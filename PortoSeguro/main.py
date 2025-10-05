import random
import sys
import time
import traceback
import glob
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSessionIdException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager


class LogManager:
    """Gerencia os logs de erros e falhas do processamento."""
    def __init__(self, pasta_base):
        self.pasta_logs = Path(pasta_base) / "logs"
        self.pasta_logs.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.arquivo_log = self.pasta_logs / f"log_falhas_{timestamp}.txt"
        self.falhas = []
        self._inicializar_log()

    def _inicializar_log(self):
        with open(self.arquivo_log, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("LOG DE FALHAS - AUTOMAÇÃO PORTO SEGURO\n")
            f.write(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")

    def registrar_falha(self, id_guia, motivo, detalhes="", traceback_erro=""):
        registro = {
            'id_guia': id_guia, 'motivo': motivo, 'detalhes': detalhes,
            'traceback': traceback_erro, 'timestamp': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        }
        if not any(f['id_guia'] == id_guia for f in self.falhas):
            self.falhas.append(registro)
            with open(self.arquivo_log, 'a', encoding='utf-8') as f:
                f.write(f"\n{'─' * 80}\n")
                f.write(f"GUIA: {id_guia}\n")
                f.write(f"Data/Hora: {registro['timestamp']}\n")
                f.write(f"Motivo: {motivo}\n")
                if detalhes: f.write(f"Detalhes: {detalhes}\n")
                if traceback_erro: f.write(f"\nTraceback:\n{traceback_erro}\n")
                f.write(f"{'─' * 80}\n")

    def gerar_resumo(self):
        with open(self.arquivo_log, 'a', encoding='utf-8') as f:
            f.write("\n\n" + "=" * 80 + "\n")
            f.write("RESUMO DE FALHAS\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"Total de falhas registradas: {len(self.falhas)}\n\n")
            if self.falhas:
                f.write("Guias que necessitam verificação manual:\n")
                for i, falha in enumerate(self.falhas, 1):
                    f.write(f"{i}. Guia {falha['id_guia']} - {falha['motivo']}\n")
        print(f"\n✓ Log de falhas salvo em: {self.arquivo_log}")
        if self.falhas:
            print(f"  → {len(self.falhas)} guia(s) necessitam verificação manual")


class PortoSeguroBot:
    def __init__(self):
        self.url = "https://www.dentistaportoseguro.com.br/"
        self.login = "00260115"
        self.senha = "@Isbi4420"
        self.driver = None
        self.wait = None
        self.log_manager = None

    def _human_delay(self, min_seconds=0.7, max_seconds=1.5):
        time.sleep(random.uniform(min_seconds, max_seconds))

    def _type_like_human(self, element, text):
        element.clear()
        for char in text:
            element.send_keys(char)
            time.sleep(random.uniform(0.05, 0.15))
    
    def handle_overlays(self, main_timeout=30):
        try:
            WebDriverWait(self.driver, 2).until(EC.visibility_of_element_located((By.ID, "spin_modal_overlay")))
            print("  - Overlay detectado. Aguardando desaparecimento...")
            WebDriverWait(self.driver, main_timeout).until(EC.invisibility_of_element_located((By.ID, "spin_modal_overlay")))
        except TimeoutException:
            pass

    def iniciar_driver(self):
        print("Iniciando navegador...")
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 20)
        print("✓ Driver iniciado com sucesso")

    def fazer_login(self):
        print("Acessando portal...")
        self.driver.get(self.url)
        campo_usuario = self.wait.until(EC.presence_of_element_located((By.ID, "usuario")))
        self._type_like_human(campo_usuario, self.login)
        campo_senha = self.driver.find_element(By.ID, "senha")
        self._type_like_human(campo_senha, self.senha)
        botao_login = self.driver.find_element(By.ID, "login-submit")
        self._human_delay()
        self.driver.execute_script("arguments[0].click();", botao_login)
        self.handle_overlays()
        print("✓ Login realizado com sucesso")

    def navegar_meus_tratamentos(self):
        try:
            menu = self.wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Meus Tratamentos")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", menu)
            self.handle_overlays()
            print("\n✓ Navegou para 'Meus Tratamentos'.")
        except Exception as e:
            print(f"✗ Erro ao navegar para 'Meus Tratamentos': {e}")
            raise

    def buscar_guia(self, id_guia):
        print(f"Buscando guia: {id_guia}")
        try:
            campo_busca = self.wait.until(EC.presence_of_element_located((By.ID, "searchGTO")))
            self._type_like_human(campo_busca, id_guia)
            self._human_delay()
            campo_busca.send_keys(Keys.ENTER)
            self.handle_overlays()
            link_guia = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[text()='{id_guia}']")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", link_guia)
            self.handle_overlays()
            print(f"✓ Guia {id_guia} aberta")
            return True
        except TimeoutException:
            print(f"✗ Guia {id_guia} não encontrada")
            return False

    def verificar_guia_ja_finalizada(self):
        try:
            WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-dollar')]")))
            return True
        except TimeoutException:
            return False

    def _extrair_data_autorizacao(self):
        try:
            elemento_p = self.wait.until(EC.presence_of_element_located((By.XPATH, "//p[strong[text()='Data da Autorização: ']]")))
            texto_completo = elemento_p.text
            match = re.search(r'\d{2}/\d{2}/\d{4}', texto_completo)
            if match:
                data = match.group(0)
                print(f"  - Data de autorização extraída: {data}")
                return data
            return None
        except TimeoutException:
            print("  - Não foi possível encontrar a data de autorização na página.")
            return None
            
    def _handle_faturamento_guia_modal(self, timeout=3):
        try:
            print("  - Verificando a presença do modal de faturamento 'GUIA'...")
            wait_curto = WebDriverWait(self.driver, timeout)
            botao_guia = wait_curto.until(EC.element_to_be_clickable((By.ID, "btnGuia")))
            print("  - Modal de faturamento detectado. Clicando em 'GUIA'...")
            self.driver.execute_script("arguments[0].click();", botao_guia)
            self.handle_overlays()
        except TimeoutException:
            print("  - Modal de faturamento não apareceu. Prosseguindo normalmente.")
            pass

    def _set_and_verify_date(self, date_element, date_to_set, checkbox):
        try:
            print(f"  - Inserindo data '{date_to_set}'...")
            textarea = self.driver.find_element(By.ID, "observacaoJustificativa")
            self.driver.execute_script("arguments[0].checked=true; arguments[1].removeAttribute('disabled'); $(arguments[0]).trigger('change');", checkbox, date_element)
            self.handle_overlays()
            date_element.send_keys(Keys.CONTROL, "a")
            date_element.send_keys(Keys.BACKSPACE)
            date_element.send_keys(date_to_set)
            self.handle_overlays()
            inserted_value = date_element.get_attribute('value')
            if inserted_value == date_to_set:
                print(f"  ✓ Verificação OK. Valor no campo: '{inserted_value}'")
                ActionChains(self.driver).move_to_element(textarea).click().perform()
                return True
            else:
                print("  ✗ FALHA CRÍTICA NA VERIFICAÇÃO DE DATA!")
                return False
        except Exception as e:
            print(f"  ✗ Erro inesperado ao tentar inserir e verificar a data: {e}")
            return False

    def confirmar_realizacao_procedimento(self):
        print("Confirmando realização do procedimento...")
        data_autorizacao = self._extrair_data_autorizacao()
        if not data_autorizacao: return False
        try:
            checkboxes = self.driver.find_elements(By.XPATH, "//input[contains(@id, 'chk_atualiza_data_')]")
            if not checkboxes: return True
            for chk in checkboxes:
                procedimento_id = chk.get_attribute('id').split('_')[-1]
                campo_data = self.driver.find_element(By.ID, f"dt_conf_{procedimento_id}")
                if not self._set_and_verify_date(campo_data, data_autorizacao, chk): return False
            botao_confirmar = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnConfirmar, #btnConfirmarData")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", botao_confirmar)
            self.handle_overlays()
            time.sleep(9)  # Pequena pausa para evitar problemas de sincronização
            self._handle_faturamento_guia_modal()
            print("✓ Procedimento confirmado. Página recarregada.")
            time.sleep(6)  # Pequena pausa para evitar problemas de sincronização

            return True
        except Exception as e:
            print(f"✗ Erro ao confirmar realização: {e}")
            return False

    def buscar_arquivos_guia(self, id_guia, pasta_base):
        print(f"Procurando arquivos para a guia {id_guia}...")
        arquivos_encontrados = []
        base_path = Path(pasta_base)
        id_parcial = id_guia[-4:]
        for f in base_path.rglob('*'):
            if not f.is_file(): continue
            nome_upper = f.stem.upper()
            if id_guia in nome_upper or id_parcial in nome_upper:
                tipo_arquivo = 'documentacao'
                if 'RX' in nome_upper or nome_upper.startswith('RF') or nome_upper.startswith('RI'):
                    tipo_arquivo = 'raio-x'
                elif 'LAUDO' in nome_upper:
                    tipo_arquivo = 'laudo'
                elif nome_upper.startswith('G') or 'GUIA' in nome_upper or 'ASS' in nome_upper:
                    tipo_arquivo = 'guia'
                arquivos_encontrados.append({'caminho': str(f.resolve()), 'tipo': tipo_arquivo})
                print(f"  → Arquivo encontrado: {f.name} (classificado como: {tipo_arquivo})")
        return arquivos_encontrados

    def _abrir_modal_anexos(self):
        time.sleep(3)
        print("  - Abrindo modal de anexos...")
        tentativas = 3
        for i in range(tentativas):
            try:
                botao_anexos = self.wait.until(EC.element_to_be_clickable((By.ID, "btnAnexosProcedimento")))
                self.driver.execute_script("arguments[0].click();", botao_anexos)
                self.handle_overlays()
                self.wait.until(EC.visibility_of_element_located((By.ID, "modalAnexosProcedimento")))
                print("  - Modal aberto com sucesso.")
                return
            except StaleElementReferenceException:
                print(f"  - AVISO: Elemento 'stale' detectado na tentativa {i + 1}/{tentativas}. Tentando novamente...")
                time.sleep(1)
        raise TimeoutException("Não foi possível abrir o modal de anexos de forma estável.")

    def _fechar_modal_anexos(self):
        print("  - Fechando modal de anexos...")
        try:
            btn_fechar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='modalAnexosProcedimento' and contains(@style,'display: block')]//button[@data-dismiss='modal']")))
            self.driver.execute_script("arguments[0].click();", btn_fechar)
            self.wait.until(EC.invisibility_of_element_located((By.ID, "modalAnexosProcedimento")))
            print("  - Modal fechado.")
        except Exception:
             print("  - Modal já estava fechado ou não foi possível fechá-lo.")

    def _anexar_um_arquivo(self, anexo_info):
        caminho_arquivo = anexo_info['caminho']
        tipo = anexo_info['tipo']
        print(f"  - Anexando: {Path(caminho_arquivo).name} como '{tipo}'")
        select_tipo = Select(self.wait.until(EC.presence_of_element_located((By.ID, "tipoAnexoModal"))))
        if tipo == 'raio-x': select_tipo.select_by_value("1")
        elif tipo == 'laudo': select_tipo.select_by_value("2")
        elif tipo == 'guia': select_tipo.select_by_value("3")
        else: select_tipo.select_by_value("4")
        self.handle_overlays()
        if tipo == 'raio-x':
            try:
                label = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='periapical']")))
                self.driver.execute_script("arguments[0].click();", label)
                print("    - Checkbox 'Periapical' marcado.")
            except TimeoutException:
                print("    - AVISO: Checkbox 'Periapical' não foi encontrado.")
        Select(self.driver.find_element(By.ID, "processoAnexo")).select_by_value("2")
        self.handle_overlays()
        self.driver.find_element(By.ID, "arquivoAnexo").send_keys(caminho_arquivo)
        self.handle_overlays()
        checkboxes = self.driver.find_elements(By.XPATH, "//table[@id='tblVincularProcedimentos']//input[@type='checkbox']")
        for checkbox in checkboxes:
            if not checkbox.is_selected():
                self.driver.execute_script("arguments[0].click();", checkbox)
        botao_incluir = self.driver.find_element(By.ID, "btnIncluirAnexos")
        self.driver.execute_script("arguments[0].click();", botao_incluir)
        self.handle_overlays()
        print("    - Arquivo incluído.")

    def anexar_todos_os_arquivos(self, lista_de_anexos):
        if not lista_de_anexos:
            print("  - Nenhum arquivo de anexo encontrado para esta guia.")
            return True
        print(f"Iniciando o anexo de {len(lista_de_anexos)} arquivo(s)...")
        try:
            for anexo in lista_de_anexos:
                self._abrir_modal_anexos()
                self._anexar_um_arquivo(anexo)
                self._fechar_modal_anexos()
            print("✓ Todos os arquivos foram anexados com sucesso.")
            return True
        except Exception as e:
            print(f"✗ Erro geral no processo de anexar arquivos: {e}")
            self.log_manager.registrar_falha("N/A", "Erro no fluxo de anexo", traceback_erro=traceback.format_exc())
            return False

    # --- FUNÇÃO QUE ESTAVA FALTANDO ---
    def submeter_para_pagamento(self):
   
        print("Submetendo para pagamento final...")
        try:
            # PASSO 1: Clicar no botão 'Confirmar' principal para abrir o modal de submissão.
            print("  - Abrindo diálogo de submissão final...")
            btn_confirmar_principal = self.wait.until(EC.element_to_be_clickable((By.ID, "btnConfirmar")))
            self.driver.execute_script("arguments[0].click();", btn_confirmar_principal)
            self.handle_overlays()
            time.sleep(7)

            # PASSO 2: DENTRO do modal, encontrar e marcar o(s) checkbox(es) com a nova lógica.
            checkboxes_finais = self.driver.find_elements(By.XPATH, "//input[starts-with(@id, 'chk_') and @type='checkbox']")
            if not checkboxes_finais:
                print("  - AVISO: Nenhum checkbox de submissão final foi encontrado no modal.")
            
            for checkbox_input in checkboxes_finais:
                # Mantemos a verificação para não desmarcar um box já marcado
                if not checkbox_input.is_selected():
                    checkbox_id = checkbox_input.get_attribute('id')
                    print(f"  - Marcando o checkbox de submissão final no modal (ID: {checkbox_id})...")
                    
                    # Encontra a label correspondente
                    label_para_clicar = self.wait.until(EC.presence_of_element_located((By.XPATH, f"//label[@for='{checkbox_id}']")))
                    
                    # Lógica de rolagem e clique que você sugeriu
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", label_para_clicar)
                    self._human_delay(0.5, 1.0) # Pequena pausa após rolar
                    self.driver.execute_script("arguments[0].click();", label_para_clicar)
                    self.handle_overlays()

            botao_confirmar = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnConfirmar, #btnConfirmarData")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", botao_confirmar)

            # Verificação final de sucesso na página principal
            self.wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-dollar')]")))
            print("✓ Guia submetida e ícone de pagamento ($) confirmado.")
            return True

        except Exception as e:
            print(f"✗ Erro na submissão final para pagamento: {e}")
            return False


def processar_guias_da_planilha(bot, caminho_planilha, pasta_arquivos_busca):
    try:
        workbook = openpyxl.load_workbook(caminho_planilha)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"✗ Erro: Planilha não encontrada em '{caminho_planilha}'")
        return

    guias_a_processar = [str(cell.value).strip() for cell in sheet['A'] if cell.value is not None]
    total_guias = len(guias_a_processar)
    print(f"\nEncontradas {total_guias} guias para processar.")
    
    sucessos = 0
    
    try:
        bot.fazer_login()
    except Exception:
        bot.log_manager.registrar_falha("GERAL", "Falha no Login", traceback_erro=traceback.format_exc())
        return

    for i, id_guia in enumerate(guias_a_processar, 1):
        print("\n" + "="*50)
        print(f"Processando Guia {i}/{total_guias}: {id_guia}")
        print("="*50)
        
        try:
            bot.navegar_meus_tratamentos()
            if not bot.buscar_guia(id_guia):
                bot.log_manager.registrar_falha(id_guia, "Guia não encontrada na busca")
                continue
            
            if bot.verificar_guia_ja_finalizada():
                print(f"✓ Guia {id_guia} já está finalizada. Pulando.")
                sucessos += 1
                continue

            if not bot.confirmar_realizacao_procedimento():
                bot.log_manager.registrar_falha(id_guia, "Falha ao confirmar realização")
                continue

            lista_de_anexos = bot.buscar_arquivos_guia(id_guia, pasta_arquivos_busca)
            if not bot.anexar_todos_os_arquivos(lista_de_anexos):
                bot.log_manager.registrar_falha(id_guia, "Falha no processo de anexo")
                continue

            if not bot.submeter_para_pagamento():
                bot.log_manager.registrar_falha(id_guia, "Falha na submissão final para pagamento")
                continue
            
            print(f"✓✓ Guia {id_guia} finalizada com SUCESSO! ✓✓")
            sucessos += 1

        except Exception as e:
            print(f"✗ Ocorreu um erro inesperado no processamento da guia {id_guia}: {e}")
            bot.log_manager.registrar_falha(id_guia, "Erro inesperado no fluxo", traceback_erro=traceback.format_exc())
        finally:
            try:
                print("Finalizando ciclo da guia, retornando para 'Meus Tratamentos'...")
                bot.navegar_meus_tratamentos()
            except Exception:
                print("✗ Falha crítica ao tentar voltar para 'Meus Tratamentos'. Abortando.")
                break

    print("\n" + "="*80)
    print("PROCESSAMENTO FINALIZADO")
    print(f"  - Total de guias na planilha: {total_guias}")
    print(f"  - Guias processadas com sucesso: {sucessos}")
    print(f"  - Guias com falha: {len(bot.log_manager.falhas)}")
    print("="*80)


if __name__ == "__main__":
    pasta_porto_seguro = Path(__file__).parent
    pasta_base = pasta_porto_seguro.parent
    pasta_arquivos = pasta_porto_seguro / "arquivos"
    print(f"Buscando por planilha (.xlsx) na pasta: {pasta_porto_seguro}")
    lista_planilhas = glob.glob(str(pasta_porto_seguro / "*.xlsx"))
    caminho_planilha = None
    if len(lista_planilhas) == 0:
        print("\n!!! ERRO: Nenhuma planilha (.xlsx) foi encontrada."); sys.exit()
    elif len(lista_planilhas) > 1:
        print("\n!!! ERRO: Mais de uma planilha (.xlsx) foi encontrada."); sys.exit()
    else:
        caminho_planilha = lista_planilhas[0]
        print(f"✓ Planilha encontrada: {Path(caminho_planilha).name}\n")
    
    log_manager = LogManager(pasta_base)
    bot = PortoSeguroBot()
    bot.log_manager = log_manager

    driver_iniciado = False
    try:
        bot.iniciar_driver()
        driver_iniciado = True
        processar_guias_da_planilha(bot, caminho_planilha, str(pasta_arquivos))
    except Exception as e:
        print(f"\n✗ Ocorreu um erro fatal na execução do robô: {e}")
        log_manager.registrar_falha("FATAL", "Erro crítico", traceback_erro=traceback.format_exc())
    finally:
        log_manager.gerar_resumo()
        if driver_iniciado and bot.driver:
            print("Fechando navegador...")
            bot.driver.quit()
        print("Fim da execução.")