import random
import sys
import time
import traceback
import glob
import re
import os
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


# ======================== CORES PARA O CONSOLE ========================
class Cores:
    """Códigos ANSI para colorir o terminal (compatível com Windows 10+)"""
    RESET = '\033[0m'
    BOLD = '\033[1m'
    
    # Cores principais
    VERDE = '\033[92m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    AZUL = '\033[94m'
    CYAN = '\033[96m'
    MAGENTA = '\033[95m'
    BRANCO = '\033[97m'
    
    # Backgrounds
    BG_VERDE = '\033[102m\033[30m'
    BG_VERMELHO = '\033[101m\033[97m'
    BG_AZUL = '\033[104m\033[97m'

# Habilita cores ANSI no Windows
if os.name == 'nt':
    os.system('')


# ======================== FUNÇÕES DE INTERFACE ========================
def limpar_tela():
    """Limpa a tela do console."""
    os.system('cls' if os.name == 'nt' else 'clear')

def exibir_banner():
    """Exibe o banner principal do sistema."""
    limpar_tela()
    print(Cores.CYAN + Cores.BOLD)
    print("╔" + "═" * 78 + "╗")
    print("║" + " " * 78 + "║")
    print("║" + "        🦷  SISTEMA DE AUTOMAÇÃO PORTO SEGURO  🦷       ".center(76) + "║")
    print("║" + " " * 78 + "║")
    print("║" + f"        Versão 1.0 | Desenvolvido em {datetime.now().year}        ".center(78) + "║")
    print("║" + " " * 78 + "║")
    print("╚" + "═" * 78 + "╝")
    print(Cores.RESET)

def exibir_separador(char="─", tamanho=80, cor=Cores.AZUL):
    """Exibe uma linha separadora."""
    print(cor + char * tamanho + Cores.RESET)

def exibir_mensagem_sucesso(mensagem):
    """Exibe mensagem de sucesso formatada."""
    print(f"{Cores.VERDE}✓ {mensagem}{Cores.RESET}")

def exibir_mensagem_erro(mensagem):
    """Exibe mensagem de erro formatada."""
    print(f"{Cores.VERMELHO}✗ {mensagem}{Cores.RESET}")

def exibir_mensagem_info(mensagem):
    """Exibe mensagem informativa formatada."""
    print(f"{Cores.CYAN}ℹ {mensagem}{Cores.RESET}")

def exibir_mensagem_alerta(mensagem):
    """Exibe mensagem de alerta formatada."""
    print(f"{Cores.AMARELO}⚠ {mensagem}{Cores.RESET}")


# ======================== CLASSES ORIGINAIS (LogManager e PortoSeguroBot) ========================
class LogManager:
    """Gerencia o sistema de logs de erros e falhas durante o processamento das guias."""
    
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
        exibir_mensagem_info(f"Log de falhas salvo em: {self.arquivo_log}")
        if self.falhas:
            exibir_mensagem_alerta(f"{len(self.falhas)} guia(s) necessitam verificação manual")


class PortoSeguroBot:
    """Bot de automação para processamento de guias no portal Porto Seguro."""
    
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
            exibir_mensagem_info("Overlay detectado. Aguardando...")
            WebDriverWait(self.driver, main_timeout).until(EC.invisibility_of_element_located((By.ID, "spin_modal_overlay")))
        except TimeoutException:
            pass

    def iniciar_driver(self):
        exibir_mensagem_info("Iniciando navegador Chrome...")
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 20)
        exibir_mensagem_sucesso("Driver iniciado com sucesso")

    def fazer_login(self):
        exibir_mensagem_info("Acessando portal Porto Seguro...")
        self.driver.get(self.url)
        campo_usuario = self.wait.until(EC.presence_of_element_located((By.ID, "usuario")))
        self._type_like_human(campo_usuario, self.login)
        campo_senha = self.driver.find_element(By.ID, "senha")
        self._type_like_human(campo_senha, self.senha)
        botao_login = self.driver.find_element(By.ID, "login-submit")
        self._human_delay()
        self.driver.execute_script("arguments[0].click();", botao_login)
        self.handle_overlays()
        exibir_mensagem_sucesso("Login realizado com sucesso")

    def navegar_meus_tratamentos(self):
        try:
            menu = self.wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Meus Tratamentos")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", menu)
            self.handle_overlays()
            exibir_mensagem_sucesso("Navegou para 'Meus Tratamentos'")
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao navegar: {e}")
            raise

    def buscar_guia(self, id_guia):
        print(f"\n{Cores.BOLD}Processando guia: {id_guia}{Cores.RESET}")
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
            exibir_mensagem_sucesso(f"Guia {id_guia} aberta")
            return True
        except TimeoutException:
            exibir_mensagem_erro(f"Guia {id_guia} não encontrada")
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
                exibir_mensagem_info(f"Data de autorização: {data}")
                return data
            return None
        except TimeoutException:
            exibir_mensagem_alerta("Data de autorização não encontrada")
            return None
            
    def _handle_faturamento_guia_modal(self, timeout=3):
        try:
            wait_curto = WebDriverWait(self.driver, timeout)
            botao_guia = wait_curto.until(EC.element_to_be_clickable((By.ID, "btnGuia")))
            exibir_mensagem_info("Modal de faturamento detectado")
            self.driver.execute_script("arguments[0].click();", botao_guia)
            self.handle_overlays()
        except TimeoutException:
            pass

    def _set_and_verify_date(self, date_element, date_to_set, checkbox):
        try:
            textarea = self.driver.find_element(By.ID, "observacaoJustificativa")
            self.driver.execute_script("arguments[0].checked=true; arguments[1].removeAttribute('disabled'); $(arguments[0]).trigger('change');", checkbox, date_element)
            self.handle_overlays()
            date_element.send_keys(Keys.CONTROL, "a")
            date_element.send_keys(Keys.BACKSPACE)
            date_element.send_keys(date_to_set)
            self.handle_overlays()
            inserted_value = date_element.get_attribute('value')
            if inserted_value == date_to_set:
                exibir_mensagem_sucesso(f"Data verificada: {inserted_value}")
                ActionChains(self.driver).move_to_element(textarea).click().perform()
                return True
            else:
                exibir_mensagem_erro("Falha na verificação de data")
                return False
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao inserir data: {e}")
            return False

    def confirmar_realizacao_procedimento(self):
        exibir_mensagem_info("Confirmando realização do procedimento...")
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
            time.sleep(9)
            self._handle_faturamento_guia_modal()
            exibir_mensagem_sucesso("Procedimento confirmado")
            time.sleep(6)
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao confirmar: {e}")
            return False

    def buscar_arquivos_guia(self, id_guia, pasta_base):
        exibir_mensagem_info(f"Buscando arquivos para guia {id_guia}...")
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
                print(f"  {Cores.VERDE}→{Cores.RESET} {f.name} ({tipo_arquivo})")
        return arquivos_encontrados

    def _abrir_modal_anexos(self):
        time.sleep(3)
        exibir_mensagem_info("Abrindo modal de anexos...")
        tentativas = 3
        for i in range(tentativas):
            try:
                botao_anexos = self.wait.until(EC.element_to_be_clickable((By.ID, "btnAnexosProcedimento")))
                self.driver.execute_script("arguments[0].click();", botao_anexos)
                self.handle_overlays()
                self.wait.until(EC.visibility_of_element_located((By.ID, "modalAnexosProcedimento")))
                return
            except StaleElementReferenceException:
                exibir_mensagem_alerta(f"Tentativa {i + 1}/{tentativas}...")
                time.sleep(1)
        raise TimeoutException("Não foi possível abrir o modal de anexos")

    def _fechar_modal_anexos(self):
        try:
            btn_fechar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='modalAnexosProcedimento' and contains(@style,'display: block')]//button[@data-dismiss='modal']")))
            self.driver.execute_script("arguments[0].click();", btn_fechar)
            self.wait.until(EC.invisibility_of_element_located((By.ID, "modalAnexosProcedimento")))
        except Exception:
            pass

    def _anexar_um_arquivo(self, anexo_info):
        caminho_arquivo = anexo_info['caminho']
        tipo = anexo_info['tipo']
        print(f"  {Cores.CYAN}→{Cores.RESET} Anexando: {Path(caminho_arquivo).name}")
        
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
            except TimeoutException:
                pass
        
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

    def anexar_todos_os_arquivos(self, lista_de_anexos):
        if not lista_de_anexos:
            exibir_mensagem_alerta("Nenhum arquivo encontrado")
            return True
        exibir_mensagem_info(f"Anexando {len(lista_de_anexos)} arquivo(s)...")
        try:
            for anexo in lista_de_anexos:
                self._abrir_modal_anexos()
                self._anexar_um_arquivo(anexo)
                self._fechar_modal_anexos()
            exibir_mensagem_sucesso("Todos os arquivos anexados")
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro no anexo: {e}")
            self.log_manager.registrar_falha("N/A", "Erro no anexo", traceback_erro=traceback.format_exc())
            return False

    def submeter_para_pagamento(self):
        exibir_mensagem_info("Submetendo para pagamento...")
        try:
            btn_confirmar_principal = self.wait.until(EC.element_to_be_clickable((By.ID, "btnConfirmar")))
            self.driver.execute_script("arguments[0].click();", btn_confirmar_principal)
            self.handle_overlays()
            time.sleep(7)

            checkboxes_finais = self.driver.find_elements(By.XPATH, "//input[starts-with(@id, 'chk_') and @type='checkbox']")
            for checkbox_input in checkboxes_finais:
                if not checkbox_input.is_selected():
                    checkbox_id = checkbox_input.get_attribute('id')
                    label_para_clicar = self.wait.until(EC.presence_of_element_located((By.XPATH, f"//label[@for='{checkbox_id}']")))
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", label_para_clicar)
                    self._human_delay(0.5, 1.0)
                    self.driver.execute_script("arguments[0].click();", label_para_clicar)
                    self.handle_overlays()

            botao_confirmar = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnConfirmar, #btnConfirmarData")))
            self._human_delay()
            self.driver.execute_script("arguments[0].click();", botao_confirmar)

            self.wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-dollar')]")))
            exibir_mensagem_sucesso("Guia submetida e pagamento confirmado ($)")
            return True

        except Exception as e:
            exibir_mensagem_erro(f"Erro na submissão: {e}")
            return False


# ======================== PROCESSAMENTO DE GUIAS ========================
def processar_guias_da_planilha(bot, caminho_planilha, pasta_arquivos_busca):
    """Função principal que processa todas as guias da planilha."""
    try:
        workbook = openpyxl.load_workbook(caminho_planilha)
        sheet = workbook.active
    except FileNotFoundError:
        exibir_mensagem_erro(f"Planilha não encontrada: {caminho_planilha}")
        return

    guias_a_processar = [str(cell.value).strip() for cell in sheet['A'] if cell.value is not None]
    total_guias = len(guias_a_processar)
    
    print(f"\n{Cores.BOLD}{Cores.AZUL}{'═' * 80}{Cores.RESET}")
    print(f"{Cores.BOLD}INICIANDO PROCESSAMENTO DE {total_guias} GUIAS{Cores.RESET}")
    print(f"{Cores.AZUL}{'═' * 80}{Cores.RESET}\n")
    
    sucessos = 0
    
    try:
        bot.fazer_login()
    except Exception:
        bot.log_manager.registrar_falha("GERAL", "Falha no Login", traceback_erro=traceback.format_exc())
        return

    for i, id_guia in enumerate(guias_a_processar, 1):
        print(f"\n{Cores.BOLD}{Cores.MAGENTA}{'═' * 80}")
        print(f"║  GUIA {i}/{total_guias}: {id_guia}")
        print(f"{'═' * 80}{Cores.RESET}\n")
        
        try:
            bot.navegar_meus_tratamentos()
            
            if not bot.buscar_guia(id_guia):
                bot.log_manager.registrar_falha(id_guia, "Guia não encontrada")
                continue
            
            if bot.verificar_guia_ja_finalizada():
                exibir_mensagem_sucesso(f"Guia {id_guia} já finalizada. Pulando.")
                sucessos += 1
                continue

            if not bot.confirmar_realizacao_procedimento():
                bot.log_manager.registrar_falha(id_guia, "Falha ao confirmar")
                continue

            lista_de_anexos = bot.buscar_arquivos_guia(id_guia, pasta_arquivos_busca)
            if not bot.anexar_todos_os_arquivos(lista_de_anexos):
                bot.log_manager.registrar_falha(id_guia, "Falha no anexo")
                continue

            if not bot.submeter_para_pagamento():
                bot.log_manager.registrar_falha(id_guia, "Falha na submissão")
                continue
            
            print(f"\n{Cores.BG_VERDE} ✓✓ GUIA {id_guia} FINALIZADA COM SUCESSO ✓✓ {Cores.RESET}\n")
            sucessos += 1

        except KeyboardInterrupt:
            exibir_mensagem_alerta("\n\nProcessamento interrompido pelo usuário!")
            break
        except Exception as e:
            exibir_mensagem_erro(f"Erro inesperado: {e}")
            bot.log_manager.registrar_falha(id_guia, "Erro inesperado", traceback_erro=traceback.format_exc())
        finally:
            try:
                bot.navegar_meus_tratamentos()
            except Exception:
                exibir_mensagem_erro("Falha ao retornar. Abortando.")
                break

    # Relatório final
    print(f"\n{Cores.BOLD}{Cores.CYAN}{'═' * 80}")
    print("║  PROCESSAMENTO FINALIZADO")
    print(f"{'═' * 80}{Cores.RESET}")
    print(f"{Cores.BOLD}Total de guias:{Cores.RESET} {total_guias}")
    print(f"{Cores.VERDE}Processadas com sucesso:{Cores.RESET} {sucessos}")
    print(f"{Cores.VERMELHO}Guias com falha:{Cores.RESET} {len(bot.log_manager.falhas)}")
    print(f"{Cores.CYAN}{'═' * 80}{Cores.RESET}\n")


# ======================== MENU INTERATIVO PROFISSIONAL ========================
def listar_logs_disponiveis(pasta_logs):
    """Lista todos os logs de falha disponíveis."""
    logs = sorted(pasta_logs.glob("log_falhas_*.txt"), reverse=True)
    return logs

def exibir_log_especifico(caminho_log):
    """Exibe o conteúdo de um log específico."""
    limpar_tela()
    exibir_banner()
    print(f"\n{Cores.BOLD}📄 LOG: {caminho_log.name}{Cores.RESET}\n")
    exibir_separador()
    try:
        with open(caminho_log, 'r', encoding='utf-8') as f:
            conteudo = f.read()
            print(conteudo)
    except FileNotFoundError:
        exibir_mensagem_erro("Arquivo de log não encontrado")
    exibir_separador()
    input(f"\n{Cores.AMARELO}Pressione ENTER para voltar ao menu...{Cores.RESET}")

def exibir_status_sistema(caminho_planilha, pasta_arquivos):
    """Exibe o status atual do sistema."""
    print(f"\n{Cores.BOLD}{Cores.AZUL}╔{'═' * 78}╗")
    print(f"║{'STATUS DO SISTEMA'.center(78)}║")
    print(f"╚{'═' * 78}╝{Cores.RESET}\n")
    
    # Verifica planilha
    if caminho_planilha and Path(caminho_planilha).exists():
        try:
            wb = openpyxl.load_workbook(caminho_planilha)
            guias = [c.value for c in wb.active['A'] if c.value]
            exibir_mensagem_sucesso(f"Planilha encontrada: {Path(caminho_planilha).name}")
            print(f"  {Cores.CYAN}→{Cores.RESET} Total de guias: {Cores.BOLD}{len(guias)}{Cores.RESET}")
        except:
            exibir_mensagem_erro("Erro ao ler planilha")
    else:
        exibir_mensagem_erro("Planilha não encontrada")
    
    # Verifica pasta de arquivos
    if pasta_arquivos and Path(pasta_arquivos).exists():
        total_arquivos = len(list(Path(pasta_arquivos).rglob('*.*')))
        exibir_mensagem_sucesso(f"Pasta de arquivos encontrada: {Path(pasta_arquivos).name}")
        print(f"  {Cores.CYAN}→{Cores.RESET} Total de arquivos: {Cores.BOLD}{total_arquivos}{Cores.RESET}")
    else:
        exibir_mensagem_erro("Pasta de arquivos não encontrada")
    
    print()

def exibir_menu_logs(pasta_logs):
    """Exibe submenu para visualizar logs."""
    while True:
        limpar_tela()
        exibir_banner()
        print(f"\n{Cores.BOLD}{Cores.AZUL}╔{'═' * 78}╗")
        print(f"║{'LOGS DE EXECUÇÃO'.center(78)}║")
        print(f"╚{'═' * 78}╝{Cores.RESET}\n")
        
        logs = listar_logs_disponiveis(pasta_logs)
        
        if not logs:
            exibir_mensagem_alerta("Nenhum log encontrado ainda")
            input(f"\n{Cores.AMARELO}Pressione ENTER para voltar...{Cores.RESET}")
            return
        
        print(f"{Cores.CYAN}Logs disponíveis:{Cores.RESET}\n")
        for i, log in enumerate(logs, 1):
            timestamp_str = log.stem.replace("log_falhas_", "")
            try:
                dt = datetime.strptime(timestamp_str, "%Y%m%d_%H%M%S")
                data_formatada = dt.strftime("%d/%m/%Y às %H:%M:%S")
            except:
                data_formatada = timestamp_str
            print(f"  {Cores.BOLD}[{i}]{Cores.RESET} {data_formatada}")
        
        print(f"\n  {Cores.BOLD}[0]{Cores.RESET} Voltar ao menu principal")
        
        exibir_separador()
        escolha = input(f"\n{Cores.AMARELO}Selecione o log para visualizar: {Cores.RESET}").strip()
        
        if escolha == '0':
            return
        
        try:
            idx = int(escolha) - 1
            if 0 <= idx < len(logs):
                exibir_log_especifico(logs[idx])
            else:
                exibir_mensagem_erro("Opção inválida!")
                time.sleep(1)
        except ValueError:
            exibir_mensagem_erro("Digite um número válido!")
            time.sleep(1)

def confirmar_inicio_processamento():
    """Solicita confirmação do usuário antes de iniciar."""
    print(f"\n{Cores.AMARELO}{Cores.BOLD}⚠  ATENÇÃO ⚠{Cores.RESET}")
    print(f"\n{Cores.AMARELO}O processamento será iniciado.{Cores.RESET}")
    print(f"{Cores.AMARELO}Você pode interromper a qualquer momento pressionando Ctrl+C{Cores.RESET}\n")
    
    while True:
        resposta = input(f"{Cores.VERDE}Deseja continuar? (S/N): {Cores.RESET}").strip().upper()
        if resposta in ['S', 'SIM', 'Y', 'YES']:
            return True
        elif resposta in ['N', 'NAO', 'NÃO', 'NO']:
            return False
        else:
            exibir_mensagem_erro("Resposta inválida! Digite S ou N")

def exibir_menu_principal_e_processar(bot, caminho_planilha, pasta_arquivos):
    """Menu principal interativo do sistema."""
    
    while True:
        exibir_banner()
        exibir_status_sistema(caminho_planilha, pasta_arquivos)
        
        print(f"{Cores.BOLD}{Cores.AZUL}╔{'═' * 78}╗")
        print(f"║{'MENU PRINCIPAL'.center(78)}║")
        print(f"╚{'═' * 78}╝{Cores.RESET}\n")
        
        print(f"  {Cores.BOLD}[1]{Cores.RESET} {Cores.VERDE}▶  Iniciar Processamento de Guias{Cores.RESET}")
        print(f"  {Cores.BOLD}[2]{Cores.RESET} {Cores.CYAN}📄 Visualizar Logs de Execução{Cores.RESET}")
        print(f"  {Cores.BOLD}[3]{Cores.RESET} {Cores.MAGENTA}📊 Verificar Status do Sistema{Cores.RESET}")
        print(f"  {Cores.BOLD}[4]{Cores.RESET} {Cores.AMARELO}ℹ  Sobre o Sistema{Cores.RESET}")
        print(f"  {Cores.BOLD}[0]{Cores.RESET} {Cores.VERMELHO}✖  Sair{Cores.RESET}")
        
        exibir_separador()
        escolha = input(f"\n{Cores.AMARELO}Digite sua opção: {Cores.RESET}").strip()
        
        if escolha == '1':
            # INICIAR PROCESSAMENTO
            limpar_tela()
            exibir_banner()
            
            if not confirmar_inicio_processamento():
                exibir_mensagem_info("Processamento cancelado pelo usuário")
                time.sleep(2)
                continue
            
            print(f"\n{Cores.VERDE}{Cores.BOLD}{'═' * 80}")
            print("║  INICIANDO SISTEMA DE AUTOMAÇÃO")
            print(f"{'═' * 80}{Cores.RESET}\n")
            
            driver_iniciado = False
            try:
                bot.iniciar_driver()
                driver_iniciado = True
                processar_guias_da_planilha(bot, caminho_planilha, pasta_arquivos)
            except KeyboardInterrupt:
                exibir_mensagem_alerta("\n\nProcessamento interrompido pelo usuário!")
            except Exception as e:
                exibir_mensagem_erro(f"\nErro fatal: {e}")
                bot.log_manager.registrar_falha("FATAL", "Erro crítico", traceback_erro=traceback.format_exc())
            finally:
                if driver_iniciado and bot.driver:
                    exibir_mensagem_info("Fechando navegador...")
                    bot.driver.quit()
            
            input(f"\n{Cores.AMARELO}Pressione ENTER para voltar ao menu...{Cores.RESET}")
        
        elif escolha == '2':
            # VISUALIZAR LOGS
            exibir_menu_logs(bot.log_manager.pasta_logs)
        
        elif escolha == '3':
            # STATUS DO SISTEMA
            limpar_tela()
            exibir_banner()
            exibir_status_sistema(caminho_planilha, pasta_arquivos)
            
            # Informações adicionais
            print(f"\n{Cores.BOLD}Informações de Configuração:{Cores.RESET}\n")
            print(f"  {Cores.CYAN}→{Cores.RESET} URL Portal: {bot.url}")
            print(f"  {Cores.CYAN}→{Cores.RESET} Usuário: {bot.login}")
            print(f"  {Cores.CYAN}→{Cores.RESET} Pasta de Logs: {bot.log_manager.pasta_logs}")
            
            input(f"\n{Cores.AMARELO}Pressione ENTER para voltar...{Cores.RESET}")
        
        elif escolha == '4':
            # SOBRE O SISTEMA
            limpar_tela()
            exibir_banner()
            print(f"\n{Cores.BOLD}{Cores.AZUL}╔{'═' * 78}╗")
            print(f"║{'SOBRE O SISTEMA'.center(78)}║")
            print(f"╚{'═' * 78}╝{Cores.RESET}\n")
            
            print(f"{Cores.BOLD}Sistema de Automação Porto Seguro{Cores.RESET}\n")
            print("Este sistema automatiza o processamento de guias de procedimentos odontológicos")
            print("no portal da Porto Seguro, realizando as seguintes tarefas:\n")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Login automático no portal")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Busca e validação de guias")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Confirmação de procedimentos realizados")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Anexação automática de documentos (RX, laudos, guias)")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Submissão para pagamento")
            print(f"  {Cores.VERDE}✓{Cores.RESET} Registro detalhado de logs e falhas")
            
            print(f"\n{Cores.BOLD}Recursos:{Cores.RESET}\n")
            print(f"  {Cores.CYAN}→{Cores.RESET} Processamento em lote via planilha Excel")
            print(f"  {Cores.CYAN}→{Cores.RESET} Detecção inteligente de arquivos por ID de guia")
            print(f"  {Cores.CYAN}→{Cores.RESET} Classificação automática de tipos de anexo")
            print(f"  {Cores.CYAN}→{Cores.RESET} Sistema de logs com timestamp")
            print(f"  {Cores.CYAN}→{Cores.RESET} Verificação de guias já finalizadas")
            print(f"  {Cores.CYAN}→{Cores.RESET} Simulação de comportamento humano")
            
            print(f"\n{Cores.BOLD}Tecnologias:{Cores.RESET}")
            print(f"  {Cores.MAGENTA}→{Cores.RESET} Python 3.x")
            print(f"  {Cores.MAGENTA}→{Cores.RESET} Selenium WebDriver")
            print(f"  {Cores.MAGENTA}→{Cores.RESET} OpenPyXL")
            
            input(f"\n{Cores.AMARELO}Pressione ENTER para voltar...{Cores.RESET}")
        
        elif escolha == '0':
            # SAIR
            limpar_tela()
            exibir_banner()
            print(f"\n{Cores.VERDE}{Cores.BOLD}Obrigado por usar o Sistema de Automação Porto Seguro!{Cores.RESET}\n")
            print(f"{Cores.CYAN}Sistema encerrado com sucesso.{Cores.RESET}\n")
            break
        
        else:
            exibir_mensagem_erro("Opção inválida! Digite um número entre 0 e 4.")
            time.sleep(1.5)


# ======================== PONTO DE ENTRADA PRINCIPAL ========================
if __name__ == "__main__":
    """
    Ponto de entrada principal do script.
    Configura os caminhos, localiza a planilha, inicializa o bot e executa o processamento.
    """
    # Configuração de caminhos
    pasta_porto_seguro = Path(__file__).parent
    pasta_base = pasta_porto_seguro.parent
    pasta_arquivos = pasta_porto_seguro / "arquivos"
    
    # Validação da estrutura de pastas
    if not pasta_arquivos.exists():
        print(f"{Cores.AMARELO}⚠ Pasta 'arquivos' não encontrada. Criando...{Cores.RESET}")
        pasta_arquivos.mkdir(exist_ok=True)
    
    # Busca por planilha Excel na pasta do script
    print(f"Buscando por planilha (.xlsx) na pasta: {pasta_porto_seguro}")
    lista_planilhas = glob.glob(str(pasta_porto_seguro / "*.xlsx"))
    caminho_planilha = None
    
    # Validação: deve existir exatamente uma planilha
    if len(lista_planilhas) == 0:
        print("\n!!! ERRO: Nenhuma planilha (.xlsx) foi encontrada.")
        sys.exit()
    elif len(lista_planilhas) > 1:
        print("\n!!! ERRO: Mais de uma planilha (.xlsx) foi encontrada.")
        sys.exit()
    else:
        caminho_planilha = lista_planilhas[0]
        print(f"✓ Planilha encontrada: {Path(caminho_planilha).name}\n")
    
    # Inicializa o sistema de logs
    log_manager = LogManager(pasta_base)
    bot = PortoSeguroBot()
    bot.log_manager = log_manager

    driver_iniciado = False
    try:
        # Ponto de entrada: Menu interativo
        exibir_menu_principal_e_processar(bot, caminho_planilha, str(pasta_arquivos))
        
    except KeyboardInterrupt:
        print(f"\n\n{Cores.AMARELO}Sistema interrompido pelo usuário.{Cores.RESET}")
    except Exception as e:
        print(f"\n✗ Ocorreu um erro fatal na execução do robô: {e}")
        log_manager.registrar_falha("FATAL", "Erro crítico", traceback_erro=traceback.format_exc())
    finally:
        # Gera o resumo final de logs
        log_manager.gerar_resumo()
        
        # Fecha o navegador se foi iniciado
        if driver_iniciado and bot.driver:
            print("Fechando navegador...")
            bot.driver.quit()
        
        print("Fim da execução.")