import csv
import os
import random
import re
import time
import traceback
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (NoSuchElementException,
                                        TimeoutException)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from constants import MAPEAMENTO_PROCEDIMENTOS, URL_PORTAL
from utils import (exibir_mensagem_alerta, exibir_mensagem_erro,
                   exibir_mensagem_info, exibir_mensagem_sucesso)


class LogManager:
    """Gerencia a criação de logs de erros e relatórios de sucesso."""
    def __init__(self, pasta_base):
        self.pasta_logs = Path(pasta_base) / "logs"
        self.pasta_logs.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        self.arquivo_log_falhas = self.pasta_logs / f"log_falhas_{timestamp}.txt"
        self.falhas = []
        self._inicializar_log_falhas()
        
        self.arquivo_relatorio_sucesso = self.pasta_logs / f"relatorio_sucesso_{timestamp}.csv"
        self._inicializar_relatorio_sucesso()

    def _inicializar_log_falhas(self):
        with open(self.arquivo_log_falhas, 'w', encoding='utf-8') as f:
            f.write(f"LOG DE FALHAS - AUTOMAÇÃO PORTO SEGURO\nData/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n{'='*80}\n\n")

    def _inicializar_relatorio_sucesso(self):
        with open(self.arquivo_relatorio_sucesso, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Guia', 'Codigo_Planilha', 'Dente_Selecionado', 'Status', 'Data_Hora'])
            
    def registrar_sucesso(self, id_guia, cod_simplificado, dente_val):
        with open(self.arquivo_relatorio_sucesso, 'a', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([id_guia, cod_simplificado, dente_val, 'Sucesso', datetime.now().strftime('%d/%m/%Y %H:%M:%S')])

    def registrar_falha(self, id_guia, motivo, traceback_erro=""):
        registro = {'id_guia': id_guia, 'motivo': motivo}
        if not any(f['id_guia'] == id_guia and f['motivo'] == motivo for f in self.falhas):
            self.falhas.append(registro)
            with open(self.arquivo_log_falhas, 'a', encoding='utf-8') as f:
                f.write(f"\n{'─' * 80}\nGUIA: {id_guia}\nData/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\nMotivo: {motivo}\n")
                if traceback_erro: f.write(f"\nTraceback:\n{traceback_erro}\n")
                f.write(f"{'─' * 80}\n")

    def gerar_resumos(self):
        exibir_mensagem_info(f"Relatório de sucessos salvo em: {self.arquivo_relatorio_sucesso}")
        if self.falhas:
            exibir_mensagem_info(f"Log de falhas salvo em: {self.arquivo_log_falhas}")
            exibir_mensagem_alerta(f"{len(self.falhas)} guia(s) necessitam de verificação manual.")

class PortoSeguroBot:
    """Encapsula toda a lógica de interação com o portal Porto Seguro Odonto."""
    def __init__(self, tipo_usuario="completo"):
        self.url = URL_PORTAL
        
        # Carrega credenciais das variáveis de ambiente
        login_completo = os.getenv("LOGIN_COMPLETO")
        senha_completo = os.getenv("SENHA_COMPLETO")
        login_simplificado = os.getenv("LOGIN_SIMPLIFICADO")
        senha_simplificado = os.getenv("SENHA_SIMPLIFICADO")

        # Configuração de usuários
        self.usuarios = {
            "completo": {
                "login": login_completo,
                "senha": senha_completo,
                "nome": "Usuário Completo (com Procedimento Complementar)",
                "usa_modal_complementar": True
            },
            "simplificado": {
                "login": login_simplificado,
                "senha": senha_simplificado,
                "nome": "Usuário Simplificado (sem Procedimento Complementar)",
                "usa_modal_complementar": False
            }
        }
        
        # Define o tipo de usuário atual
        self.tipo_usuario = tipo_usuario
        self.login = self.usuarios[tipo_usuario]["login"]
        self.senha = self.usuarios[tipo_usuario]["senha"]
        self.usa_modal_complementar = self.usuarios[tipo_usuario]["usa_modal_complementar"]
        
        self.driver = None
        self.wait = None
        self.log_manager = None

    def _human_delay(self, min_seconds=0.5, max_seconds=1.0):
        """Pausa a execução por um tempo aleatório para simular comportamento humano."""
        time.sleep(random.uniform(min_seconds, max_seconds))
    
    def _pause_between_actions(self, seconds=6):
        """Pausa estratégica entre ações importantes do fluxo."""
        exibir_mensagem_info(f"Aguardando {seconds} segundos antes da próxima ação...")
        time.sleep(seconds)

    def handle_overlays(self, main_timeout=30):
        """Aguarda o desaparecimento de pop-ups de carregamento (spinners)."""
        try:
            WebDriverWait(self.driver, 2).until(EC.visibility_of_element_located((By.ID, "spin_modal_overlay")))
            exibir_mensagem_info("Overlay detectado. Aguardando...")
            WebDriverWait(self.driver, main_timeout).until(EC.invisibility_of_element_located((By.ID, "spin_modal_overlay")))
        except TimeoutException:
            pass

    def iniciar_driver(self):
        """Configura e inicia o WebDriver do Selenium."""
        exibir_mensagem_info("Iniciando navegador Chrome...")
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 20)
        exibir_mensagem_sucesso("Driver iniciado com sucesso")

    def fazer_login(self):
        """Executa o processo de login no portal."""
        exibir_mensagem_info("Acessando portal Porto Seguro...")
        self.driver.get(self.url)
        self._human_delay(1, 2)
        
        if not self.login or not self.senha:
            raise ValueError("Credenciais não encontradas. Verifique o arquivo .env")

        self.driver.find_element(By.ID, "usuario").send_keys(self.login)
        self._human_delay(0.5, 1)
        self.driver.find_element(By.ID, "senha").send_keys(self.senha)
        self._human_delay(0.5, 1)
        self.driver.find_element(By.ID, "login-submit").click()
        self.handle_overlays()
        self._pause_between_actions(2)
        exibir_mensagem_sucesso("Login realizado com sucesso")

    def navegar_meus_tratamentos(self):
        """Navega para a página 'Meus Tratamentos'."""
        try:
            self._pause_between_actions(2)
            self.handle_overlays()
            menu = self.wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Meus Tratamentos")))
            self.driver.execute_script("arguments[0].click();", menu)
            self.handle_overlays()
            self._human_delay(2, 3)
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao navegar para 'Meus Tratamentos': {e}")
            raise

    def buscar_guia(self, id_guia):
        """Busca uma guia específica na página 'Meus Tratamentos'."""
        try:
            self.handle_overlays()
            campo_busca = self.wait.until(EC.presence_of_element_located((By.ID, "searchGTO")))
            # Otimização: insere o valor via JS para maior velocidade
            self.driver.execute_script("arguments[0].value = arguments[1];", campo_busca, id_guia)
            self._human_delay(0.5, 1)
            campo_busca.send_keys(Keys.ENTER)
            self.handle_overlays()
            self._human_delay(1, 2)
            link_guia = self.wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[text()='{id_guia}']")))
            self.driver.execute_script("arguments[0].click();", link_guia)
            self.handle_overlays()
            self._pause_between_actions(5)
            exibir_mensagem_sucesso(f"Guia {id_guia} aberta")
            return True
        except TimeoutException:
            exibir_mensagem_erro(f"Guia {id_guia} não encontrada")
            return False

    def verificar_procedimento_ja_confirmado(self):
        """Verifica se o procedimento já foi confirmado (ícone de dólar laranja #de9045)."""
        try:
            self.handle_overlays()
            icones_dollar = self.driver.find_elements(By.XPATH, "//i[contains(@class, 'fa-dollar') and contains(@style, 'color')]")
            
            if not icones_dollar:
                return False
            
            for icone_dollar in icones_dollar:
                cor_icone = self.driver.execute_script(
                    "return window.getComputedStyle(arguments[0]).getPropertyValue('color');", 
                    icone_dollar
                )
                style_attr = icone_dollar.get_attribute('style') or ""
                
                tem_cor_laranja = '#de9045' in style_attr.lower() or 'rgb(222, 144, 69)' in cor_icone
                
                if tem_cor_laranja:
                    exibir_mensagem_info("✓ Ícone $ LARANJA detectado - procedimento já confirmado. Indo direto para anexos.")
                    return True
            
            return False
        except Exception:
            return False

    def verificar_guia_ja_finalizada(self):
        """Verifica se a guia já foi completamente finalizada e paga."""
        try:
            self.handle_overlays()
            icones_dollar = self.driver.find_elements(By.XPATH, "//i[contains(@class, 'fa-dollar') and contains(@style, 'color')]")
            
            if not icones_dollar:
                return False
            
            icone_dollar = icones_dollar[0]
            cor_icone = self.driver.execute_script(
                "return window.getComputedStyle(arguments[0]).getPropertyValue('color');", 
                icone_dollar
            )
            style_attr = icone_dollar.get_attribute('style') or ""
            
            nao_e_laranja = '#de9045' not in style_attr.lower() and 'rgb(222, 144, 69)' not in cor_icone
            
            if nao_e_laranja:
                exibir_mensagem_info(f"✓ Guia já completamente finalizada (ícone $ detectado, cor: {cor_icone}).")
                return True
            return False
        except Exception:
            return False

    def verificar_ja_esta_na_tela_anexos(self):
        """Verifica se já está na tela de anexos."""
        try:
            self.handle_overlays()
            titulo_presente = len(self.driver.find_elements(By.XPATH, "//title[contains(text(), 'Gestão Rede')]")) > 0
            botao_anexos_presente = len(self.driver.find_elements(By.ID, "btnAnexosProcedimento")) > 0
            
            if titulo_presente or botao_anexos_presente:
                exibir_mensagem_info("Detectado que a guia já está na tela de anexos. Pulando confirmação de procedimento.")
                return True
            return False
        except Exception:
            return False

    def _extrair_data_autorizacao(self):
        """Extrai a data de autorização da página da guia."""
        try:
            self.handle_overlays()
            elemento_p = self.wait.until(EC.presence_of_element_located((By.XPATH, "//p[strong[text()='Data da Autorização: ']]")))
            match = re.search(r'\d{2}/\d{2}/\d{4}', elemento_p.text)
            return match.group(0) if match else None
        except TimeoutException:
            exibir_mensagem_alerta("Data de autorização não encontrada")
            return None

    def _set_and_verify_date(self, date_element, date_to_set, checkbox):
        """Preenche e verifica um campo de data."""
        try:
            self.handle_overlays()
            self.driver.execute_script("arguments[0].checked=true; arguments[1].removeAttribute('disabled'); $(arguments[0]).trigger('change');", checkbox, date_element)
            date_element.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            date_element.send_keys(date_to_set)
            ActionChains(self.driver).move_to_element(self.driver.find_element(By.ID, "observacaoJustificativa")).click().perform() # Tira o foco
            self.handle_overlays()
            return date_element.get_attribute('value') == date_to_set
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao inserir data: {e}")
            return False
        
    def _handle_faturamento_guia_modal(self, timeout=3):
        """Lida com o modal 'Faturamento através de' se ele aparecer."""
        try:
            botao_guia = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((By.ID, "btnGuia")))
            self.driver.execute_script("arguments[0].click();", botao_guia)
            self.handle_overlays(main_timeout=15)
        except TimeoutException:
            pass

    def _verificar_se_requer_dente(self, codigo_procedimento):
        """Verifica se um procedimento requer seleção de dente baseado no mapeamento."""
        for key, dados in MAPEAMENTO_PROCEDIMENTOS.items():
            if dados["codigo_completo"] == codigo_procedimento:
                requer = dados.get("requer_dente", True)
                return requer
        exibir_mensagem_alerta(f"Procedimento {codigo_procedimento} não encontrado no mapeamento. Assumindo que requer dente.")
        return True        

    def _handle_procedimento_complementar(self, proc_code_alvo, proc_desc_alvo, dente_valor_alvo, observacao_alvo="", codigo_guia=""):
        """Lida com o modal 'Procedimento Complementar' com validação rigorosa."""
        try:
            exibir_mensagem_info("Aguardando modal de procedimento complementar...")
            self.handle_overlays()
            wait = WebDriverWait(self.driver, 15)
            
            requer_dente = self._verificar_se_requer_dente(proc_code_alvo)
            celula_procedimento_xpath = f"//td[starts-with(normalize-space(.), '{proc_code_alvo} -')]"
            celula_procedimento = wait.until(EC.presence_of_element_located((By.XPATH, celula_procedimento_xpath)))
            exibir_mensagem_sucesso(f"Modal e procedimento '{proc_code_alvo}' encontrados. Validando...")

            linha = celula_procedimento.find_element(By.XPATH, "./..")
            texto_procedimento_site = linha.find_element(By.XPATH, ".//td[2]").text.upper()
            
            if proc_desc_alvo.upper() not in texto_procedimento_site:
                exibir_mensagem_erro(f"VALIDAÇÃO FALHOU! O nome do procedimento '{proc_code_alvo}' não corresponde.")
                return False
            exibir_mensagem_sucesso("Validação de nome OK.")
            
            checkbox_proc = linha.find_element(By.XPATH, ".//input[@type='checkbox']")
            self.driver.execute_script("arguments[0].click();", checkbox_proc)
            self.handle_overlays()
            self._human_delay(0.5, 1)

            if requer_dente and dente_valor_alvo:
                exibir_mensagem_info(f"Selecionando dente {dente_valor_alvo} para procedimento {proc_code_alvo}...")
                Select(linha.find_element(By.XPATH, ".//select")).select_by_value(dente_valor_alvo)
                self.handle_overlays()
                self._human_delay(0.5, 1)
            else:
                exibir_mensagem_info(f"Procedimento {proc_code_alvo} NÃO requer seleção de dente. Pulando...")

            if observacao_alvo and codigo_guia:
                try:
                    codigo_base = codigo_guia[:8] if len(codigo_guia) >= 8 else codigo_guia
                    campo_obs_id = f"obs_{proc_code_alvo}_{codigo_base}"
                    exibir_mensagem_info(f"Tentando preencher observação no campo '{campo_obs_id}'...")
                    
                    campo_obs = self.driver.find_element(By.ID, campo_obs_id)
                    campo_obs.clear()
                    campo_obs.send_keys(observacao_alvo)
                    self._human_delay(0.5, 1)
                    exibir_mensagem_sucesso(f"Observação inserida: '{observacao_alvo[:50]}...'")
                except Exception:
                    pass

            self.driver.find_element(By.ID, "btnAvancarAlteracao").click()
            self.handle_overlays()
            time.sleep(9)
            exibir_mensagem_sucesso("Procedimento complementar salvo.")
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro inesperado ao processar o modal complementar: {e}")
            self.log_manager.registrar_falha(proc_code_alvo, "Erro no modal complementar", traceback.format_exc())
            return False


    def confirmar_realizacao_procedimento(self, proc_code, proc_desc, dente_valor, observacao="", codigo_guia=""):
        """Confirma a data de realização e lida com os modais subsequentes."""
        exibir_mensagem_info("Confirmando realização do procedimento...")
        self._pause_between_actions(4)
        data_autorizacao = self._extrair_data_autorizacao()
        if not data_autorizacao: return False
        try:
            checkboxes = self.driver.find_elements(By.XPATH, "//input[contains(@id, 'chk_atualiza_data_')]")
            if checkboxes:
                for chk in checkboxes:
                    proc_id_chk = chk.get_attribute('id').split('_')[-1]
                    campo_data = self.driver.find_element(By.ID, f"dt_conf_{proc_id_chk}")
                    if not self._set_and_verify_date(campo_data, data_autorizacao, chk): return False
            
            self._human_delay(1, 2)
            self.wait.until(EC.element_to_be_clickable((By.ID, "btnConfirmar"))).click()
            self.handle_overlays()
            self._pause_between_actions(6)
            
            if self.usa_modal_complementar:
                if proc_code and dente_valor:
                    if not self._handle_procedimento_complementar(proc_code, proc_desc, dente_valor, observacao, codigo_guia):
                        return False
            
            self._pause_between_actions(10)
            self._handle_faturamento_guia_modal()
            self._pause_between_actions(6)
            exibir_mensagem_sucesso("Etapa de confirmação concluída.")
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro ao confirmar procedimento: {e}")
            return False

    def buscar_arquivos_guia(self, id_guia, pasta_base):
        """Busca arquivos locais correspondentes a uma guia."""
        exibir_mensagem_info(f"Buscando arquivos para guia {id_guia}...")
        arquivos_encontrados = []
        id_parcial = id_guia[-4:]
        for f in Path(pasta_base).rglob('*'):
            if not f.is_file(): continue
            nome_upper = f.stem.upper()
            if id_guia in nome_upper or id_parcial in nome_upper:
                tipo_arquivo = 'documentacao'
                if nome_upper.startswith('FI'): tipo_arquivo = 'foto-inicial'
                elif nome_upper.startswith('FF'): tipo_arquivo = 'foto-final'
                elif 'RX' in nome_upper or nome_upper.startswith(('RF', 'RI')): tipo_arquivo = 'raio-x'
                elif 'LAUDO' in nome_upper: tipo_arquivo = 'laudo'
                elif nome_upper.startswith(('G', 'GUIA', 'ASS')): tipo_arquivo = 'guia'
                arquivos_encontrados.append({'caminho': str(f.resolve()), 'tipo': tipo_arquivo})
        return arquivos_encontrados

    def _abrir_modal_anexos(self):
        """Aguarda a página estabilizar e abre o modal de anexos."""
        exibir_mensagem_info("Preparando para anexar arquivos...")
        try:
            time.sleep(5)
            self.handle_overlays()
            botao_anexos = self.wait.until(EC.element_to_be_clickable((By.ID, "btnAnexosProcedimento")))
            time.sleep(2) 
            self.driver.execute_script("arguments[0].click();", botao_anexos)
            time.sleep(5)
            self.handle_overlays()
            self.wait.until(EC.visibility_of_element_located((By.ID, "modalAnexosProcedimento")))
            exibir_mensagem_sucesso("Modal de anexos aberto.")
        except Exception as e:
            exibir_mensagem_erro(f"Não foi possível abrir o modal de anexos de forma estável: {e}")
            raise

    def _fechar_modal_anexos(self):
        """Fecha o modal de anexos."""
        try:
            self.handle_overlays()
            self.driver.find_element(By.XPATH, "//div[@id='modalAnexosProcedimento' and contains(@style,'display: block')]//button[@data-dismiss='modal']").click()
            self.wait.until(EC.invisibility_of_element_located((By.ID, "modalAnexosProcedimento")))
            self.handle_overlays()
        except Exception: pass

    def _anexar_um_arquivo(self, anexo_info):
        """Executa a lógica de upload para um único arquivo."""
        caminho_arquivo, tipo = anexo_info['caminho'], anexo_info['tipo']
        exibir_mensagem_info(f"Anexando: {Path(caminho_arquivo).name}")
        
        self.handle_overlays()
        select_tipo = Select(self.wait.until(EC.presence_of_element_located((By.ID, "tipoAnexoModal"))))
        if tipo == 'raio-x': select_tipo.select_by_value("1")
        elif tipo == 'laudo': select_tipo.select_by_value("2")
        elif tipo == 'guia': select_tipo.select_by_value("3")
        else: select_tipo.select_by_value("4")
        
        self.handle_overlays()
        self._human_delay(0.3, 0.5)

        if tipo == 'raio-x':
            try: WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//label[@for='periapical']"))).click()
            except TimeoutException: pass
        
        Select(self.driver.find_element(By.ID, "processoAnexo")).select_by_value("2")
        self.handle_overlays()
        self._human_delay(0.3, 0.5)
        
        self.driver.find_element(By.ID, "arquivoAnexo").send_keys(caminho_arquivo)
        self._human_delay(0.5, 1)
        self.handle_overlays()
        
        for checkbox in self.driver.find_elements(By.XPATH, "//table[@id='tblVincularProcedimentos']//input[@type='checkbox']"):
            if not checkbox.is_selected(): checkbox.click()
        
        self.handle_overlays()
        self.driver.find_element(By.ID, "btnIncluirAnexos").click()
        self.handle_overlays(main_timeout=15)

    def anexar_todos_os_arquivos(self, lista_de_anexos):
        """Itera sobre uma lista de arquivos e anexa cada um."""
        if not lista_de_anexos:
            exibir_mensagem_alerta("Nenhum arquivo encontrado para anexar.")
            return True
        exibir_mensagem_info(f"Anexando {len(lista_de_anexos)} arquivo(s)...")
        self._pause_between_actions(4)
        try:
            for anexo in lista_de_anexos:
                self._abrir_modal_anexos()
                self._anexar_um_arquivo(anexo)
                self._fechar_modal_anexos()
                self._human_delay(1, 2)
            exibir_mensagem_sucesso("Todos os arquivos foram anexados.")
            self._pause_between_actions(8)
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro durante o anexo de arquivos: {e}")
            self.log_manager.registrar_falha("N/A", "Erro no anexo", traceback.format_exc())
            return False

    def submeter_para_pagamento(self):
        """Clica no botão final para submeter a guia para pagamento."""
        exibir_mensagem_info("Submetendo guia para pagamento...")
        self._pause_between_actions(4)
        try:
            self.handle_overlays()
            checkboxes = self.driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            
            if checkboxes:
                for checkbox in checkboxes:
                    try:
                        if not checkbox.is_selected():
                            self.driver.execute_script("arguments[0].click();", checkbox)
                            self._human_delay(0.3, 0.5)
                    except Exception: pass
            
            self.handle_overlays()
            self._human_delay(1, 2)
            
            self.wait.until(EC.element_to_be_clickable((By.ID, "btnConfirmar"))).click()
            self.handle_overlays()
            self._pause_between_actions(6)
            self.wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-dollar')]")))
            exibir_mensagem_sucesso("Guia submetida e pagamento confirmado ($)")
            return True
        except Exception as e:
            exibir_mensagem_erro(f"Erro na submissão final: {e}")
            return False
