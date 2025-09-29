import pandas as pd
import pyperclip 
import glob
import os
import re

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time

options = Options()
options.add_argument("--start-maximized")

# --- Configurações de Caminho ---
NOME_PLANILHA = 'guias_porto_seguro.xlsx'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVOS_FOLDER = os.path.join(SCRIPT_DIR, 'arquivos')
CAMINHO_PLANILHA = os.path.join(ARQUIVOS_FOLDER, NOME_PLANILHA)

CHROME_DRIVER_PATH = os.path.join(SCRIPT_DIR, 'chromedriver.exe') 

# --- Função de Ação do Convênio (MÓDULO PORTO SEGURO) ---
def processar_porto_seguro(dados_guia, documentos_encontrados):
    """
    Função dedicada a interagir com o portal da Porto Seguro (Usando Selenium).
    """
    # Credenciais
    LOGIN = "00260115"
    SENHA = "@Isbi4420"
    URL_LOGIN = "https://www.dentistaportoseguro.com.br/" 

    data_atendimento = dados_guia['Data de Atendimento']
    numero_guia = dados_guia['ID da Guia/Beneficiário']
    
    driver = None
    
    print(f"   ➡️ INICIANDO MÓDULO SELENIUM. Data a ser usada: {data_atendimento}")
    
    # 1. Configurar e Abrir o Navegador
    try:
        service = Service(executable_path=CHROME_DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(URL_LOGIN)        
        wait = WebDriverWait(driver, 15)

        # 2. Módulo de Login: SELECTORES REAIS
        print("   ➡️ Tentando Login...")
        
        OVERLAY_SELECTOR = "spin_modal_overlay"
        # Seletores que o cliente validou como funcionais
        campo_login = wait.until(EC.presence_of_element_located((By.ID, "usuario"))) 
        campo_senha = wait.until(EC.presence_of_element_located((By.ID, "senha")))
        botao_entrar = wait.until(EC.element_to_be_clickable((By.ID, "login-submit")))
        
        campo_login.send_keys(LOGIN)
        campo_senha.send_keys(SENHA)
        botao_entrar.click()
        
        wait.until(EC.url_changes(URL_LOGIN)) 
        
        print("   ⭐ Navegação para a área logada confirmada!")

        # 3. Módulo de Navegação: Clicar em "Meus Tratamentos "
        print("   ➡️ Tentando navegar para a área de Meus Tratamentos...")

        try:
            link_tratamento = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Meus Tratamentos"))) 
            wait.until(EC.invisibility_of_element_located((By.ID, OVERLAY_SELECTOR)))

            driver.execute_script("arguments[0].click();", link_tratamento)

            # 4. MÓDULO DE BUSCA
            CAMPO_BUSCA_ID = "searchGTO"
            wait.until(EC.presence_of_element_located((By.ID, CAMPO_BUSCA_ID)))
            time.sleep(4) 
            
            print(f"   ➡️ Buscando paciente pelo ID: {numero_guia}")
            
            campo_busca = driver.find_element(By.ID, CAMPO_BUSCA_ID)
            campo_busca.send_keys(numero_guia) 
            time.sleep(2) 

            campo_busca.send_keys(Keys.ENTER) 

            time.sleep(2) 
            
            LINK_PRONTUARIO_XPATH = f"//a[text()='{numero_guia}']"
            
            wait.until(EC.invisibility_of_element_located((By.ID, OVERLAY_SELECTOR)))
            
            link_prontuario = wait.until(EC.element_to_be_clickable((By.XPATH, LINK_PRONTUARIO_XPATH)))
            link_prontuario.click()

            time.sleep(2)
            print("   ✅ Prontuário aberto. Próximo passo: Inserir Dados e Anexar Arquivos.")
            
        except Exception as e:
            print(f"   ❌ ERRO AO NAVEGAR PARA 'Meus Tratamentos': {e}")
            return
                    
        # IDs dos campos
        CHECKBOX_ID = "chk_atualiza_data_1"
        CAMPO_DATA_ID = "dt_conf_1"
        
        data_atendimento_obj = dados_guia['Data de Atendimento']
        
        if hasattr(data_atendimento_obj, 'strftime'):
            data_atendimento_str = data_atendimento_obj.strftime('%d/%m/%Y')
        else:
            data_atendimento_str = str(data_atendimento_obj).split(' ')[0]
            if '-' in data_atendimento_str:
                partes = data_atendimento_str.split('-')
                data_atendimento_str = f"{partes[2]}/{partes[1]}/{partes[0]}"
                
        print(f"   ➡️ Data formatada: {data_atendimento_str}")


        XPATH_TODOS_CHECKBOXES = "//input[starts-with(@id, 'chk_atualiza_data_')]"
        todos_checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_TODOS_CHECKBOXES)))
        
        time.sleep(1)

        for checkbox_input in todos_checkboxes:
            
            seq_id = checkbox_input.get_attribute('id').split('_')[-1]
            CAMPO_DATA_ID_LOOP = f"dt_conf_{seq_id}"
            
            campo_data = driver.find_element(By.ID, CAMPO_DATA_ID_LOOP) 
            
            driver.execute_script("arguments[0].click();", campo_data)

            driver.execute_script(
                f"""
                var cb = arguments[0]; 
                var dt = arguments[1]; 
                
                // 1. Força o estado CHECKED
                cb.checked = true; 
                
                // 2. Remove 'disabled' (Habilita o campo de data)
                dt.removeAttribute('disabled'); 
                
                // 3. Dispara o evento CHANGE no checkbox (Gatilho para a validação do sistema)
                $(cb).trigger('change'); 
                """, checkbox_input, campo_data
            )
                
            print(f"   ➡️ Inserindo data '{dados_guia['Data de Atendimento']}'...")
            pyperclip.copy(data_atendimento_str)
            campo_data = wait.until(EC.element_to_be_clickable((By.ID, CAMPO_DATA_ID)))
            campo_data.click()
            campo_data.clear()
            campo_data.send_keys(Keys.CONTROL, 'v') 
            
        # 7. MÓDULO DE UPLOAD DE ARQUIVOS
        
        BOTAO_ANEXOS_ID = "btnAnexosProcedimento"
        CAMPO_UPLOAD_ID = "arquivoAnexo"
        
        print("   ➡️ Clicando no link para Anexar Documentos...")
        
        botao_anexos = wait.until(EC.element_to_be_clickable((By.ID, BOTAO_ANEXOS_ID)))
        botao_anexos.click()
        
        print("   ✅ Pop-up de Anexos aberto. Iniciando upload dinâmico.")

        
        # MÓDULO DE UPLOAD EM SI:
        for arquivo_path in documentos_encontrados:
            
            # Localiza o campo de upload dentro do modal
            campo_upload = wait.until(EC.presence_of_element_located((By.ID, CAMPO_UPLOAD_ID)))
            
            # Injeta o caminho ABSOLUTO do arquivo (simula a abertura da janela de arquivos)
            campo_upload.send_keys(os.path.abspath(arquivo_path))
            
            print(f"   ⬆️ Arquivo {os.path.basename(arquivo_path)} injetado.")
            
            # 7.3. PREENCHER DADOS DO ANEXO E CLICAR EM 'INCLUIR'
            
            # Classificação do arquivo (Baseado no nome: RF_ ou G_)
            if os.path.basename(arquivo_path).startswith('RF_'):
                tipo_anexo = "Raio-X" 
            elif os.path.basename(arquivo_path).startswith('G_'):
                tipo_anexo = "GTO - Guia Tratamento Odontológico"
            else:
                tipo_anexo = "Documentação (Fotos)" # Fallback

            # Seleciona o TIPO de Anexo (select com ID "tipoAnexoModal")
            tipo_anexo_modal = driver.find_element(By.ID, "tipoAnexoModal")
            tipo_anexo_modal.send_keys(tipo_anexo) 
            
            # Seleciona o PROCESSO: Pagamento (Opção 2)
            processo_anexo_modal = driver.find_element(By.ID, "processoAnexo")
            processo_anexo_modal.send_keys("Pagamento") 
            
            # Clica no botão Incluir no modal
            botao_incluir = driver.find_element(By.ID, "btnIncluirAnexos")
            botao_incluir.click()
            
            # Espera o processo de upload e inclusão terminar
            time.sleep(3) 

        
        # 7.4. FINALIZAÇÃO DA GUIA
        
        # O botão 'Confirmar' que finaliza tudo está no footer (ID: btnConfirmar)
        driver.find_element(By.ID, "btnConfirmar").click() 
        
        print(f"   ✅ Processamento da Guia {numero_guia} FINALIZADO COM SUCESSO.")
                    
    except (TimeoutException, NoSuchElementException) as e:
        print(f"   ❌ ERRO NO SELENIUM: Falha ao encontrar um elemento. Motivo: {type(e).__name__}")
        print("   ⚠️ O robô falhou no link de busca ou no link do prontuário.")
    except Exception as e:
        print(f"   ❌ ERRO INESPERADO: {e}")
    finally:
        if driver:
             # driver.quit() # Descomente APENAS quando o robô estiver pronto
             pass 
        
# --- Função Principal: O Robô Orquestrador ---
def iniciar_automacao():
    try:
        # 1. Leitura da Planilha
        df = pd.read_excel(CAMINHO_PLANILHA)
        print(f"Planilha '{NOME_PLANILHA}' lida com sucesso.\n")

        # 2. Loop de Processamento por Paciente
        for index, row in df.iterrows():
            numero_guia = str(row['ID da Guia/Beneficiário']).strip()
            convenio = row['Convênio']
            
            print(f"==========================================================================================")
            print(f"ORQUESTRADOR: Iniciando Processo | Convênio: {convenio} | Guia: {numero_guia}")
            
            # --- Roteamento do Orquestrador ---
            if convenio == 'Porto Seguro':
                
                padrao_busca = os.path.join(ARQUIVOS_FOLDER, f"*{numero_guia}*.*")
                documentos_encontrados = glob.glob(padrao_busca)
                
                if documentos_encontrados:
                    print(f"✅ {len(documentos_encontrados)} Documentos encontrados para anexar.")
                    
                    # Prepara os dados do paciente
                    dados_paciente = {
                        'ID da Guia/Beneficiário': numero_guia,
                        'Nome do Paciente': row['Nome do Paciente'],
                        'Data de Atendimento': row['Data de Atendimento'] 
                    }
                    
                    # 3. CHAMA O MÓDULO DE AÇÃO
                    processar_porto_seguro(dados_paciente, documentos_encontrados)
                    
                else:
                    print(f"❌ ATENÇÃO: Nenhum documento encontrado. Pulando paciente.")
            
            else:
                print(f"⚠️ CONVÊNIO NÃO MAPEADO: Não há módulo para {convenio}. Pulando.")
                
            print("==========================================================================================\n")

    except FileNotFoundError:
        print(f"\n❌ ERRO FATAL: O arquivo '{CAMINHO_PLANILHA}' não foi encontrado.")
    except KeyError as e:
        print(f"\n❌ ERRO FATAL: A coluna {e} não foi encontrada na planilha. Verifique o nome das colunas.")
    except Exception as e:
        print(f"\n❌ Ocorreu um erro inesperado: {e}")

iniciar_automacao()