<div align="center">

# 🦷 Dental RPA — Porto Seguro & SulAmérica Odonto

**Automação de Faturamento Odontológico de Ponta a Ponta**

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=for-the-badge&logo=Selenium&logoColor=white)
![Automation](https://img.shields.io/badge/RPA-Production%20Ready-blue?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Deployed%20%26%20Active-success?style=for-the-badge)

> Sistema RPA desenvolvido como solução freelancer real para uma clínica
> odontológica. O código foi higienizado e modularizado para exibição
> em portfólio, preservando a privacidade e segurança do cliente.

</div>

---

## 📌 Contexto do Projeto

Clínicas odontológicas credenciadas por operadoras como Porto Seguro e
SulAmérica precisam enviar **guias de atendimento** mensalmente para
receber os repasses financeiros. Esse processo envolve:

- Acessar portais web distintos para cada operadora
- Preencher manualmente dados clínicos de cada procedimento
- Conferir códigos de dentes, faces e tratamentos
- Fazer upload de documentos (raios-X, fotos, laudos) vinculados
  individualmente a cada guia

Na clínica cliente, esse fluxo consumia **horas semanais** da equipe
administrativa, com alto risco de **glosas** — rejeições de faturamento
causadas por erros de preenchimento — que impactavam diretamente o
caixa da clínica.

**O objetivo era claro: eliminar o trabalho manual sem abrir mão da
precisão.**

---

## 🎯 Resultados Entregues

| Métrica | Antes | Depois |
|---|---|---|
| Tempo médio por guia | ~8 minutos | ~40 segundos |
| Erros de vinculação de anexos | Frequentes | Zero |
| Capacidade de processamento | ~20 guias/dia | Ilimitado |
| Rastreabilidade | Nenhuma | Log completo por execução |

---

## 🛠️ Como o Sistema Funciona

O robô executa o fluxo completo de faturamento de forma autônoma,
replicando com precisão cada ação que um operador humano realizaria
no portal — mas em uma fração do tempo.
```
┌─────────────────────────────────────────────────────────┐
│                    FLUXO DE EXECUÇÃO                    │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  📊 Planilha Excel (.xlsx)                              │
│         │                                               │
│         ▼                                               │
│  🔐 Login seguro no portal da operadora                 │
│         │                                               │
│         ▼                                               │
│  🔍 Busca e validação da guia                           │
│         │                                               │
│         ▼                                               │
│  ✍️  Preenchimento dos dados clínicos                   │
│     (procedimentos, dentes, faces, observações)         │
│         │                                               │
│         ▼                                               │
│  📎 Identificação e upload automático de anexos         │
│     (raios-X, fotos, laudos — por correspondência)      │
│         │                                               │
│         ▼                                               │
│  📋 Geração de log de auditoria (.csv / .txt)           │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

---

## 💎 Funcionalidades Técnicas

### 🤖 Core do RPA
- **Processamento em batch** direto de planilhas `.xlsx` — sem
  necessidade de reformatação
- **Navegação resiliente** com detecção automática de overlays,
  modais e mudanças de DOM
- **Delays randômicos** entre ações para simular comportamento humano
  e evitar bloqueios anti-bot
- **Retry inteligente** em caso de falhas de rede ou timeouts

### 📎 Gestão de Documentos
- Varredura automática da pasta de anexos
- **Correspondência parcial** entre nome do arquivo e número da guia
- Vinculação e upload sem intervenção manual
- Validação de sucesso após cada upload

### 🔐 Segurança
- Credenciais isoladas em `.env` — nunca expostas no código
- `.gitignore` configurado para proteger dados sensíveis do cliente
- Dados clínicos tratados localmente, sem persistência em servidores
  externos

### 📊 Auditoria e Rastreabilidade
- Log gerado automaticamente a cada execução
- Registro de guias processadas com sucesso e falhas com motivo
- Relatório exportado em `.csv` para análise pela equipe da clínica

---

## 🚀 Stack Tecnológica

| Tecnologia | Função |
|---|---|
| **Python 3.8+** | Linguagem core |
| **Selenium WebDriver** | Automação e manipulação do DOM |
| **Webdriver-Manager** | Gerenciamento automático do ChromeDriver |
| **Openpyxl** | Leitura e parsing das planilhas Excel |
| **Python-Dotenv** | Gestão segura de credenciais |
| **Logging + CSV** | Auditoria e relatórios de execução |

---

## 📂 Arquitetura
```
Dental-RPA/
├── src/
│   ├── main.py          # Orquestrador principal e menu interativo
│   ├── bot.py           # Engine do RPA — lógica Selenium e fluxos
│   ├── constants.py     # Mapeamento de procedimentos e configurações
│   └── utils.py         # Helpers de UI, parsing e utilitários gerais
│
├── arquivos/            # Banco local de anexos (RX, laudos, fotos)
├── logs/                # Relatórios automáticos de cada execução
├── .env.example         # Template de variáveis de ambiente
├── .gitignore           # Proteção de dados sensíveis
└── iniciar_bot.bat      # Instalador + inicializador (Windows)
```

---

## ⚙️ Como Executar
```bash
# 1. Clone o repositório
git clone https://github.com/kauankelvin7/dental-rpa.git

# 2. Configure as credenciais
cp .env.example .env
# Edite o .env com suas credenciais do portal

# 3. No Windows, execute o inicializador
iniciar_bot.bat
# Ele instala as dependências e abre o menu automaticamente

# 4. Ou manualmente
pip install -r requirements.txt
python src/main.py
```

---

## 🧠 Decisões de Engenharia

**Por que Selenium e não uma API?**
Os portais das operadoras não disponibilizam APIs públicas. A única
forma de automatizar é via interface web — tornando o Selenium a
escolha natural e mais robusta para o problema.

**Por que arquitetura modular?**
Separar `bot.py`, `constants.py` e `utils.py` permite que novos
fluxos (ex: SulAmérica) sejam adicionados sem reescrever a lógica
central. O `constants.py` concentra todos os mapeamentos de
procedimentos, facilitando manutenção.

**Por que delays randômicos?**
Portais com proteção anti-bot detectam padrões de clique regulares.
Delays variáveis entre 1.5s e 3.5s tornam o robô indistinguível de
um operador humano nas métricas do servidor.

---

<div align="center">

**Desenvolvido por Kauan Kelvin**
*Software Engineer & RPA Developer*

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/kauan-kelvin)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=flat-square&logo=github&logoColor=white)](https://github.com/kauankelvin7)

</div>
