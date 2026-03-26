# 🦷 Dental Billing & RPA - Porto Seguro Odonto Automation

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-43B02A?style=for-the-badge&logo=Selenium&logoColor=white)
![Automation](https://img.shields.io/badge/RPA-Software%20Bot-blue?style=for-the-badge)

Este projeto é um sistema de **Automação de Processos Robóticos (RPA)** de nível profissional desenvolvido para otimizar o fluxo de faturamento de clínicas odontológicas credenciadas pela Porto Seguro. O sistema elimina o trabalho manual repetitivo, reduz erros de processamento e acelera o recebimento de repasses da seguradora.

## 🌟 O Desafio (Business Case)

Clínicas odontológicas enfrentam um gargalo administrativo no faturamento:
1.  **Entrada Manual de Dados**: Milhares de guias processadas uma a uma.
2.  **Validação Complexa**: Conferência manual de códigos de procedimentos e dentes.
3.  **Gestão de Documentação**: Upload de centenas de exames (Raios-X, fotos e laudos) que devem ser vinculados corretamente a cada guia.

**Resultado antes da automação**: Horas de trabalho administrativo desperdiçadas e alto risco de glosas (faturamentos rejeitados) por erros humanos.

## 🛠️ A Solução (The Tech Solution)

Desenvolvi um robô robusto em **Python** que interage diretamente com o portal web, simulando o comportamento humano com inteligência e precisão.

### 💎 Principais Funcionalidades

*   **Processamento em Batch**: Lê guias diretamente de planilhas Excel (`.xlsx`).
*   **Gestão Inteligente de Arquivos**: Identifica e anexa automaticamente documentos locais (RX, fotos, laudos) utilizando lógica de correspondência parcial.
*   **Security-First**: Implementação de variáveis de ambiente (`.env`) e `.gitignore` para proteção total de dados sensíveis e credenciais do cliente.
*   **Modularização de Código**: Arquitetura modular separando lógica de negócio, utilitários, constantes e interface.
*   **Resiliência & Human Bypass**: Implementação de delays randômicos e detecção dinâmica de overlays para evitar bloqueios de bots e garantir estabilidade.
*   **Logs & Auditoria**: Geração de relatórios detalhados (`.csv` e `.txt`) de sucessos e falhas para controle total da operação.

## 🚀 Stack Tecnológica

*   **Python 3.8+**: Linguagem core pela versatilidade e bibliotecas de automação.
*   **Selenium WebDriver**: Manipulação dinâmica do DOM e automação de navegação web.
*   **Webdriver-Manager**: Gerenciamento automático de binários (Chrome).
*   **Python-Dotenv**: Gestão segura de segredos e credenciais.
*   **Openpyxl**: Manipulação eficiente de planilhas Excel.

## 📂 Arquitetura do Projeto

```text
Automacao-PortoSeguro/
├── src/
│   ├── main.py        # Orquestrador e Menu Interativo
│   ├── bot.py         # Lógica Core do RPA & Selenium
│   ├── constants.py   # Mapeamento de Procedimentos & Configurações
│   └── utils.py       # Interface de Usuário & Helpers
├── arquivos/          # Banco de anexos (Raios-X, Laudos, etc)
├── logs/              # Relatórios automáticos de auditoria
├── .env               # Credenciais de acesso (Protegido)
└── iniciar_bot.bat    # Executável de instalação e inicialização
```

## ⚙️ Como Executar

1.  Clone este repositório.
2.  Crie um arquivo `.env` baseado no `.env.example` e insira suas credenciais.
3.  Execute o `iniciar_bot.bat` no Windows (ele instalará as dependências e iniciará o menu).

---
### 👨‍💻 Desenvolvedor
**Kauan Kelvin** - *Software Engineer & RPA Developer*
> Este projeto foi desenvolvido como solução freelancer real. O código foi higienizado e modularizado para exibição no portfólio, garantindo a privacidade e segurança do cliente original.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/SEU-LINK)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=flat-square&logo=github&logoColor=white)](https://github.com/kauankelvin7)
