import os
from datetime import datetime
from constants import Cores

# ======================== FUNÇÕES DE INTERFACE DE USUÁRIO ========================
def exibir_banner():
    """Limpa a tela e exibe o banner principal do sistema."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print(Cores.CYAN + Cores.BOLD)
    print("╔" + "═" * 78 + "╗")
    print("║" + " 🦷 SISTEMA DE AUTOMAÇÃO PORTO SEGURO 🦷 ".center(84) + "║")
    print("║" + f" Versão 2.1 | {datetime.now().year} ".center(84) + "║")
    print("╚" + "═" * 78 + "╝")
    print(Cores.RESET)

def exibir_mensagem_sucesso(mensagem):
    print(f"{Cores.VERDE}✓ {mensagem}{Cores.RESET}")

def exibir_mensagem_erro(mensagem):
    print(f"{Cores.VERMELHO}✗ {mensagem}{Cores.RESET}")

def exibir_mensagem_info(mensagem):
    print(f"{Cores.CYAN}ℹ {mensagem}{Cores.RESET}")

def exibir_mensagem_alerta(mensagem):
    print(f"{Cores.AMARELO}⚠ {mensagem}{Cores.RESET}")
