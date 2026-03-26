import os

# ======================== CONFIGURAÇÃO DE CORES PARA O CONSOLE ========================
class Cores:
    """Códigos de escape ANSI para formatação de texto no terminal."""
    RESET = '\033[0m'
    BOLD = '\033[1m'
    VERDE = '\033[92m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    AZUL = '\033[94m'
    CYAN = '\033[96m'
    MAGENTA = '\033[95m'
    BG_VERDE = '\033[102m\033[30m'

if os.name == 'nt':
    os.system('')

# ======================== DADOS E MAPEAMENTOS ========================
MAPEAMENTO_PROCEDIMENTOS = {
    "1": {"codigo_completo": "82000468", "descricao": "CONTROLE DE HEMORRAGIA COM APLICACAO", "requer_dente": False},
    "2": {"codigo_completo": "82000484", "descricao": "CONTROLE DE HEMORRAGIA SEM APLICACAO", "requer_dente": False},
    "3": {"codigo_completo": "82000859", "descricao": "EXODONTIA DE RAIZ RESIDUAL", "requer_dente": True},
    "4": {"codigo_completo": "82000875", "descricao": "EXODONTIA SIMPLES DE PERMANENTE", "requer_dente": True},
    "5": {"codigo_completo": "82001022", "descricao": "INCISAO E DRENAGEM EXTRA-ORAL", "requer_dente": False},
    "6": {"codigo_completo": "82001030", "descricao": "INCISAO E DRENAGEM INTRA-ORAL", "requer_dente": False},
    "7": {"codigo_completo": "82001197", "descricao": "REDUCAO SIMPLES DE LUXACAO", "requer_dente": False},
    "8": {"codigo_completo": "82001251", "descricao": "REIMPLANTE DENTARIO COM CONTENCAO", "requer_dente": True},
    "9": {"codigo_completo": "82001308", "descricao": "REMOCAO DE DRENO EXTRA-ORAL", "requer_dente": False},
    "10": {"codigo_completo": "82001316", "descricao": "REMOCAO DE DRENO INTRA-ORAL", "requer_dente": False},
    "11": {"codigo_completo": "82001499", "descricao": "SUTURA DE FERIDA EM REGIAO BUCO-MAXILO-FACIAL", "requer_dente": False},
    "12": {"codigo_completo": "82001650", "descricao": "TRATAMENTO DE ALVEOLITE", "requer_dente": True},
    "13": {"codigo_completo": "83000089", "descricao": "EXODONTIA SIMPLES DE DECIDUO", "requer_dente": True},
    "14": {"codigo_completo": "85100048", "descricao": "COLAGEM DE FRAGMENTOS DENTARIOS", "requer_dente": True},
    "15": {"codigo_completo": "85100196", "descricao": "RESTAURACAO EM RESINA FOTOPOLIMERIZAVEL 1 FACE", "requer_dente": True},
    "16": {"codigo_completo": "85100200", "descricao": "RESTAURACAO EM RESINA FOTOPOLIMERIZAVEL 2 FACES", "requer_dente": True},
    "17": {"codigo_completo": "85100218", "descricao": "RESTAURACAO EM RESINA FOTOPOLIMERIZAVEL 3 FACES", "requer_dente": True},
    "18": {"codigo_completo": "85100226", "descricao": "RESTAURACAO EM RESINA FOTOPOLIMERIZAVEL 4 FACES", "requer_dente": True},
    "19": {"codigo_completo": "85200034", "descricao": "PULPECTOMIA", "requer_dente": True},
    "20": {"codigo_completo": "85200085", "descricao": "RESTAURACAO TEMPORARIA / TRATAMENTO EXPECTANTE", "requer_dente": True},
    "21": {"codigo_completo": "85300020", "descricao": "IMOBILIZACAO DENTARIA EM DENTES PERMANENTES", "requer_dente": False},
    "22": {"codigo_completo": "85300063", "descricao": "TRATAMENTO DE ABSCESSO PERIODONTAL AGUDO", "requer_dente": False},
    "23": {"codigo_completo": "85300080", "descricao": "TRATAMENTO DE PERICORONARITE", "requer_dente": True},
    "24": {"codigo_completo": "85400084", "descricao": "COROA PROVISORIA SEM PINO", "requer_dente": True},
    "25": {"codigo_completo": "85400467", "descricao": "RECIMENTACAO DE TRABALHOS PROTETICOS", "requer_dente": True}
}

URL_PORTAL = "https://www.dentistaportoseguro.com.br/"
