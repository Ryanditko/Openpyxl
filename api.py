import requests
import time
import datetime
import json
from openpyxl import load_workbook

# Configuração das APIs
GENESYS_URL = "https://api.genesyscloud.com/v2/your_endpoint"
SALESFORCE_URL = "https://your-salesforce-instance.com/services/data/vXX.0/query/?q=YOUR_QUERY"
HEADERS = {
    "Authorization": "Bearer SEU_TOKEN",
    "Content-Type": "application/json"
}

# Caminho da planilha e dos logs
ARQUIVO_EXCEL = "planilha.xlsx"
LOG_FILE = "logs.json"

# Colunas da planilha
COLUNAS = [
    "Data", "Inicio do intervalo", "Fim do intervalo", "Tipo de média", "ID de agente", "Nome do agente",
    "Atendidas", "Tratamento", "Tratamento médio", "Conversação média", "Espera média",
    "TPC média", "Em espera", "Transferidas"
]

# Horário de funcionamento do script
HORARIO_INICIO = datetime.time(8, 0, 0)   # 08:00 (8 da manhã)
HORARIO_FECHAMENTO = datetime.time(20, 0, 0)  # 20:00 (8 da noite)


def buscar_dados_api(url):
    """Faz requisição para a API e retorna os dados."""
    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Erro ao buscar dados da API {url}: {e}")
        return None

def consolidar_dados():
    """Busca e consolida os dados do Genesys e Salesforce."""
    genesys_dados = buscar_dados_api(GENESYS_URL)
    salesforce_dados = buscar_dados_api(SALESFORCE_URL)

    if not genesys_dados or not salesforce_dados:
        print("⚠️ Dados não disponíveis. Tentando novamente mais tarde.")
        return None

    dados_finais = []
    
    for item in genesys_dados["resultados"]:
        data_atual = datetime.datetime.now().strftime("%Y-%m-%d")
        dados_finais.append({
            "Data": data_atual,
            "Inicio do intervalo": item.get("inicio_intervalo", ""),
            "Fim do intervalo": item.get("fim_intervalo", ""),
            "Tipo de média": item.get("tipo_media", ""),
            "ID de agente": item.get("id_agente", ""),
            "Nome do agente": item.get("nome_agente", ""),
            "Atendidas": item.get("atendidas", ""),
            "Tratamento": item.get("tratamento", ""),
            "Tratamento médio": item.get("tratamento_medio", ""),
            "Conversação média": item.get("conversacao_media", ""),
            "Espera média": item.get("espera_media", ""),
            "TPC média": item.get("tpc_media", ""),
            "Em espera": item.get("em_espera", ""),
            "Transferidas": item.get("transferidas", "")
        })

    return dados_finais

def atualizar_planilha(dados):
    """Atualiza a planilha com os dados consolidados."""
    try:
        wb = load_workbook(ARQUIVO_EXCEL)
        ws = wb.active

        colunas_indices = {COLUNAS[i]: i + 1 for i in range(len(COLUNAS))}

        for item in dados:
            nova_linha = [""] * len(COLUNAS)
            for coluna, valor in item.items():
                indice = colunas_indices[coluna] - 1
                nova_linha[indice] = valor
            ws.append(nova_linha)

        wb.save(ARQUIVO_EXCEL)
        print("Planilha atualizada com sucesso!")

    except Exception as e:
        print(f"Erro ao atualizar a planilha: {e}")

def salvar_log(dados):
    """Salva logs das alterações feitas no sistema."""
    try:
        log_data = []
        try:
            with open(LOG_FILE, "r") as f:
                log_data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass  # Arquivo ainda não existe ou está vazio

        log_data.append({
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "alteracoes": dados
        })

        with open(LOG_FILE, "w") as f:
            json.dump(log_data, f, indent=4)

        print("Log salvo com sucesso!")

    except Exception as e:
        print(f"Erro ao salvar log: {e}")

if __name__ == "__main__":

    while True:
        agora = datetime.datetime.now().time()

        # Verifica se é horário de fechamento do expediente
    if HORARIO_INICIO <= agora <= HORARIO_FECHAMENTO:
        print("Consolidando dados do expediente...")
    dados_consolidados = consolidar_dados()
    if dados_consolidados:
        atualizar_planilha(dados_consolidados)
        salvar_log(dados_consolidados)
    print("Aguardando a próxima execução...")
else:
    print("Fora do horário de expediente. Aguardando início...")

