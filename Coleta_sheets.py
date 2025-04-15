# =============================================================================
# As 36 primeiras linhas NÃO devem ser alteradas
# Início do programa (já fornecido)
import os.path
import time
import ctypes

from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configuração do Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
PLANILHA_GERAL = "1DwA7MRS0QAGTLsc1-grEEHHzZ_SBsFXh60zz1fSpiLk"

credenciais = None

# Verifica se o arquivo de token existe
if os.path.exists("token.json"):
    credenciais = Credentials.from_authorized_user_file("token.json", SCOPES)

# Se as credenciais não forem válidas, tenta renovar ou pedir uma nova autenticação
if not credenciais or not credenciais.valid:
    if credenciais and credenciais.expired and credenciais.refresh_token:
        try:
            credenciais.refresh(Request())
        except Exception as e:
            print(f"Erro ao atualizar o token: {e}")
            credenciais = None  # Forçar nova autenticação

    if not credenciais or not credenciais.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            "client_secret.json", SCOPES
        )
        credenciais = flow.run_local_server(port=0)

        # Salva o novo token para evitar autenticação toda vez
        with open("token.json", "w") as token:
            token.write(credenciais.to_json())

# =============================================================================
# A partir daqui, a lógica continua com a API oficial (googleapiclient)

# Conectar ao Google Sheets API
service = build("sheets", "v4", credentials=credenciais)
sheet = service.spreadsheets()

NOME_ABA = "Solicitação de Produção"

def coleta():
    try:
        # Preparando variáveis
        codigos = []
        quantidades = []


        # Lê os valores da planilha da aba desejada (colunas A até K)
        result = sheet.values().get(
            spreadsheetId=PLANILHA_GERAL,
            range=f"{NOME_ABA}!A1:K"
        ).execute()

        linhas = result.get("values", [])

        # Percorrer as linhas (ignorando cabeçalho)
        for i in range(1, len(linhas)):
            linha = linhas[i]

            # Verifica se a coluna K (índice 10) contém "TRUE"
            if (
                (len(linha) > 10 and linha[10].strip().upper() == "TRUE") and # Verifica a caixa 
                (len(linha) <= 4 or not linha[4].strip())  # Coluna E vazia
            ):
                Questor = linha[1] if len(linha) > 1 else ""
                QNT = linha[3] if len(linha) > 3 else ""

                codigos.append(Questor)
                quantidades.append(QNT)

        return codigos, quantidades

    except HttpError as erro:
        print(f"Ocorreu um erro ao acessar a planilha: {erro}")

codigos, quantidades = coleta()
print(f'Códigos {codigos}')
print(f'Quantidades {quantidades}')