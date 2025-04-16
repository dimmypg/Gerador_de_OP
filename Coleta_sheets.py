# =============================================================================
# As 36 primeiras linhas NÃO devem ser alteradas
# Início do programa (já fornecido)
import os.path
import time
import ctypes
import pyautogui as py

"""
pyautogui.write -> escrever um texto
pyautogui.press -> apertar 1 tecla
pyautogui.click -> clicar em algum lugar da tela
pyautogui.hotkey -> combinação de teclas
pyautogui.scroll -> rola o scroll do mouse
"""

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


def ler_planilha():
    os.system('cls') # Limpa o terminal

    # Lê os valores da planilha da aba desejada (colunas A até K)
    result = sheet.values().get(
        spreadsheetId=PLANILHA_GERAL,
        range=f"{NOME_ABA}!A1:K"
    ).execute()

    linhas = result.get("values", [])        # Lê os valores da planilha da aba desejada (colunas A até K)
    result = sheet.values().get(
        spreadsheetId=PLANILHA_GERAL,
        range=f"{NOME_ABA}!A1:K"
    ).execute()

    linhas = result.get("values", [])

    return linhas


def verificador(colunas):

    # Preparando variáveis
    codigos = []
    quantidades = []
    nome = []

        # Percorrer as colunas (ignorando cabeçalho)
    for i in range(1, len(colunas)):
        linha = colunas[i]
       
        if (                                                                # Verificadores:
            (len(linha) > 10 and linha[10].strip().upper() == "TRUE") and   # Verifica a caixa na coluna K
            (len(linha) <= 4 or not linha[4].strip()) and                   # Coluna E vazia
            (linha[1].strip()) and                                          # Coluna B preenchida
            (linha[3].strip())                                              # Coluna D preenchida
        ):
            
            codigos.append(linha[1])
            nome.append(linha[2])
            quantidades.append(linha[3])

    return codigos, quantidades, nome


def abrir_op():
    py.click(587,11)#Questor
    for i in range(5):
        py.press("esc")
    py.click(203,33)#Estoque
    py.click(291,232)#Indústria
    py.click(460,235)#Ordem de produção
    time.sleep(1)#Aguarda
    py.click(203,207)#Campo Cliente
    py.write("7615")#Preenche o código da IPA
    py.press("F6")#Salva
    time.sleep(1)#Aguarda
    
    
def preencher_op(produto,quantidade):
    py.click(196,231)
    py.hotkey("ctrl","p")#Abre o produto produzido
    py.write(produto)
    py.press("enter")
    py.write(quantidade)
    py.press("F6")
    py.press("enter")
    py.press("enter")
    py.press("enter")
    py.press("enter")
    py.press("esc")


def impressao():
    py.click(196,231)
    py.hotkey('ctrl','i')
    py.press("enter")
    time.sleep(5)
    py.press('F11')
    time.sleep(2)
    py.press("enter")
    time.sleep(3)
    py.click(196,231)
    for i in range(8):
        py.press("esc")


def main():

    py.PAUSE = 1

    try:   
        colunas = ler_planilha()                    # Carrega a planilha

        codigos, quantidades, nome = verificador(colunas) # Adiciona na lista
        
        total = len(codigos)
        for i, (codigo, quant, nomes) in enumerate(zip(codigos, quantidades, nome), 1):
            print(f"Processando {i}/{total} - Código: {codigo}, Quantidade: {quant}, Nome: {nomes}")

            cod_questor = str(codigo)
            quantidade = str(quant)

            abrir_op()

            preencher_op(cod_questor,quantidade)
            
            impressao()  

    except HttpError as erro:
        print(f"Ocorreu um erro ao acessar a planilha: {erro}")


main()
