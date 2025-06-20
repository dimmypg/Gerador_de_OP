# pyinstaller --onefile --name="GerOP" --icon="automacao.ico" --hidden-import=PIL --hidden-import=cv2 --hidden-import=pyscreeze "GerOP.py"

# =============================================================================
# As 36 primeiras linhas NÃO devem ser alteradas
# Início do programa (já fornecido)
import os.path
import time
import ctypes
import pyautogui as py
import pygetwindow as gw
import pyperclip
import sys
import win32gui

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


def esperar_ate_aparecer(imagem, timeout=10, confidence=0.9, intervalo=0.5):
    """
    Tenta localizar a imagem na tela até o tempo limite (timeout).
    """
    start_time = time.time()
    while True:
        try:
            pos = py.locateOnScreen(imagem, confidence=confidence)
            if pos:
                return pos
        except py.ImageNotFoundException:
            pass  # Não faz nada, apenas tenta de novo

        if time.time() - start_time > timeout:
            raise Exception(f"Imagem '{imagem}' não encontrada após {timeout} segundos.")
            
        time.sleep(intervalo)


def clicar_quando_aparecer(imagem, timeout=5):
    """
    Espera a imagem aparecer e clica nela.
    Se não encontrar dentro do timeout, lança uma exceção.
    """
    pos = esperar_ate_aparecer(imagem, timeout)
    if pos:
        py.click(pos)
    else:
        raise Exception(f"Imagem '{imagem}' não encontrada em {timeout} segundos.")
    

def clicar_dir_quando_aparecer(imagem, timeout=5):
    """
    Espera a imagem aparecer e clica com o clique direito nela.
    Se não encontrar dentro do timeout, lança uma exceção.
    """
    pos = esperar_ate_aparecer(imagem, timeout)
    if pos:
        py.click(pos, button='right')
    else:
        raise Exception(f"Imagem '{imagem}' não encontrada em {timeout} segundos.")


def abrir_op(lista_op):
    clicar_quando_aparecer("Questor.png")
    clicar_quando_aparecer("Estoque.png")
    clicar_quando_aparecer("Industria.png")
    clicar_quando_aparecer("Ordem_de_Producao.png")
    time.sleep(1)#Aguarda
    clicar_dir_quando_aparecer("NumeroOP.png")
    clicar_quando_aparecer("Copiar.png")
    num_op = pyperclip.paste().strip()
    lista_op.append(num_op)
    pyperclip.copy('')
    py.press("tab")
    py.write("7615")#Preenche o código da IPA
    py.press("F6")#Salva
    time.sleep(1)#Aguarda
    
    
def preencher_op(produto,quantidade):
    clicar_quando_aparecer("Questor.png")
    py.hotkey("ctrl","p")#Abre o produto produzido
    esperar_ate_aparecer("Produto.png")
    time.sleep(1)
    py.write(produto)
    py.press("enter")
    py.write(quantidade)
    py.press("F6")
    py.press("enter")
    while True:
        try:
            esperar_ate_aparecer("OK.png",timeout= 2,confidence=0.5)
            py.press('enter')
            time.sleep(0.2)
        except Exception:
            time.sleep(0.5)
            py.press('esc')
            break


def esperar_enquanto_aparecer(imagem, intervalo=0.5, timeout=None, confidence=0.9):
    """
    Espera enquanto uma imagem estiver presente na tela.

    :param imagem: Nome do arquivo da imagem.
    :param intervalo: Tempo (em segundos) entre cada verificação.
    :param timeout: Tempo máximo (em segundos) para aguardar. Se None, espera indefinidamente.
    :param confidence: Precisão da detecção da imagem (padrão 0.9).
    """
    inicio = time.time()

    while True:
        try:
            if not py.locateOnScreen(imagem, confidence=confidence):
                break
        except py.ImageNotFoundException:
            break  # Se não encontrou, sai do loop
        except Exception as e:
            print(f"Erro inesperado durante a busca da imagem '{imagem}': {e}")
            break

        if timeout and (time.time() - inicio) > timeout:
            raise Exception(f"A imagem '{imagem}' ficou tempo demais na tela (>{timeout}s).")

        time.sleep(intervalo)



def impressao():
    clicar_quando_aparecer("Questor.png")
    py.hotkey('ctrl','i')
    clicar_quando_aparecer("Gerar.png")
    esperar_enquanto_aparecer("AguardaGerarOPParaImpressao.png")
    time.sleep(0.5)
    esperar_ate_aparecer("FolhaOPImpressao.png")
    time.sleep(0.5)
    max_tentativas = 2
    for tentativa in range(max_tentativas):
        clicar_quando_aparecer("Imprimir.png")
        try:
            esperar_ate_aparecer("JanelaDeImpressao.png")
            break # Continua o programa se aparecer a janela
        except Exception as e:
            print(f"Tentativa {tentativa + 1} falhou: {e}")
            if tentativa == max_tentativas - 1:
                raise # Se atingir o máximo de tentativas, relança o erro.
            time.sleep(0.5)
    py.press("enter")
    time.sleep(2)


def AbrirPrograma(titulo_parcial):

    janelas = gw.getWindowsWithTitle(titulo_parcial)

    if janelas:
        janela = janelas[0]

        # print(f"Janela encontrada: {janela.title}")

        if janela.isMinimized:
            janela.restore()
        janela.maximize()
        time.sleep(0.5)  # espera meio segundo

        # Tenta usar win32gui para focar a janela
        try:
            win32gui.SetForegroundWindow(janela._hWnd)
            # print("Janela ativada com win32gui.")
        except Exception as e:
            print(f"Erro ao ativar com win32gui: {e}")
            # fallback para pygetwindow
            try:
                janela.activate()
                # print("Janela ativada com pygetwindow.activate().")
            except Exception as e2:
                print(f"Erro ao ativar com pygetwindow: {e2}")

    else:
        print(f'Nenhuma janela encontrada com o nome: {titulo_parcial}')


def fechar_janelas():
    # os.system('cls')
    # print("Fechando janelas...")

    py.PAUSE = 1

    try:
        clicar_quando_aparecer("Questor.png")
    except Exception as e:
        AbrirPrograma('QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46')
        clicar_quando_aparecer("Questor.png")
        time.sleep(0.5)

    clicar_quando_aparecer("Janelas.png")
    clicar_quando_aparecer("FecharTodasasJanelas.png")

    # print("Janelas fechadas")


def main():

    py.PAUSE = 1

    try:   
        colunas = ler_planilha()                    # Carrega a planilha

        codigos, quantidades, nome = verificador(colunas) # Adiciona na lista
        for i, (codigo, quant, nomes) in enumerate(zip(codigos, quantidades, nome), 1):
            
            print(f"OP N°{i} - Cód. Questor: {codigo}, Quantidade: {quant}, Nome: {nomes}")

        confirmador = input(f"\nDeseja continuar com a lista acima?\n\n[Enter] - Sim \n[Espaço + Enter] - Não\n\n")

        if confirmador != "":

            os.system('cls')
            sys.exit()

        os.system('cls')

        total = len(codigos)
      
        lista_op = []

        for i, (codigo, quant, nomes) in enumerate(zip(codigos, quantidades, nome), 1):
            print(f"Gerando OP {i}/{total} - Cód. Questor: {codigo}, Quantidade: {quant}, Nome: {nomes}")

            cod_questor = str(codigo)
            quantidade = str(quant)

            AbrirPrograma('QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46')

            fechar_janelas()

            abrir_op(lista_op)            

            preencher_op(cod_questor,quantidade)
            
            impressao()  

        fechar_janelas()

        os.system('cls')

        codigos, quantidades, nome = verificador(colunas) # Adiciona na lista
        for i, (lista,codigo, quant, nomes) in enumerate(zip(lista_op,codigos, quantidades, nome), 1):
            
            print(f"OP N°{lista} - Cód. Questor: {codigo}, Quantidade: {quant}, Nome: {nomes}")


        encerrador = input(f"\n\nPressione Enter para encerrar...")

        if encerrador != "":

            os.system('cls')
            sys.exit()


        if total == 0:
            print("\nTodas as OP's foram geradas!\n")

    except HttpError as erro:
        print(f"Ocorreu um erro ao acessar a planilha: {erro}")


main()