# def AbrirQuestor():
    
#     # Use parte fixa do nome da janela do Questor
#     titulo_parcial = 'QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46'

#     # Procura todas as janelas com esse título
#     janelas = gw.getWindowsWithTitle(titulo_parcial)

#     if janelas:
#         janela = janelas[0]
#         if janela.isMinimized:
#             janela.restore()
#         janela.maximize()
#         janela.activate()
#     else:
#         print(f'Nenhuma janela encontrada com o nome: {titulo_parcial}')
# QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46 - [Consulta Estoque - SOMENTE LEITURA]
# QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46 - [Impressão - Ordem De Produção Pendente.]
# QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46 - [Ordens de Produção Pendentes]
# QUESTOR EMPRESARIAL 1-MZ RETIFICA DE MOTORES CNPJ: 94.748.894/0001-46



import pygetwindow as gw
import time

# Lista todas as janelas e imprime seus títulos (para você saber o nome exato)
for w in gw.getWindowsWithTitle(''):
    print(w.title)

# Aguarde um pouco para verificar o nome da janela
time.sleep(3)

# Substitua pelo nome (ou parte) do título da janela que deseja restaurar
titulo = 'Bloco de Notas'  # exemplo

# Encontra a primeira janela que contém o título
janelas = gw.getWindowsWithTitle(titulo)

if janelas:
    janela = janelas[0]
    if janela.isMinimized:
        janela.restore()      # Restaura se estiver minimizada
    janela.maximize()         # Maximiza
    janela.activate()         # Traz para frente
else:
    print(f"Nenhuma janela encontrada com o título: {titulo}")
