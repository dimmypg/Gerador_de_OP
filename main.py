import pyautogui as py
import time

# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar 1 tecla
# pyautogui.click -> clicar em algum lugar da tela
# pyautogui.hotkey -> combinação de teclas
# pyautogui.scroll -> rola o scroll do mouse
py.PAUSE = 0.5

        
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
    py.press("F6")#Salba
    py.hotkey("ctrl","p")#Abre o produto produzido


def preencher_op(produto,quantidade):
    py.click(196,231)
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
    time.sleep(1)
    py.press("enter")
    time.sleep(1)
    for i in range(5):
        py.press("esc")


def main():
    cod_questor = str(input("Informe o código Questor da peça produzida: "))
    quantidade = str(input("Informe quantidade de peças produzidas: "))

    abrir_op()

    preencher_op(cod_questor,quantidade)
    
    impressao()

main()