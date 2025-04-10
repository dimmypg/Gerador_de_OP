import pyautogui as py
import time

# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar 1 tecla
# pyautogui.click -> clicar em algum lugar da tela
# pyautogui.hotkey -> combinação de teclas
# pyautogui.scroll -> rola o scroll do mouse
py.PAUSE = 0.5

        
'''
Questor
Point(x=587, y=11)

Estoque
Point(x=203, y=33)

Indústria
Point(x=291, y=232)

Ordem de produção
Point(x=460, y=235)

Campo Cliente
Point(x=203, y=207)

'''

def abrir_op():
    py.click(587,11)#Questor
    py.click(203,33)#Estoque
    py.click(291,232)#Indústria
    py.click(460,235)#Ordem de produção
    time.sleep(0.5)#Aguarda
    py.click(203,207)#Campo Cliente
    py.write("7615")#Preenche o código da IPA
    py.press("F6")#Salba
    py.hotkey("ctrl","p")#Abre o produto produzido

def main():
    abrir_op()

main()