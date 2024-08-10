import pyautogui
import time

controlador = 1

quantidade = int(input('quantidade de NF: ')) * 2
tempo = int(input('Segundos de um comando para outro: '))

time.sleep(5)

while (controlador <= quantidade):
    pyautogui.click(-1580, 880) 
    pyautogui.press('tab')
    pyautogui.press('space')
    time.sleep(tempo)
    controlador = controlador + 1