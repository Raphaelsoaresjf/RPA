#cmd
#pip install mouseinfo
#python
#from mouseinfo import mouseInfo
#mouseInfo()

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('user.xlsx')
usuarios_sheet = workbook['usuarios']

for linha in usuarios_sheet.iter_rows(min_row=2):
    #[Python] [python.teste] [python.meuinc.com.br] [123@Mudar] [123@Mudar] [20] [123]

    #Nome
    pyautogui.click(-1641,588, duration=0.5)
    pyautogui.write(linha[0].value)

    #Login
    pyautogui.click(-1080,576, duration=0.5)
    pyautogui.write(linha[1].value)

    #Email
    pyautogui.click(-1535,708, duration=0.5)
    pyautogui.write(linha[2].value)

    #Perfil
    pyautogui.click(-1096,696, duration=0.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(-1025,794, duration=0.5)

    #Senha
    pyautogui.click(-1475,814, duration=0.5)
    pyautogui.write(linha[4].value)

    #Repetir Senha
    pyautogui.click(-1027,814, duration=0.5)
    pyautogui.write(linha[5].value)

    #Codigo mega
    pyautogui.click(-1501,934, duration=0.5)
    pyautogui.write(str(linha[6].value))