import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

#Acessando o navegador
pyautogui.hotkey("win")
pyautogui.write("Chrome")
pyautogui.press("enter")
pyautogui.hotkey("ctrl", "t")
#Acessando o Link
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)
#Navegando no sistema
pyautogui.click(x=289, y=256, clicks=2) #Cick duplo na pasta
time.sleep(2)
pyautogui.click(x=362, y=362) #Click no arquivo
pyautogui.click(x=1711, y=155) #Click em "mais opçoes"
pyautogui.click(x=1465, y=560) #Click em download

time.sleep(5)
#Calculando indicadores
tabela = pd.read_excel(r"D:\Users\raulv\Downloads\Vendas - Dez.xlsx")
#print(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#Entrando no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)
#Enviando email
pyautogui.click(x=80, y=175) #Click em compose
pyautogui.write("raulvinicius245+PYTHON@gmail.com")
pyautogui.press("tab") #Seleciona o email
pyautogui.press("tab") #Muda para o campo assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v") #Escreve assunto
pyautogui.press("tab") #Muda para o corpo do email

texto = f"""
Prezados, 

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,} unidades

Segue a tabela em anexo.

Att,
Raul Correia."""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=475, y=475)
pyautogui.write("Vendas - Dez.xlsx")
pyautogui.press("enter")
time.sleep(3)
pyautogui.hotkey("ctrl", "enter")#Envia email
