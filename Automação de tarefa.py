import pandas 
import pyautogui
import time
import pyperclip

pyautogui.PAUSE = 3

#Abre o Navegador
pyautogui.press("win")
time.sleep(2)
pyautogui.write("crome")
pyautogui.press("enter")


#Abre o Drive
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("Ctrl", "v")
pyautogui.press("enter")

#Seleciona e baixa a Planilha
pyautogui.click(x=369, y=273, clicks= 2)
pyautogui.click(x=369, y=273) 
pyautogui.click(x=1392, y=154)
pyautogui.click(x=1201, y=561)


#Le a tabela do Excel e atribui colunas á variaveis
tabela = pandas.read_excel(r"C:\Users\Henrique\Downloads\Vendas - Dez.xlsx")
print(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)


#Abre uma nova guia e entre no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#Coloca dados e envia Email ao destinatario selecionado
pyautogui.click(x=218, y=168)
pyperclip.copy("HenriqueSericovM@outlook.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

#Corpo do Email com informações do faturamento e quantidade da planilha do Excell
texto = f"""
Prezados, Bom dia!

O faturamento de ontem foi de R${faturamento:,.2f}
A quantidade de vendas foi de R${quantidade:,.2f}

Abs
Henrique Sericov"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.click(x=690, y=693)
pyautogui.click(x=147, y=337)
pyautogui.click(x=428, y=340)




