import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)


# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=419, y=262, clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=419, y=262)
pyautogui.click(x=1168, y=158)
pyautogui.click(x=973, y=605)
time.sleep(2)
pyautogui.click(x=942, y=21)

# Passo 4: Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r"D:\Downloads\Vendas - Dez.xlsx") #Preencher com o caminho do arquivo
display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=86, y=171)
time.sleep(4)
pyautogui.write("emailteste@gmail.com")
pyautogui.press("tab") # seleciona o email
pyautogui.press("tab") # pula pro campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v") # escrever o assunto
pyautogui.press("tab") #pular pro corpo do email
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Gustavo Bortolon"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.click(x=957, y=695) #anexar arquivo
pyautogui.click(x=248, y=194, clicks = 2) #anexa a tabela

# clicar no botão enviar

# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")



#As coordenadas X e Y variam conforma a resolução
#Para descobrir as coordenadas do cursor do mouse usar:
#time.sleep(5) 
#pyautogui.position()
