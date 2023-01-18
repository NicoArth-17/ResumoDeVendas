# Passo 1: Entrar no sistema da empresa (link do drive)
# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
# Passo 3: Exportar o relatório (fazer o download)
# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos vendidos)
# Passo 5: Enviar um email para a diretoria

import pyautogui
# pyautogui.click -> Clicar
# pyautogui.write -> Escrever
# pyautogui.press -> Pressionar
# pyautogui.hotkey -> Atalho
# pyautogui.position -> Descobrir a posição em que o mouse vai clicar

import pyperclip
# pyperclip.copy -> Copiar

import time
# time.sleep -> Tempo de espera para continuar executando o código


pyautogui.PAUSE = 1
# Reduzindo a velocidade em que as linhas de código serão aplicados para o computador conseguir aplicá-los

pyautogui.FAILSAFE = True
# Abortar a automação durante a execução levando o mouse para o canto superior esquerdo da tela

# PASSO 1 - Entrar no sistema da empresa (link do drive)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(5)

#PASSO 2 - Navegar até o local do relatório (entrar na pasta Exportar)

# print(pyautogui.position()) -> Pasta exportar
pyautogui.click(x=298, y=291, clicks=2)
time.sleep(3)
# print(pyautogui.position()) -> Arquivo
pyautogui.click(x=298, y=291)
time.sleep(3)

# PASSO 3 - Exportar o relatório (fazer o download)

# print(pyautogui.position()) -> 3 pontinhos
pyautogui.click(x=1152, y=193)
time.sleep(3)
# print(pyautogui.position()) -> Download
pyautogui.click(x=946, y=628)
time.sleep(5)

# PASSO 4 - Calcular os indicadores (faturamento e quantidade de produtos vendidos)

import pandas as pd
# Apelidando a biblioteca 'pandas' como 'pd', ou seja, sempre que chamar um comando poderá usar 'pd' ao invés de 'pandas'

tabela = pd.read_excel(r"C:\Users\mobishopgamer\Documents\Estudo\CursoEmVideo\Python\IntensivãoPython\ProjetoAutomaçao\ResumoDeVendas\Vendas.xlsx")

faturamento = tabela['Valor Final'].sum() # [nome da coluna].somar
quantidade = tabela['Quantidade'].sum()

# PASSO 5 - Enviar um email para a diretoria
    # 5.1 - abrir aba e entrar no email
    # 5.2 - clicar no botão escrever
    # 5.3 - preencher as informações
    # 5.4 - enviar email

# 5.1
# 5.2 (link do email salvo com 'escrever' aberto)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox?compose=new")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)

# 5.3
# print(pyautogui.position()) -> Selecionar destinatario
#pyautogui.click(x=875, y=299)
pyautogui.write('nicolarthur17@hotmail.com')
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy('Relatótio de vendas')
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
texto = f'''
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos vendida foi de: {quantidade:,}

Ats
Nicolas Arthur'''
# f'''... = .format
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

#5.4
pyautogui.hotkey("ctrl","enter")