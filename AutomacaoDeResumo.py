from PyPDF2 import PdfReader
import pyautogui
from time import sleep

dados = ""

# lendo o pdf
reader = PdfReader(f'dados//1.pdf')

# quantidades de paginas
qtd_paginas = len(reader.pages)

# extraindo texto de todas as paginas
for p in range(qtd_paginas):
    pagina = reader.pages[p].extract_text()
    dados += f"\n{pagina}\n"

# salvando os textos em um arquivo .txt
with open("dados.txt", "w", encoding="UTF-8") as arquivo:
    arquivo.write(dados)


# PyAutoGui

# Definindo as variaveis  
link = "https://www.bing.com/"
prompt = "Escreva um resumo detalhado em forma de topicos baseado no texto abaixo:\n"
nome_do_arquivo = "Resumo-slides"
caminho = "F:\\workspace\\projetos\\AutomacaoDeResumos\\resumo"

pyautogui.PAUSE = 1

# abrindo e copiando o texto do arquivo txt gerado da extração de textos dos slides 
pyautogui.hotkey('win', 'r')
pyautogui.write('F:\\workspace\\projetos\\AutomacaoDeResumos\\dados.txt')
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('alt', 'f4')

# Acessando o microsoft copilot
pyautogui.hotkey('win', 'r')
pyautogui.write('Chrome')
pyautogui.press('enter')
pyautogui.hotkey('win', 'up')

pyautogui.write(link)
pyautogui.press('enter')

pyautogui.click(x=1004, y=236)
pyautogui.click(x=456, y=124)

pyautogui.press('tab')
pyautogui.write(prompt)
pyautogui.hotkey('ctrl', 'v')

pyautogui.click(x=412, y=658)

sleep(120)

# botão download
pyautogui.click(x=586, y=662)

# botão pdf
pyautogui.click(x=594, y=577)

# botão imprimir
pyautogui.press('enter')

sleep(1)
# nome do arquivo
pyautogui.write(nome_do_arquivo)

# caminho do arquivo
pyautogui.click(x=563, y=300)
pyautogui.write(caminho)
pyautogui.press('enter')

# salvar arquivo
pyautogui.click(x=708, y=697)
pyautogui.hotkey('alt', 'f4')

# abrir resumo
pyautogui.hotkey('win', 'r')
pyautogui.write(f'{caminho}\\{nome_do_arquivo}.pdf')
pyautogui.press('enter')
pyautogui.hotkey('win', 'up')
