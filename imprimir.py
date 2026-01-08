#Pega todos arquivos de uma pasta e imprime.

import os
import time
import pyautogui
import subprocess

# Caminho da pasta dos PDFs
pasta_pdf = r"C:\SEI_PDF_BLOCO_número do bloco_CONSOLIDADO - Copia"   

# Lista todos os PDFs ordenados alfabeticamente
lista_processar = sorted([arq for arq in os.listdir(pasta_pdf) if arq.lower().endswith(".pdf")])

if not lista_processar:
    print("Nenhum PDF encontrado.")
    exit()

for pdf in lista_processar:
    caminho_pdf = os.path.join(pasta_pdf, pdf)
    print(f"Processando: {pdf}")

    # Abre o PDF no Edge
    subprocess.Popen(["cmd", "/c", "start", "", caminho_pdf])

    time.sleep(3)  # tempo para abrir

    # CTRL + P
    pyautogui.hotkey("ctrl", "p")
    time.sleep(2)

    # ENTER para imprimir
    pyautogui.press("enter")
    time.sleep(3)

    # Fecha a aba do navegador
    pyautogui.hotkey("ctrl", "w")
    time.sleep(1)

print("Processo concluído.")
