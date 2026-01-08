# Acessa cada processo da lista em ordem, acessa o documento expedição e consolida os arquivos que ten link no documento.

import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import os

# --- Configurações ---
load_dotenv()
login = os.getenv("LOGIN")
senha = os.getenv("SENHA")

# --- Abrir navegador ---
driver = webdriver.Edge()
driver.get("https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI&infra_url=L3NlaS8=")
driver.maximize_window()
time.sleep(3)

# --- Login com PyAutoGUI ---
pyautogui.write(login, interval=0.1)
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
pyautogui.press("enter")
pyautogui.write(senha, interval=0.1)
time.sleep(2)

# --- Selecionar órgão ---
encontrar_orgao = driver.find_element(By.ID, "selOrgao")
select = Select(encontrar_orgao)
select.select_by_value("7")
time.sleep(2)

# --- Clicar em acessar ---
driver.find_element(By.ID, "Acessar").click()
time.sleep(5)

wait = WebDriverWait(driver, 15)
processos_processados = set()  # IDs de processos já abertos

# --- Loop de processos ---
while True:
    try:
        # Espera e pega todos os links de processos (independente do status)
        processos = wait.until(
            EC.presence_of_all_elements_located((
                By.XPATH,
                "//a[contains(@href, 'acao=procedimento_trabalhar')]"
            ))
        )

        # Filtra o primeiro processo ainda não processado
        processo_link = None
        for p in processos:
            texto = p.text.strip()
            if texto and texto not in processos_processados:
                processo_link = p
                break

        if not processo_link:
            print("⚠ Nenhum novo processo encontrado. Encerrando loop.")
            break

        texto = processo_link.text.strip()
        processos_processados.add(texto)
        print(f"✔ Abrindo processo {texto}")

        # --- Salva abas atuais
        abas_antes = driver.window_handles

        # Abre o processo em nova aba via JavaScript
        driver.execute_script("window.open(arguments[0].href, '_blank');", processo_link)

        # Espera a nova aba aparecer
        wait.until(lambda d: len(d.window_handles) > len(abas_antes))
        abas_apos = driver.window_handles
        aba_nova = [aba for aba in abas_apos if aba not in abas_antes][0]

        # Troca para a aba do processo
        driver.switch_to.window(aba_nova)
        time.sleep(5)

        # --- Interagir com documentos ---
        time.sleep(5)

        # --- Fecha aba do processo ---
        driver.close()

        # --- Volta para aba da lista ---
        driver.switch_to.window(abas_antes[0])
        time.sleep(2)

    except Exception as e:
        print(f"⚠ Erro ao processar processo: {e}")
        break

# --- Fecha navegador ---
driver.quit()
