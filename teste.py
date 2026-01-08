# Documento para testar a

import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import os
import pikepdf
import base64

# CONFIGURAÇÕES
load_dotenv()
login = os.getenv("LOGIN")
senha = os.getenv("SENHA")

processo_desejado = "número do processo"
PASTA_SAIDA = r"C:\PDFs_SEI_número do processo"
os.makedirs(PASTA_SAIDA, exist_ok=True)

# ABRIR NAVEGADOR
driver = webdriver.Edge()
driver.get(
    "https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php"
    "?sigla_orgao_sistema=MGI&sigla_sistema=SEI&infra_url=L3NlaS8="
)
driver.maximize_window()
time.sleep(3)

# LOGIN COM PYAUTOGUI
pyautogui.write(login, interval=0.1)
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
pyautogui.press("enter")
pyautogui.write(senha, interval=0.1)
time.sleep(2)

# SELECIONAR ÓRGÃO
encontrar_orgao = driver.find_element(By.ID, "selOrgao")
select = Select(encontrar_orgao)
select.select_by_value("7")
time.sleep(2)
driver.find_element(By.ID, "Acessar").click()
time.sleep(5)

wait = WebDriverWait(driver, 25)

# FUNÇÃO: CLICAR DOCUMENTO
def clicar_documento(texto_alvo, driver, max_tries=10, delay=1):
    for _ in range(max_tries):
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for iframe in iframes:
            driver.switch_to.frame(iframe)
            try:
                span = driver.find_element(
                    By.XPATH,
                    f"//span[contains(normalize-space(text()), '{texto_alvo}')]"
                )
                link = span.find_element(By.XPATH, "..")
                driver.execute_script("arguments[0].scrollIntoView(true);", link)
                driver.execute_script("arguments[0].click();", link)
                driver.switch_to.default_content()
                return True
            except:
                driver.switch_to.default_content()
        time.sleep(delay)
    return False

try:
    processos = wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//a[contains(@href, 'acao=procedimento_trabalhar')]")
        )
    )

    processo_link = next(
        (p for p in processos if processo_desejado in p.text.strip()),
        None
    )

    if not processo_link:
        print(f"⚠ Processo {processo_desejado} não encontrado.")
    else:
        abas_antes = driver.window_handles
        driver.execute_script(
            "window.open(arguments[0].href, '_blank');",
            processo_link
        )
        wait.until(lambda d: len(d.window_handles) > len(abas_antes))
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(5)

        if clicar_documento("Formulário de", driver):
            time.sleep(5)

            wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.XPATH, "//iframe[contains(@id, 'ifrVisualizacao')]")
                )
            )
            wait.until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.XPATH, "//iframe[contains(@id, 'ifrArvoreHtml')]")
                )
            )

            links = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "a.ancora_sei")
                )
            )

            aba_formulario = driver.current_window_handle
            pdfs_gerados = []

            for i, link in enumerate(links, start=1):
                href = link.get_attribute("href")
                texto = link.text.strip() or f"Documento_{i}"

                driver.execute_script(
                    "window.open(arguments[0], '_blank');",
                    href
                )
                time.sleep(3)
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(5)

                nome_pdf = os.path.join(
                    PASTA_SAIDA,
                    f"{i:02d}_{texto.replace(' ', '_')}.pdf"
                )

                print(f"🖨 Gerando PDF (HTML) → {nome_pdf}")

                pdf_data = driver.execute_cdp_cmd(
                    "Page.printToPDF",
                    {"printBackground": True}
                )
                pdf_bytes = base64.b64decode(pdf_data["data"])

                with open(nome_pdf, "wb") as f:
                    f.write(pdf_bytes)

                pdfs_gerados.append(nome_pdf)

                driver.close()
                driver.switch_to.window(aba_formulario)
                driver.switch_to.frame(
                    driver.find_element(By.XPATH, "//iframe[contains(@id, 'ifrVisualizacao')]")
                )
                driver.switch_to.frame(
                    driver.find_element(By.XPATH, "//iframe[contains(@id, 'ifrArvoreHtml')]")
                )

            consolidado = os.path.join(
                PASTA_SAIDA,
                f"{processo_desejado.replace('/', '-')}.pdf"
            )

            pdf_final = pikepdf.Pdf.new()
            for pdf in pdfs_gerados:
                with pikepdf.open(pdf) as temp:
                    pdf_final.pages.extend(temp.pages)

            pdf_final.save(consolidado)
            pdf_final.close()

            print("✅ PDFs HTML exportados e consolidados com sucesso.")

except Exception as e:
    print(f"⚠ Erro geral: {e}")

driver.quit()

