# Acessa um processo especifico, acessa o documento expedição e consolida os arquivos que ten link no documento. PARA OFICIOS

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
driver.get("https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI&infra_url=L3NlaS8=")
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
        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for iframe in iframes:
                driver.switch_to.frame(iframe)
                try:
                    span = driver.find_element(By.XPATH, f"//span[contains(normalize-space(text()), '{texto_alvo}')]")
                    link = span.find_element(By.XPATH, "..")
                    driver.execute_script("arguments[0].scrollIntoView(true);", link)
                    driver.execute_script("arguments[0].click();", link)
                    driver.switch_to.default_content()
                    print(f"✔ '{texto_alvo}' clicado com sucesso")
                    return True
                except:
                    driver.switch_to.default_content()
                    continue
            time.sleep(delay)
        except:
            driver.switch_to.default_content()
            time.sleep(delay)
    print(f"⚠ Não foi possível clicar em '{texto_alvo}' após {max_tries} tentativas")
    return False


try:
    # LISTA DE PROCESSOS
    processos = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'acao=procedimento_trabalhar')]"))
    )

    processo_link = next((p for p in processos if processo_desejado in p.text.strip()), None)

    if not processo_link:
        print(f"⚠ Processo {processo_desejado} não encontrado.")
    else:
        print(f"✔ Abrindo processo {processo_desejado}")

        abas_antes = driver.window_handles
        driver.execute_script("window.open(arguments[0].href, '_blank');", processo_link)
        wait.until(lambda d: len(d.window_handles) > len(abas_antes))
        nova_aba = [a for a in driver.window_handles if a not in abas_antes][0]
        driver.switch_to.window(nova_aba)
        time.sleep(5)

        time.sleep(2)
        pyautogui.click(76, 311)
        time.sleep(2)

        # ABRIR FORMULÁRIO
        if clicar_documento("Formulário de", driver):
            time.sleep(7)
            print("📄 Formulário aberto, tentando acessar conteúdo...")

            try:
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@id, 'ifrVisualizacao')]")))
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@id, 'ifrArvoreHtml')]")))

                links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.ancora_sei")))
                print(f"🔗 {len(links)} links encontrados dentro do Formulário.")

                aba_formulario = driver.current_window_handle
                pdfs_gerados = []

                # LOOP DOS LINKS
                for i, link in enumerate(links, start=1):
                    href = link.get_attribute("href")
                    texto = link.text.strip() or f"Link_{i}"

                    print(f"➡ ({i}/{len(links)}) Abrindo {texto}...")

                    driver.execute_script("window.open(arguments[0], '_blank');", href)
                    time.sleep(2)

                    abas_atuais = driver.window_handles
                    driver.switch_to.window(abas_atuais[-1])
                    time.sleep(5)

                    nome_pdf = os.path.join(PASTA_SAIDA, f"{i:02d}_{texto.replace(' ', '_')}.pdf")

                    # DETECTAR SE É PDF NATIVO
                    is_pdf = False
                    try:
                        driver.find_element(By.XPATH, "//embed[contains(@type,'pdf')]")
                        is_pdf = True
                    except:
                        try:
                            driver.find_element(By.XPATH, "//iframe[contains(@src,'pdf')]")
                            is_pdf = True
                        except:
                            is_pdf = False

                    # MODO HTML OU PDF NATIVO
                    if not is_pdf:
                        # MODO HTML → printToPDF
                        print(f"🖨 Gerando PDF por impressão → {nome_pdf}")

                        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                        pdf_bytes = base64.b64decode(pdf_data["data"])

                        with open(nome_pdf, "wb") as f:
                            f.write(pdf_bytes)

                    else:
                        # MODO PDF NATIVO → PyAutoGUI
                        print(f"💾 Salvando PDF nativo via PyAutoGUI → {nome_pdf}")

                        time.sleep(2)
                        pyautogui.hotkey("ctrl", "s")
                        time.sleep(2)

                        caminho_completo = nome_pdf
                        pyautogui.write(caminho_completo, interval=0.05)
                        time.sleep(1)

                        pyautogui.press("enter")
                        time.sleep(3)

                    pdfs_gerados.append(nome_pdf)

                    driver.close()
                    driver.switch_to.window(aba_formulario)
                    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[contains(@id, 'ifrVisualizacao')]"))
                    driver.switch_to.frame(driver.find_element(By.XPATH, "//iframe[contains(@id, 'ifrArvoreHtml')]"))

                # CONSOLIDAR PDFs
                consolidado = os.path.join(PASTA_SAIDA, f"{processo_desejado.replace('/', '-')}.pdf")
                print(f"📚 Consolidando em: {consolidado}")

                pdf_final = pikepdf.Pdf.new()
                for caminho in pdfs_gerados:
                    try:
                        with pikepdf.open(caminho) as pdf_temp:
                            pdf_final.pages.extend(pdf_temp.pages)
                    except Exception as e:
                        print(f"Erro ao anexar {caminho}: {e}")

                pdf_final.save(consolidado)
                pdf_final.close()

                print("✅ Todos os PDFs foram exportados e consolidados com sucesso.")

            except Exception as e:
                print(f"⚠ Erro ao processar os links do Formulário: {e}")
                driver.switch_to.default_content()

        driver.close()
        driver.switch_to.window(abas_antes[0])
        time.sleep(2)

except Exception as e:
    print(f"⚠ Erro geral: {e}")

driver.quit()
