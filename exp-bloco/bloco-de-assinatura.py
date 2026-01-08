# Baixa documentos do processos dentro de um bloco de assinatura e depois consolida eles.

import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import time
import os
import shutil
from pikepdf import Pdf
import traceback
import pyperclip

# CONFIGURAÇÕES
load_dotenv()
login = os.getenv("LOGIN")
senha = os.getenv("SENHA")

INICIO = "número do processo"
FIM    = "número do processo"

processos_processados = set()
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

# ABRIR NAVEGADOR
driver = webdriver.Edge()
driver.get("https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI&infra_url=L3NlaS8=")
driver.maximize_window()
time.sleep(5)

# LOGIN
pyautogui.write(login, interval=0.1)
time.sleep(2)
pyautogui.press("enter")
time.sleep(3)
pyautogui.press("enter")
pyautogui.write(senha, interval=0.1)
time.sleep(4)

# SELECIONAR ÓRGÃO
select = Select(driver.find_element(By.ID, "selOrgao"))
select.select_by_value("7")
time.sleep(2)
driver.find_element(By.ID, "Acessar").click()
time.sleep(7)

# ACESSAR BLOCOS E ASSINATURA
driver.find_element(By.XPATH, "//span[text()='Blocos']").click()
time.sleep(4)
driver.find_element(By.XPATH, "//span[text()='Assinatura']").click()
time.sleep(5)

# ABRIR BLOCO 
driver.find_element(By.LINK_TEXT, "960915").click()
time.sleep(8)

# CLICA NA PRÓXIMA ABA
pyautogui.click(968, 404)
time.sleep(8)

wait = WebDriverWait(driver, 30)

# FUNÇÃO ABRIR DOCUMENTO 
def abrir_documento(nome_doc):
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for iframe in iframes:
            driver.switch_to.frame(iframe)
            try:
                elemento = driver.find_element(
                    By.XPATH,
                    f"//span[contains(normalize-space(text()), '{nome_doc}')]"
                )
                link = elemento.find_element(By.XPATH, "..")
                driver.execute_script("arguments[0].scrollIntoView(true);", link)
                time.sleep(3)
                link.click()
                driver.switch_to.default_content()
                print(f"✔ Documento aberto: {nome_doc}")
                return True
            except:
                driver.switch_to.default_content()
                continue
        print(f"⚠ Documento não encontrado: {nome_doc}")
        return False
    finally:
        driver.switch_to.default_content()

# FUNÇÃO SALVAR VIA PyAutoGUI (COM SUPORTE A ACENTOS)
def salvar_pdf_com_focus(caminho_completo, click_x=600, click_y=200, retries=3):

    for attempt in range(1, retries + 1):
        try:
            try:
                driver.execute_script("window.focus();")
            except:
                pass
            
            time.sleep(0.5)
            pyautogui.click(click_x, click_y)
            time.sleep(1.2)

            pyautogui.hotkey("ctrl", "s")
            time.sleep(2)

            # SOLUÇÃO PARA ACENTOS: COPIAR E COLAR 
            pyperclip.copy(caminho_completo)
            time.sleep(0.6)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.6)

            pyautogui.press("enter")
            time.sleep(3)

            print(f"✔ Tentativa {attempt}: PDF salvo corretamente em: {caminho_completo}")
            return caminho_completo

        except Exception as e:
            print(f"⚠ Tentativa {attempt} falhou ao salvar PDF: {e}")
            traceback.print_exc()
            time.sleep(2)

    raise RuntimeError("Falha ao salvar PDF após múltiplas tentativas.")

# PDF RECENTE
def pegar_pdf_recente():
    arquivos = [os.path.join(PASTA_DOWNLOADS, f) for f in os.listdir(PASTA_DOWNLOADS) if f.lower().endswith(".pdf")]
    if not arquivos:
        return None
    return max(arquivos, key=os.path.getmtime)

# LOOP PRINCIPAL 
while True:
    try:
        processos = wait.until(
            EC.presence_of_all_elements_located(
                (By.XPATH, "//a[contains(@class,'protocoloAberto') and contains(@class,'aIdProcedimento')]")
            )
        )

        lista_textos = [p.text.strip() for p in processos]

        if INICIO not in lista_textos:
            print(f"❌ Processo inicial '{INICIO}' não encontrado.")
            break
        
        if FIM not in lista_textos:
            print(f"❌ Processo final '{FIM}' não encontrado.")
            break

        idx_inicio = lista_textos.index(INICIO)
        idx_fim    = lista_textos.index(FIM)

        intervalo = processos[idx_fim:idx_inicio + 1] if idx_inicio > idx_fim else processos[idx_inicio:idx_fim + 1]
        intervalo = [p for p in intervalo if p.text not in processos_processados]

        if not intervalo:
            print("✔ Todos os processos no intervalo foram processados.")
            break

        processo_link = intervalo[0]
        texto_proc = processo_link.text.strip()
        processos_processados.add(texto_proc)

        abas_antes = driver.window_handles
        processo_link.click()
        print(f"➡ Abrindo processo: {texto_proc}")

        wait.until(lambda d: len(d.window_handles) > len(abas_antes))
        nova_aba = [a for a in driver.window_handles if a not in abas_antes][0]
        driver.switch_to.window(nova_aba)

        time.sleep(7)

        pasta_base = r"C:\SEI_PDF_BLOCO_960915"
        pasta_processo = os.path.join(pasta_base, texto_proc.replace("/", "-"))
        os.makedirs(pasta_processo, exist_ok=True)

        documentos = [
            "Notificação",
            "Memória de Cálculo",
            "Nota Técnica",
            "Tabela PSS",
            "Tabela Selic",
            "Notícia RFB",
            "Ficha - Financeira",
            "Formulário - Resposta"
        ]

        pdfs_salvos = []

        # PROCESSAMENTO DOS DOCUMENTOS
        for index, doc in enumerate(documentos, start=1):

            abriu = abrir_documento(doc)
            if not abriu:
                time.sleep(5)
                continue

            # NOTIFICAÇÃO
            if doc == "Notificação":
                print("📄 Notificação encontrada — Gerando PDF manualmente...")

                botao_pdf_clicado = False
                iframes = driver.find_elements(By.TAG_NAME, "iframe")

                for iframe in iframes:
                    driver.switch_to.frame(iframe)
                    time.sleep(3)
                    try:
                        btn_pdf = driver.find_element(By.XPATH, "//img[contains(@src,'documento_gerar_pdf')]")
                        driver.execute_script("arguments[0].scrollIntoView(true);", btn_pdf)
                        time.sleep(2)
                        btn_pdf.click()
                        botao_pdf_clicado = True
                        driver.switch_to.default_content()
                        print("✔ Botão Gerar PDF clicado.")
                        break
                    except:
                        driver.switch_to.default_content()
                        continue

                if not botao_pdf_clicado:
                    print("❌ Botão Gerar PDF não encontrado.")
                    continue

                time.sleep(10)

                botao_gerar_clicado = False
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    driver.switch_to.frame(iframe)
                    time.sleep(2)
                    try:
                        botao_gerar = driver.find_element(By.NAME, "btnGerar")
                        driver.execute_script("arguments[0].scrollIntoView(true);", botao_gerar)
                        time.sleep(2)
                        botao_gerar.click()
                        botao_gerar_clicado = True
                        driver.switch_to.default_content()
                        print("✔ Botão GERAR clicado.")
                        break
                    except:
                        driver.switch_to.default_content()
                        continue

                if not botao_gerar_clicado:
                    print("❌ Botão GERAR não encontrado.")
                    continue

                print("⏳ Aguardando download...")
                time.sleep(12)

                pdf_recente = pegar_pdf_recente()
                if pdf_recente:
                    novo_nome = os.path.join(pasta_processo, f"{index:02d} - {doc}.pdf")
                    shutil.move(pdf_recente, novo_nome)
                    pdfs_salvos.append(novo_nome)
                    print(f"✔ PDF movido: {novo_nome}")
                else:
                    print("❌ Nenhum PDF baixado para Notificação.")

                continue

            # DOCUMENTOS NORMAIS
            try:
                print(f"📄 Abrindo PDF de: {doc}")
                time.sleep(6)

                abas_antes_pdf = driver.window_handles
                link_clicado = False

                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    driver.switch_to.frame(iframe)
                    time.sleep(1)
                    try:
                        link_aqui = driver.find_element(By.XPATH, "//a[contains(normalize-space(text()), 'aqui') or contains(@class,'ancoraVisualizacaoArvore')]")
                        driver.execute_script("arguments[0].scrollIntoView(true);", link_aqui)
                        time.sleep(1)
                        link_aqui.click()
                        link_clicado = True
                        driver.switch_to.default_content()
                        print("✔ Link 'aqui' clicado.")
                        break
                    except:
                        driver.switch_to.default_content()
                        continue

                if not link_clicado:
                    print("⚠ Selenium não achou 'aqui' — clicando coordenadas (400,273)")
                    try:
                        driver.execute_script("window.focus();")
                    except:
                        pass
                    time.sleep(0.5)
                    pyautogui.click(400, 273)
                    link_clicado = True
                    time.sleep(0.5)

                time.sleep(4)

                handles = driver.window_handles
                if len(handles) > len(abas_antes_pdf):
                    aba_pdf = handles[-1]
                    driver.switch_to.window(aba_pdf)
                    print("✔ Mudou para aba do PDF.")
                else:
                    aba_pdf = driver.current_window_handle
                    print("⚠ Salvando na aba atual.")

                time.sleep(4)

                nome_pdf = f"{index:02d} - {doc}.pdf"
                caminho_pdf = os.path.join(pasta_processo, nome_pdf)

                salvar_pdf_com_focus(caminho_pdf, click_x=600, click_y=200, retries=3)
                pdfs_salvos.append(caminho_pdf)

                if len(handles) > len(abas_antes_pdf):
                    driver.close()
                    time.sleep(1)
                    driver.switch_to.window(nova_aba)
                else:
                    try:
                        driver.switch_to.window(nova_aba)
                    except:
                        pass

                time.sleep(1)
                print(f"✔ Documento salvo: {caminho_pdf}")

            except Exception as ex_doc:
                print("⚠ Erro ao processar documento:", ex_doc)
                traceback.print_exc()

                try:
                    handles = driver.window_handles
                    for h in handles:
                        if h != nova_aba and h not in abas_antes:
                            try:
                                driver.switch_to.window(h)
                                driver.close()
                            except:
                                pass
                    driver.switch_to.window(nova_aba)
                except:
                    pass

                continue

        # CONSOLIDAR PDFs DO PROCESSO

        pdf_consolidado = os.path.join(
            pasta_processo,
            f"Consolidado - {texto_proc.replace('/', '-')}.pdf"
        )

        pdf_final = Pdf.new()
        for pdf in pdfs_salvos:
            try:
                with Pdf.open(pdf) as pdf_temp:
                    pdf_final.pages.extend(pdf_temp.pages)
            except Exception as e:
                print(f"Erro ao consolidar {pdf}: {e}")

        pdf_final.save(pdf_consolidado)
        pdf_final.close()

        print(f"✔ Consolidado gerado: {pdf_consolidado}")


        # MOVER O CONSOLIDADO PARA C:\SEI_PDF_BLOCO_960915_CONSOLIDADO

        pasta_final = r"C:\SEI_PDF_BLOCO_960915_CONSOLIDADO"
        os.makedirs(pasta_final, exist_ok=True)

        destino_final = os.path.join(
            pasta_final,
            os.path.basename(pdf_consolidado)
        )

        shutil.copy(pdf_consolidado, destino_final)
        print(f"✔ Consolidado movido para pasta final: {destino_final}")

        # ENCERRAR

        driver.close()
        driver.switch_to.window(abas_antes[0])
        time.sleep(2)

    except Exception as erro_final:
        print("Erro no loop principal:", erro_final)
        traceback.print_exc()
        break

print("✔ Rotina finalizada.")

