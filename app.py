import streamlit as st
import pandas as pd
import json
import time
import re
import io
import os
import shutil
import pdfplumber
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from pytesseract import image_to_string
from pdf2image import convert_from_path

# === CONFIGURAÇÃO DO AMBIENTE ===

from selenium.common.exceptions import TimeoutException, NoSuchElementException

def habilitar_download_headless(driver, download_dir):
    """Libera downloads em headless via CDP (se suportado)."""
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(download_dir)
        })
    except Exception:
        pass

def esperar_pdf_baixar(download_dir, timeout=90):
    """
    Aguarda até existir pelo menos 1 PDF completo (sem .crdownload) no diretório,
    com tamanho estabilizado, para evitar arquivo incompleto.
    """
    t0 = time.time()
    ultimo_total = -1
    while time.time() - t0 < timeout:
        arquivos = os.listdir(download_dir)
        pdfs = [a for a in arquivos if a.lower().endswith(".pdf")]
        crds = [a for a in arquivos if a.lower().endswith(".crdownload")]
        if pdfs and not crds:
            total = sum(os.path.getsize(os.path.join(download_dir, p)) for p in pdfs)
            if total == ultimo_total and total > 0:
                return True
            ultimo_total = total
        time.sleep(1.2)
    raise TimeoutException("Tempo excedido aguardando download do PDF.")

def switch_to_default_and_return(driver):
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def switch_to_frame_contendo(driver, by, locator, max_depth=6):
    """
    Busca recursivamente por iframes até encontrar um elemento (by, locator).
    Mantém o driver DENTRO do frame onde o elemento foi encontrado.
    Retorna True/False.
    """
    switch_to_default_and_return(driver)

    def _search_in_frame(depth=0):
        if depth > max_depth:
            return False
        try:
            driver.find_element(by, locator)
            return True
        except Exception:
            pass

        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for idx in range(len(frames)):
            switch_to_default_and_return(driver)
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            try:
                driver.switch_to.frame(frames[idx])
                if _search_in_frame(depth + 1):
                    return True
            except Exception:
                continue
        return False

    return _search_in_frame()

def mudar_para_contexto_relatorio(driver, wait, janela_principal, janela_sistema):
    """
    Após clicar em 'Imprimir' ou 'Outras Despesas':
      1) Se abrir nova janela/aba -> troca pra ela ('window').
      2) Senão, tenta achar o ReportViewer embutido via iframe ('iframe').
    Retorna: {"modo": "window"|"iframe"|None, "success": bool}
    """
    # 1) tenta nova janela/aba
    try:
        wait.until(lambda d: len(d.window_handles) > 2)
        for h in driver.window_handles:
            if h not in [janela_principal, janela_sistema]:
                driver.switch_to.window(h)
                return {"modo": "window", "success": True}
    except Exception:
        pass

    # 2) tenta localizar ReportViewer embutido via iframe (procurando a toolbar)
    try:
        driver.switch_to.window(janela_sistema)
        xp_toolbar = "//*[self::select or self::button or self::input][contains(@id,'Export') or contains(@title,'Export') or contains(@aria-label,'Export')]"
        achou = switch_to_frame_contendo(driver, By.XPATH, xp_toolbar, max_depth=6)
        if achou:
            return {"modo": "iframe", "success": True}
    except Exception:
        pass

    return {"modo": None, "success": False}

def exportar_pdf_reportviewer_generico(driver, wait, download_dir):
    """
    Fallback genérico para o ReportViewer:
      - Seleciona 'PDF' no dropdown (ids/titles genéricos)
      - Clica no 'Export' (botão genérico)
      - Aguarda download
    """
    dropdown_xpaths = [
        "//select[contains(@id,'Export') or contains(@title,'Export') or contains(@aria-label,'Export')]",
        "//select[contains(@id,'FormatList') or contains(@name,'FormatList')]"
    ]
    export_btn_xpaths = [
        "//*[@id and (contains(@id,'Export') or contains(@title,'Export')) and (self::button or self::input)]",
        "//input[@type='submit' and (contains(@id,'Export') or contains(@title,'Export'))]"
    ]

    dropdown = None
    for xp in dropdown_xpaths:
        try:
            dropdown = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            break
        except Exception:
            continue
    if dropdown is None:
        raise NoSuchElementException("Dropdown de formato do ReportViewer não encontrado.")

    # Seleciona PDF pelo texto visível ou pelo value
    try:
        Select(dropdown).select_by_visible_text("PDF")
    except Exception:
        try:
            Select(dropdown).select_by_value("PDF")
        except Exception:
            try:
                Select(dropdown).select_by_visible_text("Pdf")
            except Exception:
                pass

    export_btn = None
    for xp in export_btn_xpaths:
        try:
            export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            break
        except Exception:
            continue
    if export_btn is None:
        raise NoSuchElementException("Botão Export do ReportViewer não encontrado.")

    driver.execute_script("arguments[0].click();", export_btn)
    esperar_pdf_baixar(download_dir, timeout=90)


def configurar_driver():
    download_dir = os.path.join(os.getcwd(), "temp_pdfs")
    # Limpeza preventiva para teste limpo
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    os.makedirs(download_dir)

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    opts.add_experimental_option("prefs", prefs)
    
    chrome_bin = os.environ.get("CHROME_BINARY", "/usr/bin/chromium")
    if os.path.exists(chrome_bin):
        opts.binary_location = chrome_bin

    try:
        driver = webdriver.Chrome(options=opts)
    except:
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)
    return driver, download_dir

# === NAVEGAÇÃO ENTRE FRAMES (SUA LÓGICA ORIGINAL) ===

def entrar_no_frame_do_elemento(driver, element_id):
    driver.switch_to.default_content()
    try:
        driver.find_element(By.ID, element_id)
        return True 
    except:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for i, frame in enumerate(iframes):
            driver.switch_to.default_content()
            driver.switch_to.frame(i)
            try:
                driver.find_element(By.ID, element_id)
                return True
            except:
                continue
    return False

# === MOTOR DE EXTRAÇÃO (INTELIGÊNCIA GABMA) ===

def extrair_texto_pdf(caminho_pdf):
    texto_full = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: texto_full += t + "\n"
    except Exception as e:
        st.error(f"Erro ao ler PDF nativo: {e}")
    
    # Se o texto for nulo ou imagem (comum no AMHP), usa OCR
    if len(texto_full.strip()) < 50:
        try:
            paginas_img = convert_from_path(caminho_pdf, dpi=200)
            for img in paginas_img:
                texto_full += image_to_string(img, lang='por') + "\n"
        except Exception as e:
            st.error(f"Erro no OCR (verifique packages.txt): {e}")
    
    return texto_full

def processar_arquivos_baixados(diretorio, numero_guia):
    dados_lista = []
    # Regex flexível para capturar dados de faturamento
    padrao = re.compile(
        r"(\d{2}/\d{2}/\d{4})"  # Data
        r".*?"                  # Salto preguiçoso
        r"(\d[\d\.\-]{5,15})"   # Código TUSS
        r"\s+(.*?)\s+"          # Descrição
        r"(\d+)\s+"             # Qtd
        r"([\d,.]+)\s+"         # Unit
        r"([\d,.]+)",           # Total
        re.DOTALL
    )
    
    for arquivo in os.listdir(diretorio):
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(diretorio, arquivo)
            texto = extrair_texto_pdf(caminho)
            texto_limpo = re.sub(r"[ \t]+", " ", texto) # Normaliza espaços
            matches = padrao.findall(texto_limpo)
            
            for m in matches:
                dados_lista.append({
                    "Guia": numero_guia,
                    "Data": m[0],
                    "Código": m[1],
                    "Descrição": m[2].replace("\n", " ").strip(),
                    "Qtd": m[3],
                    "Valor Unit": m[4],
                    "Valor Total": m[5],
                    "Arquivo Origem": arquivo
                })
    return pd.DataFrame(dados_lista)

# === FUNÇÃO PRINCIPAL DE BUSCA ===

def extrair_detalhes_site_amhp(numero_guia):
    driver, download_dir = configurar_driver()
    # Garantir caminho absoluto para o Chrome
    download_dir = os.path.abspath(download_dir) 
    wait = WebDriverWait(driver, 30)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    janela_principal = driver.current_window_handle
    
    try:
        # 1. Login (Mantido)
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # 2. Transição para AMHPTISS
        time.sleep(7)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        # Esperar nova janela abrir e focar nela
        wait.until(lambda d: len(d.window_handles) > 1)
        for handle in driver.window_handles:
            if handle != janela_principal:
                driver.switch_to.window(handle)
                break
        
        janela_sistema = driver.current_window_handle

        # 3. Busca (Navegação Direta)
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(5)

        # Preenchimento Robusto
        input_atendimento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_rtbNumeroAtendimento")))
        driver.execute_script(f"arguments[0].value = '{valor_solicitado}';", input_atendimento)
        
        btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
        driver.execute_script("arguments[0].click();", btn_buscar)
        
        # 4. Abrir Relatório
        time.sleep(5)
        link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
        driver.execute_script("arguments[0].click();", link_guia)
        
       
        # 5. O PULO DO GATO: Download robusto (somente melhorias; passos 1–4 intactos)
        # Libera download em headless via CDP (não impacta navegação)
        try:
            habilitar_download_headless(driver, download_dir)
        except Exception:
            pass

        # Vamos tentar os dois botões (Imprimir e Outras Despesas) — sua ordem original
        botoes = ["ctl00_MainContent_btnImprimir_input", "ctl00_MainContent_rbtOutrasDespesas_input"]
        
        for id_btn in botoes:
            driver.switch_to.window(janela_sistema)
            if entrar_no_frame_do_elemento(driver, id_btn):
                try:
                    btn_export = driver.find_element(By.ID, id_btn)
                    if btn_export.is_enabled():
                        # Clica no botão (mesma lógica)
                        driver.execute_script("arguments[0].click();", btn_export)
                        
                        # === CAMINHO ORIGINAL: nova janela do relatório ===
                        tentou_popup = False
                        try:
                            wait.until(lambda d: len(d.window_handles) > 2)
                            tentou_popup = True

                            # Vai para a janela do relatório
                            for handle in driver.window_handles:
                                if handle not in [janela_principal, janela_sistema]:
                                    driver.switch_to.window(handle)
                                    break
                            
                            # Tenta primeiro pelos IDs fixos que você usava
                            try:
                                drop = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                                Select(drop).select_by_value("PDF")
                                time.sleep(2)
                                btn_final = driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export")
                                driver.execute_script("arguments[0].click();", btn_final)
                            except Exception:
                                # Se IDs mudarem, usa o fallback genérico
                                exportar_pdf_reportviewer_generico(driver, wait, download_dir)

                            # Espera o PDF finalizar o download (robusto). Se der timeout, aplica seu sleep como reserva.
                            try:
                                esperar_pdf_baixar(download_dir, timeout=90)
                            except Exception:
                                time.sleep(8)

                            # Fecha a janela do relatório e volta
                            driver.close()
                            driver.switch_to.window(janela_sistema)

                        except Exception:
                            # === FALLBACK: relatório pode ter carregado em iframe na mesma aba ===
                            if not tentou_popup:
                                # Se nem abriu popup, tenta localizar a toolbar dentro de um iframe
                                try:
                                    ctx = mudar_para_contexto_relatorio(driver, wait, janela_principal, janela_sistema)
                                    if ctx["success"]:
                                        exportar_pdf_reportviewer_generico(driver, wait, download_dir)
                                        if ctx["modo"] == "window":
                                            driver.close()
                                            driver.switch_to.window(janela_sistema)
                                        else:
                                            driver.switch_to.default_content()
                                    else:
                                        st.write(f"Aviso: Cliquei em {id_btn}, mas não localizei relatório (popup/iframe).")
                                except Exception as e2:
                                    st.write(f"Aviso: falha no fallback do relatório: {e2}")

                except Exception as e:
                    st.write(f"Aviso: Falha ao tentar clicar em {id_btn}: {e}")
                    try:
                        driver.save_screenshot("erro_download.png")
                    except Exception:
                        pass
                    continue

        driver.switch_to.window(janela_sistema)
        
        # 6. Extração
        df_final = processar_arquivos_baixados(download_dir, valor_solicitado)
        return {"status": "Sucesso", "dados": df_final, "diretorio": download_dir}

    except Exception as e:
        driver.save_screenshot("erro_download.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === INTERFACE STREAMLIT ===

st.set_page_config(page_title="GABMA - Consulta AMHP", page_icon="🏥", layout="wide")
st.title("🏥 Inteligência de Faturamento AMHP")

if "credentials" not in st.secrets:
    st.error("Configure as credenciais em Secrets.")
else:
    guia = st.text_input("Número do Atendimento:")
    
    if st.button("🚀 Processar e Analisar"):
        if not guia:
            st.warning("Informe a guia.")
        else:
            with st.spinner("Navegando no portal e baixando documentos..."):
                res = extrair_detalhes_site_amhp(guia)
                
                
                if "erro" in res:
                    st.error(f"Erro: {res['erro']}")
                    if os.path.exists("erro_download.png"):
                        st.image("erro_download.png", caption="Screenshot do Erro")
                else:
                    st.success("Automação concluída!")
                    
                    # --- TESTE DE DOWNLOAD (Para você conferir se baixou) ---
                    with st.expander("📂 Conferência de Arquivos Baixados"):
                        arquivos = os.listdir(res["diretorio"])
                        if arquivos:
                            for arq in arquivos:
                                caminho = os.path.join(res["diretorio"], arq)
                                tamanho = os.path.getsize(caminho) / 1024
                                st.write(f"📄 {arq} ({tamanho:.1f} KB)")
                                with open(caminho, "rb") as f:
                                    st.download_button(f"📥 Baixar {arq}", f, file_name=arq)
                        else:
                            st.warning("Nenhum arquivo encontrado na pasta de download.")

                    # --- EXIBIÇÃO DOS DADOS ---
                    df = res["dados"]
                    if not df.empty:
                        st.subheader("📋 Dados Extraídos")
                        st.dataframe(df, use_container_width=True)
                        csv = df.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 Baixar Planilha de Resultados", csv, "faturamento.csv", "text/csv")
                    else:
                        st.info("Os arquivos foram baixados, mas o motor de extração não encontrou o padrão de faturamento (verifique a Regex ou se é imagem).")
