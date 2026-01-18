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

def configurar_driver():
    # Usamos caminho absoluto para evitar problemas no Linux/Streamlit Cloud
    download_dir = os.path.abspath(os.path.join(os.getcwd(), "temp_pdfs"))
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
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0 # Libera popups de download
    }
    opts.add_experimental_option("prefs", prefs)
    
    chrome_bin = os.environ.get("CHROME_BINARY", "/usr/bin/chromium")
    if os.path.exists(chrome_bin):
        opts.binary_location = chrome_bin

    try:
        driver = webdriver.Chrome(options=opts)
    except:
        driver = webdriver.Chrome(service=Service("/usr/bin/chromedriver"), options=opts)
    
    return driver, download_dir

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
            except: continue
    return False

# === MOTOR DE EXTRAÇÃO (DADOS GABMA) ===

def extrair_texto_pdf(caminho_pdf):
    texto_full = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: texto_full += t + "\n"
    except: pass
    
    if len(texto_full.strip()) < 50:
        try:
            paginas_img = convert_from_path(caminho_pdf, dpi=200)
            for img in paginas_img:
                texto_full += image_to_string(img, lang='por') + "\n"
        except: pass
    return texto_full

def processar_arquivos_baixados(diretorio, numero_guia):
    dados_lista = []
    padrao = re.compile(r"(\d{2}/\d{2}/\d{4}).*?(\d[\d\.\-]{5,15})\s+(.*?)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)", re.DOTALL)
    
    for arquivo in os.listdir(diretorio):
        if arquivo.lower().endswith(".pdf"):
            texto = extrair_texto_pdf(os.path.join(diretorio, arquivo))
            texto_limpo = re.sub(r"[ \t]+", " ", texto)
            matches = padrao.findall(texto_limpo)
            for m in matches:
                dados_lista.append({
                    "Atendimento": numero_guia, "Data": m[0], "Código": m[1],
                    "Descrição": m[2].replace("\n", " ").strip(), "Qtd": m[3],
                    "Vlr Unit": m[4], "Vlr Total": m[5], "Arquivo": arquivo
                })
    return pd.DataFrame(dados_lista)

# === FUNÇÃO DE BUSCA E DOWNLOAD (AMHP) ===

def extrair_detalhes_site_amhp(numero_guia):
    driver, download_dir = configurar_driver()
    wait = WebDriverWait(driver, 30)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    
    try:
        # 1. Login
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # 2. Entrar no AMHPTISS
        time.sleep(6)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        # Mudar para a nova aba do sistema
        wait.until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[-1])
        janela_sistema = driver.current_window_handle

        # 3. Ir para Busca
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(4)
        
        # Preencher Guia
        input_id = "ctl00_MainContent_rtbNumeroAtendimento"
        if entrar_no_frame_do_elemento(driver, input_id):
            el = driver.find_element(By.ID, input_id)
            driver.execute_script(f"arguments[0].value = '{valor_solicitado}';", el)
            driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input").click()

        # 4. Abrir Relatório
        time.sleep(4)
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]"))).click()
        
        # 5. Fluxo de Impressão (Seus Botões)
        time.sleep(3)
        btn_imprimir_id = "ctl00_MainContent_btnImprimir_input"
        if entrar_no_frame_do_elemento(driver, btn_imprimir_id):
            driver.find_element(By.ID, btn_imprimir_id).click()
            
            # Espera abrir a aba do PDF/Relatório
            wait.until(lambda d: len(d.window_handles) > 2)
            driver.switch_to.window(driver.window_handles[-1])
            
            # Selecionar PDF no Dropdown que você passou
            dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
            Select(dropdown).select_by_value("PDF")
            
            # Clicar no link Exportar
            time.sleep(2)
            btn_export = driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export")
            driver.execute_script("arguments[0].click();", btn_export)
            
            # Aguardar o download
            time.sleep(8)
            driver.close() # Fecha aba do relatório
            driver.switch_to.window(janela_sistema)

        # 6. Extração Final
        df_final = processar_arquivos_baixados(download_dir, valor_solicitado)
        return {"status": "Sucesso", "dados": df_final, "diretorio": download_dir}

    except Exception as e:
        driver.save_screenshot("erro_final.png")
        return {"erro": str(e), "dados": pd.DataFrame()} # Garante que a chave 'dados' exista mesmo no erro
    finally:
        driver.quit()

# === INTERFACE ===

st.set_page_config(page_title="GABMA - AMHP", layout="wide")
st.title("🏥 Extrator de Faturamento AMHP")

if "credentials" not in st.secrets:
    st.error("Configure os Secrets.")
else:
    guia = st.text_input("Número da Guia:")
    if st.button("🚀 Iniciar"):
        with st.spinner("Processando..."):
            res = extrair_detalhes_site_amhp(guia)
            
            if "erro" in res and not res.get("status"):
                st.error(f"Erro: {res['erro']}")
                if os.path.exists("erro_final.png"): st.image("erro_final.png")
            else:
                st.success("Concluído!")
                df = res["dados"]
                if not df.empty:
                    st.dataframe(df, use_container_width=True)
                else:
                    st.warning("Nenhum dado extraído. Verifique se o PDF baixou corretamente no expander abaixo.")
                
                with st.expander("📂 Arquivos no Servidor"):
                    arquivos = os.listdir(res["diretorio"])
                    for f in arquivos:
                        st.write(f"📄 {f}")
                        with open(os.path.join(res["diretorio"], f), "rb") as file:
                            st.download_button(f"Download {f}", file, file_name=f)
