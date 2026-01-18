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
    wait = WebDriverWait(driver, 30)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    janela_principal = None
    
    try:
        # 1. Login
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # 2. Acesso ao Módulo TISS
        time.sleep(6)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        time.sleep(5)
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
        janela_principal = driver.current_window_handle

        # 3. Navegação Direta
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(4)

        # 4. Preenchimento via RadInput
        input_id = "ctl00_MainContent_rtbNumeroAtendimento"
        state_id = "ctl00_MainContent_rtbNumeroAtendimento_ClientState"
        entrar_no_frame_do_elemento(driver, input_id)

        client_state = json.dumps({
            "enabled": True, "emptyMessage": "", "validationText": valor_solicitado,
            "valueAsString": valor_solicitado, "lastSetTextBoxValue": valor_solicitado
        })

        driver.execute_script("""
            var el = document.getElementById(arguments[0]);
            var state = document.getElementById(arguments[1]);
            if(el) {
                el.value = arguments[2];
                if(state) state.value = arguments[3];
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
        """, input_id, state_id, valor_solicitado, client_state)

        # 5. Buscar
        btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
        driver.execute_script("arguments[0].click();", btn_buscar)
        
        # 6. Abrir Guia
        time.sleep(4)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rgMasterTable")))
        link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
        driver.execute_script("arguments[0].click();", link_guia)
        time.sleep(3)

        # 7. Exportar PDF Principal
        btn_imprimir_id = "ctl00_MainContent_btnImprimir_input"
        if entrar_no_frame_do_elemento(driver, btn_imprimir_id):
            driver.find_element(By.ID, btn_imprimir_id).click()
            time.sleep(6) 
            for handle in driver.window_handles:
                if handle != janela_principal:
                    driver.switch_to.window(handle)
                    break
            try:
                dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                Select(dropdown).select_by_value("PDF")
                driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                time.sleep(6) # Tempo para download concluir
                driver.close()
            except: pass
            driver.switch_to.window(janela_principal)

        # 8. Exportar Outras Despesas
        entrar_no_frame_do_elemento(driver, "ctl00_MainContent_rbtOutrasDespesas_input")
        try:
            btn_outras = driver.find_element(By.ID, "ctl00_MainContent_rbtOutrasDespesas_input")
            if btn_outras.is_enabled():
                btn_outras.click()
                time.sleep(6)
                for handle in driver.window_handles:
                    if handle != janela_principal:
                        driver.switch_to.window(handle)
                        dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                        Select(dropdown).select_by_value("PDF")
                        driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                        time.sleep(6)
                        driver.close()
                        break
                driver.switch_to.window(janela_principal)
        except: pass

        # 9. Extração de Dados dos Arquivos
        df_final = processar_arquivos_baixados(download_dir, valor_solicitado)
        
        return {"status": "Sucesso", "dados": df_final, "diretorio": download_dir}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
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
                    if os.path.exists("erro_amhptiss.png"):
                        st.image("erro_amhptiss.png", caption="Screenshot do Erro")
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
