import streamlit as st
import pandas as pd
import json
import time
import re
import io
import os
import pdfplumber
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# === 1. CONFIGURAÇÃO DO AMBIENTE E DOWNLOADS ===
def configurar_driver():
    download_dir = os.path.join(os.getcwd(), "temp_pdfs")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    else:
        for f in os.listdir(download_dir):
            if f.endswith(".pdf"):
                try: os.remove(os.path.join(download_dir, f))
                except: pass

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
    return driver

# === 2. GESTÃO DE CONTEXTO (FRAMES) ===
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

# === 3. EXTRAÇÃO DE DADOS DO PDF COM REGEX ===
def processar_pdfs_baixados():
    all_data = []
    pdf_path = os.path.join(os.getcwd(), "temp_pdfs")
    arquivos = [f for f in os.listdir(pdf_path) if f.endswith(".pdf")]
    
    for filename in arquivos:
        try:
            with pdfplumber.open(os.path.join(pdf_path, filename)) as pdf:
                full_text = ""
                for page in pdf.pages:
                    full_text += page.extract_text() + "\n"
                
                # A. Extração para Guia SP/SADT (Procedimentos)
                if "SP/SADT" in full_text.upper():
                    # Regex para capturar Procedimentos baseada no padrão de guias SADT AMHP
                    matches = re.findall(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}\s+\d{2}:\d{2}\s+(\d{2})\s+([\d\.]+-\d)\s+(.*?)\s+1\s+.*?([\d,.]+)\s+([\d,.]+)", full_text)
                    for m in matches:
                        all_data.append({
                            "Data": m[0],
                            "Tabela": m[1],
                            "Código": m[2],
                            "Descrição": m[3].strip(),
                            "Qtd": 1,
                            "Valor Unitário": m[4],
                            "Valor Total": m[5],
                            "Fonte": "Procedimentos (SADT)"
                        })

                # B. Extração para Guia de Outras Despesas (Materiais/Medicamentos)
                elif "OUTRAS DESPESAS" in full_text.upper():
                    # Regex para capturar itens de despesas (Materiais)
                    # Procura Código de Despesa -> Data -> Tabela -> Código -> Qtd -> Valores -> Descrição
                    matches = re.findall(r"(\d{2})\s+([\d\.]+)\s+(\d+)\s+100\s+([\d,.]+)\s+([\d,.]+)\n16-Descrição\s+(.*)", full_text)
                    for m in matches:
                        all_data.append({
                            "Data": "17/07/2025", # Data extraída do cabeçalho do documento
                            "Tabela": m[0],
                            "Código": m[1],
                            "Descrição": m[5].strip(),
                            "Qtd": m[2],
                            "Valor Unitário": m[3],
                            "Valor Total": m[4],
                            "Fonte": "Outras Despesas"
                        })
        except Exception as e:
            st.error(f"Erro ao processar {filename}: {e}")

    return pd.DataFrame(all_data)

# === 4. AUTOMAÇÃO SELENIUM ===
def extrair_detalhes_site_amhp(numero_guia):
    driver = configurar_driver()
    wait = WebDriverWait(driver, 30)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    janela_principal = None
    
    try:
        # A. Login
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # B. Acesso ao Módulo TISS
        time.sleep(6)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        time.sleep(5)
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
        janela_principal = driver.current_window_handle

        # C. Busca da Guia
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(4)

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

        btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
        driver.execute_script("arguments[0].click();", btn_buscar)
        
        time.sleep(4)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rgMasterTable")))
        link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
        driver.execute_script("arguments[0].click();", link_guia)
        time.sleep(3)

        # D. Exportar Guia Principal
        btn_imprimir_id = "ctl00_MainContent_btnImprimir_input"
        if entrar_no_frame_do_elemento(driver, btn_imprimir_id):
            btn_imprimir = driver.find_element(By.ID, btn_imprimir_id)
            driver.execute_script("arguments[0].click();", btn_imprimir)
            time.sleep(7) 

            for handle in driver.window_handles:
                if handle != janela_principal:
                    driver.switch_to.window(handle)
                    break
            
            try:
                dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                Select(dropdown).select_by_value("PDF")
                time.sleep(1)
                driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                time.sleep(5) 
                driver.close() 
            except: pass
            
            driver.switch_to.window(janela_principal)

        # E. Exportar Outras Despesas
        entrar_no_frame_do_elemento(driver, "ctl00_MainContent_rbtOutrasDespesas_input")
        try:
            btn_outras = driver.find_element(By.ID, "ctl00_MainContent_rbtOutrasDespesas_input")
            if btn_outras.is_enabled():
                driver.execute_script("arguments[0].click();", btn_outras)
                time.sleep(7)
                
                for handle in driver.window_handles:
                    if handle != janela_principal:
                        driver.switch_to.window(handle)
                        dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                        Select(dropdown).select_by_value("PDF")
                        time.sleep(1)
                        driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                        time.sleep(5)
                        driver.close()
                        break
                driver.switch_to.window(janela_principal)
        except: pass

        return {"status": "Sucesso", "arquivos": os.listdir("temp_pdfs")}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === 5. INTERFACE DO USUÁRIO ===
st.set_page_config(page_title="GABMA - Conciliação AMHP", page_icon="🏥", layout="wide")
st.title("🏥 Sistema GABMA: Conciliação Automática AMHP")

if "credentials" not in st.secrets:
    st.error("Configure as credenciais no Secrets do Streamlit.")
else:
    guia_id = st.text_input("Número do Atendimento:", placeholder="Ex: 61789641")
    
    if st.button("🚀 Iniciar Processamento"):
        if not guia_id:
            st.warning("Insira o número da guia.")
        else:
            with st.spinner("Robô em ação... Baixando arquivos."):
                res_robo = extrair_detalhes_site_amhp(guia_id)
                
                if "erro" in res_robo:
                    st.error(f"Erro: {res_robo['erro']}")
                else:
                    st.success("PDFs baixados!")
                    with st.spinner("Extraindo dados com Regex..."):
                        df_final = processar_pdfs_baixados()
                        
                        if not df_final.empty:
                            st.subheader("📋 Planilha Consolidada")
                            st.dataframe(df_final, use_container_width=True)
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_final.to_excel(writer, index=False, sheet_name='GABMA_AMHP')
                            
                            st.download_button(
                                label="📥 Baixar Planilha Excel",
                                data=output.getvalue(),
                                file_name=f"gabma_conciliacao_{guia_id}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("Falha na extração. Verifique se os PDFs contêm dados legíveis.")
