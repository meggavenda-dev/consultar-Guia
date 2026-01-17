import streamlit as st
import pandas as pd
import json
import time
import re
import io
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# === CONFIGURAÇÃO DO AMBIENTE ===
def configurar_driver():
    # AJUSTE 1: Configurar pasta de download e preferências
    download_dir = os.path.join(os.getcwd(), "temp_pdfs")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    
    # Preferências para baixar PDF automaticamente sem abrir visualizador
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

# === NAVEGAÇÃO ENTRE FRAMES ===
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

# === FUNÇÃO DE BUSCA NO PORTAL AMHP ===
def extrair_detalhes_site_amhp(numero_guia):
    driver = configurar_driver()
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

        # 4. RadInput (Preenchimento)
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

        # AJUSTE 2: Exportar Guia Principal
        btn_imprimir_id = "ctl00_MainContent_btnImprimir_input"
        if entrar_no_frame_do_elemento(driver, btn_imprimir_id):
            btn_imprimir = driver.find_element(By.ID, btn_imprimir_id)
            driver.execute_script("arguments[0].click();", btn_imprimir)
            time.sleep(6) # Esperar pop-up

            # Gerenciar Janelas (Ir para o Relatório)
            for handle in driver.window_handles:
                if handle != janela_principal:
                    driver.switch_to.window(handle)
                    break
            
            # Exportar PDF no Pop-up
            try:
                dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                Select(dropdown).select_by_value("PDF")
                time.sleep(1)
                driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                time.sleep(4) # Tempo de download
                driver.close() # Fecha janela do relatório
            except:
                pass
            
            driver.switch_to.window(janela_principal)

        # AJUSTE 3: Verificar Outras Despesas
        entrar_no_frame_do_elemento(driver, "ctl00_MainContent_rbtOutrasDespesas_input")
        try:
            btn_outras = driver.find_element(By.ID, "ctl00_MainContent_rbtOutrasDespesas_input")
            if btn_outras.is_enabled():
                driver.execute_script("arguments[0].click();", btn_outras)
                time.sleep(6)
                
                # Gerenciar Janelas novamente para o novo relatório
                for handle in driver.window_handles:
                    if handle != janela_principal:
                        driver.switch_to.window(handle)
                        dropdown = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                        Select(dropdown).select_by_value("PDF")
                        time.sleep(1)
                        driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                        time.sleep(4)
                        driver.close()
                        break
                driver.switch_to.window(janela_principal)
        except:
            pass

        return {"status": "Sucesso", "arquivos": os.listdir("temp_pdfs")}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === INTERFACE ===
st.set_page_config(page_title="GABMA - Consulta AMHP", page_icon="🏥")
st.title("🏥 Consulta e Download AMHP")

if "credentials" not in st.secrets:
    st.error("Configure os Secrets.")
else:
    guia = st.text_input("Número do Atendimento:")
    if st.button("🚀 Processar e Baixar PDFs"):
        with st.spinner("Executando fluxo de impressão..."):
            res = extrair_detalhes_site_amhp(guia)
            if "erro" in res:
                st.error(res["erro"])
            else:
                st.success("Processo concluído!")
                st.write("Arquivos baixados:", res["arquivos"])
