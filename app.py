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

# === 1. NOVAS FUNÇÕES DE INTELIGÊNCIA (EXTRAÇÃO) ===

def extrair_texto_pdf(caminho_pdf):
    texto_full = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: texto_full += t + "\n"
    except Exception as e:
        st.error(f"Erro ao abrir PDF {os.path.basename(caminho_pdf)}: {e}")
    
    # Debug: Mostra o que foi extraído nativamente no log do Streamlit
    if len(texto_full.strip()) < 50:
        st.info(f"Ativando OCR para: {os.path.basename(caminho_pdf)}")
        try:
            # Aumentamos o DPI para 200 para melhorar a nitidez da leitura
            paginas_img = convert_from_path(caminho_pdf, dpi=200)
            for img in paginas_img:
                texto_full += image_to_string(img, lang='por') + "\n"
        except Exception as e:
            st.error(f"Erro no OCR: {e}. Verifique se poppler e tesseract estão no packages.txt")
    
    return texto_full

def processar_arquivos_baixados(diretorio, numero_guia):
    dados_lista = []
    
    # REGEX FLEXÍVEL: 
    # Agora aceita datas com ou sem espaços, e códigos TUSS com qualquer pontuação
    padrao = re.compile(
        r"(\d{2}/\d{2}/\d{4})"  # Data
        r".*?"                  # Qualquer coisa no meio
        r"(\d[\d\.\-]{5,15})"   # Código (mínimo 5 dígitos/pontos)
        r"\s+(.*?)\s+"          # Descrição
        r"(\d+)\s+"             # Quantidade
        r"([\d,.]+)\s+"         # Valor Unit
        r"([\d,.]+)",           # Valor Total
        re.DOTALL
    )
    
    arquivos = [f for f in os.listdir(diretorio) if f.lower().endswith(".pdf")]
    
    for arquivo in arquivos:
        texto = extrair_texto_pdf(os.path.join(diretorio, arquivo))
        
        # DEBUG: Se quiser ver o texto bruto para ajustar a regex, descomente a linha abaixo:
        # st.text_area(f"Texto extraído de {arquivo}", texto, height=200)

        # Normalização agressiva: remove múltiplos espaços e tabulações
        texto_limpo = re.sub(r"[ \t]+", " ", texto)
        matches = padrao.findall(texto_limpo)
        
        for m in matches:
            # Limpeza básica na descrição para remover quebras de linha residuais
            desc = m[2].replace("\n", " ").strip()
            
            dados_lista.append({
                "Atendimento": numero_guia,
                "Data": m[0],
                "Código TUSS": m[1],
                "Descrição": desc,
                "Qtd": m[3],
                "Valor Unit": m[4],
                "Valor Total": m[5],
                "Origem": arquivo
            })
            
    return pd.DataFrame(dados_lista)
    
# === 2. SUAS FUNÇÕES ORIGINAIS (MANTIDAS) ===

def configurar_driver():
    download_dir = os.path.join(os.getcwd(), "temp_pdfs")
    # Limpa a pasta antes de começar para não misturar dados
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

        # 4. RadInput (Preenchimento via ClientState)
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
            time.sleep(6)

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

        # AJUSTE 3: Verificar Outras Despesas
        entrar_no_frame_do_elemento(driver, "ctl00_MainContent_rbtOutrasDespesas_input")
        try:
            btn_outras = driver.find_element(By.ID, "ctl00_MainContent_rbtOutrasDespesas_input")
            if btn_outras.is_enabled():
                driver.execute_script("arguments[0].click();", btn_outras)
                time.sleep(6)
                
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

        # === NOVO: Processamento de Dados Pós-Download ===
        df_final = processar_arquivos_baixados(download_dir, valor_solicitado)
        
        return {"status": "Sucesso", "dados": df_final, "arquivos": os.listdir(download_dir)}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === 3. INTERFACE STREAMLIT ===

st.set_page_config(page_title="GABMA - Consulta AMHP", page_icon="🏥", layout="wide")
st.title("🏥 Consulta e Extração Automática AMHP")

if "credentials" not in st.secrets:
    st.error("Configure as credenciais nos Secrets do Streamlit.")
else:
    guia = st.text_input("Número do Atendimento:")
    if st.button("🚀 Processar e Extrair Dados"):
        if not guia:
            st.warning("Por favor, digite o número da guia.")
        else:
            with st.spinner("Executando automação e extraindo dados (isso pode levar um minuto)..."):
                res = extrair_detalhes_site_amhp(guia)
                
                if "erro" in res:
                    st.error(f"Erro no processo: {res['erro']}")
                else:
                    st.success("Processo concluído!")
                    
                    df = res["dados"]
                    if not df.empty:
                        st.subheader("📋 Dados de Faturamento Extraídos")
                        st.dataframe(df, use_container_width=True)
                        
                        # Botão de download
                        csv = df.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            label="📥 Baixar Planilha CSV",
                            data=csv,
                            file_name=f"faturamento_guia_{guia}.csv",
                            mime="text/csv",
                        )
                    else:
                        st.warning("Nenhum dado de procedimento foi encontrado nos PDFs baixados.")
                    
                    with st.expander("Ver arquivos baixados"):
                        st.write(res["arquivos"])
