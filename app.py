# -*- coding: utf-8 -*-
import os, re, time, shutil, json
import streamlit as st
import pandas as pd
import pdfplumber
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# === CONFIGURAÇÃO DE AMBIENTE ===
DOWNLOAD_DIR = os.path.join(os.getcwd(), "temp_pdfs")

def preparar_ambiente():
    if os.path.exists(DOWNLOAD_DIR):
        shutil.rmtree(DOWNLOAD_DIR)
    os.makedirs(DOWNLOAD_DIR)

def configurar_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    opts.add_experimental_option("prefs", prefs)
    
    # Busca binários no secrets para Streamlit Cloud
    chrome_bin = st.secrets.get("env", {}).get("CHROME_BINARY")
    if chrome_bin: opts.binary_location = chrome_bin
    
    return webdriver.Chrome(options=opts)

# === MOTOR DE EXTRAÇÃO PDF (ITEM A ITEM) ===
def processar_pdfs_baixados(guia_referencia):
    dados_finais = []
    arquivos = [os.path.join(DOWNLOAD_DIR, f) for f in os.listdir(DOWNLOAD_DIR) if f.lower().endswith(".pdf")]
    
    for arq in arquivos:
        with pdfplumber.open(arq) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text: full_text += text + "\n"
            
            # Identifica se é SADT ou Outras Despesas
            tipo_guia = "SP/SADT" if "SADT" in full_text.upper() else "Outras Despesas"
            
            # Captura nome do paciente (Campo 10)
            paciente_match = re.search(r"10-Nome\s*\n(.*?)\n", full_text)
            paciente = paciente_match.group(1).strip() if paciente_match else "Não Identificado"

            # Regex para itens (SADT)
            regex_sadt = re.compile(r"(\d{2}/\d{2}/\d{4}).*?(\d{2}).*?([\d\.]+-\d).*?\n?(.*?)\s+(\d+)\s+.*?([\d,.]+)\s+([\d,.]+)", re.DOTALL)
            for match in regex_sadt.finditer(full_text):
                res = match.groups()
                dados_factual = {
                    "Nº Guia": guia_referencia, "Tipo": tipo_guia, "Paciente": paciente,
                    "Data": res[0], "Tab": res[1], "Código": res[2], "Descrição": res[3].strip(),
                    "Qtd": res[4], "Unit": res[5], "Total": res[6], "Glosa": "", "Pago": ""
                }
                dados_finais.append(dados_factual)
            
            # Regex para itens (Outras Despesas)
            regex_outras = re.compile(r"(\d{2})\s+(\d{8})\s+(\d+)\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)\s*\n\s*(.*?)\n", re.DOTALL)
            for match in regex_outras.finditer(full_text):
                res = match.groups()
                dados_finais.append({
                    "Nº Guia": guia_referencia, "Tipo": tipo_guia, "Paciente": paciente,
                    "Data": "", "Tab": res[0], "Código": res[1], "Descrição": res[5].strip(),
                    "Qtd": res[2], "Unit": res[3], "Total": res[4], "Glosa": "", "Pago": ""
                })
                
    return pd.DataFrame(dados_finais)

# === INTERFACE E AUTOMAÇÃO ===
st.set_page_config(page_title="GABMA - AMHP Pro", layout="wide")
st.title("🏥 Automação AMHP: Download e Estruturação")

guia_alvo = st.text_input("Número do Atendimento/Guia:")

if st.button("🚀 Iniciar Processo Completo"):
    if not guia_alvo:
        st.error("Informe o número da guia.")
    else:
        preparar_ambiente()
        driver = configurar_driver()
        wait = WebDriverWait(driver, 35)
        
        try:
            with st.status("Executando fluxo no Portal...", expanded=True) as status:
                # 1. LOGIN
                driver.get("https://portal.amhp.com.br/")
                wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
                driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)
                
                # 2. ENTRAR NO TISS
                time.sleep(6)
                btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
                driver.execute_script("arguments[0].click();", btn_tiss)
                time.sleep(5)
                
                # Gerenciar Janela do TISS
                driver.switch_to.window(driver.window_handles[-1])
                janela_tiss = driver.current_window_handle
                
                # 3. BUSCA
                driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
                st.write("🔍 Filtrando atendimento...")
                input_f = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_rtbNumeroAtendimento")))
                input_f.send_keys(guia_alvo + Keys.ENTER)
                time.sleep(4)
                
                # Clicar no link do atendimento na tabela
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rgMasterTable")))
                driver.find_element(By.XPATH, f"//a[contains(text(), '{guia_alvo}')]").click()
                time.sleep(3)

                # 4. DOWNLOAD SADT E OUTRAS
                botoes_impressao = [
                    ("SADT", "ctl00_MainContent_btnImprimir_input"),
                    ("Outras Despesas", "ctl00_MainContent_rbtOutrasDespesas_input")
                ]

                for nome, btn_id in botoes_impressao:
                    try:
                        st.write(f"📥 Gerando PDF: {nome}...")
                        btn = wait.until(EC.element_to_be_clickable((By.ID, btn_id)))
                        driver.execute_script("arguments[0].click();", btn)
                        time.sleep(6)
                        
                        # Trocar para janela do ReportView
                        for handle in driver.window_handles:
                            if handle != janela_tiss:
                                driver.switch_to.window(handle)
                                # Exportar PDF
                                drop = wait.until(EC.presence_of_element_located((By.ID, "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList")))
                                Select(drop).select_by_value("PDF")
                                driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export").click()
                                time.sleep(5) # Tempo de download
                                driver.close()
                                driver.switch_to.window(janela_tiss)
                                break
                    except:
                        st.warning(f"Aviso: {nome} não disponível para esta guia.")

                # 5. PROCESSAMENTO FINAL
                st.write("📊 Estruturando dados...")
                df_final = processar_pdfs_baixados(guia_alvo)
                
                if not df_final.empty:
                    st.dataframe(df_final, use_container_width=True)
                    csv = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("💾 Baixar Planilha Final", csv, f"faturamento_{guia_alvo}.csv", "text/csv")
                    status.update(label="✅ Tudo pronto!", state="complete")
                else:
                    st.error("Nenhum dado extraído dos PDFs.")

        except Exception as e:
            st.error(f"Erro no processo: {e}")
        finally:
            driver.quit()
