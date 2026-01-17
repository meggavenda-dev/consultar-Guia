# -*- coding: utf-8 -*-
import os, re, time, shutil
import streamlit as st
import pandas as pd
import pdfplumber

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# ========= 1. CONFIGURAÇÕES DE AMBIENTE E BINÁRIOS =========
try:
    chrome_bin = st.secrets.get("env", {}).get("CHROME_BINARY")
    driver_bin = st.secrets.get("env", {}).get("CHROMEDRIVER_BINARY")
    if chrome_bin: os.environ["CHROME_BINARY"] = chrome_bin
    if driver_bin: os.environ["CHROMEDRIVER_BINARY"] = driver_bin
except Exception:
    pass

DOWNLOAD_TEMPORARIO = os.path.join(os.getcwd(), "temp_downloads")

def preparar_pasta_downloads():
    if os.path.exists(DOWNLOAD_TEMPORARIO):
        shutil.rmtree(DOWNLOAD_TEMPORARIO)
    os.makedirs(DOWNLOAD_TEMPORARIO)

# ========= 2. SANITIZAÇÃO E LIMPEZA DE DADOS =========
_ILLEGAL_CTRL_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")

def _sanitize_text(s: str) -> str:
    if s is None: return ""
    s = s.replace("\x00", "")
    s = _ILLEGAL_CTRL_RE.sub("", s)
    s = s.replace("\u00A0", " ").strip()
    return s

def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(lambda x: _sanitize_text(str(x)) if pd.notnull(x) else "")
    return df

# ========= 3. UTILITÁRIOS SELENIUM (CLIQUE SEGURO) =========
def js_safe_click(driver, by, value, timeout=30):
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    driver.execute_script("arguments[0].scrollIntoView(true);", el)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", el)

# ========= 4. MOTOR DE EXTRAÇÃO DE PDF (SADT + OUTRAS) =========
def processar_pdf_faturamento(caminho_pdf, n_guia_usuario):
    dados = []
    full_text = ""
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            texto_pag = page.extract_text()
            if texto_pag:
                # Normalização: reduz espaços múltiplos e remove NoneTypes
                texto_pag = re.sub(r"[ \t]+", " ", texto_pag)
                full_text += texto_pag + "\n"

    if not full_text.strip():
        return pd.DataFrame()

    # Captura de Cabeçalho (Paciente)
    paciente_match = re.search(r"10-Nome\s*\n(.*?)\n", full_text)
    paciente = paciente_match.group(1).strip() if paciente_match else "Nome Não Identificado"

    # Regex SADT (Procedimentos) - Uso de re.DOTALL para capturar quebras de linha na descrição
    regex_sadt = re.compile(
        r"(\d{2}/\d{2}/\d{4}).*?(\d{2}).*?([\d\.]+-\d).*?\n?(.*?)\s+(\d+)\s+.*?([\d,.]+)\s+([\d,.]+)",
        re.DOTALL
    )

    for match in regex_sadt.finditer(full_text):
        m = match.groups()
        dados.append({
            "Nº Guia": n_guia_usuario, "Tipo Guia": "SP/SADT", "Nome do Paciente": paciente,
            "Data de Atendimento": m[0], "Tabela": m[1], "Código": m[2],
            "Descrição": m[3].replace("\n", " ").strip(), "Quantidade": m[4],
            "Valor Unitário (R$)": m[5], "Valor Total (R$)": m[6],
            "Valor Glosado": "", "Motivo de Glosa": "", "Valor Pago": ""
        })

    # Regex Outras Despesas (Materiais/Taxas)
    regex_outras = re.compile(
        r"(\d{2})\s+(\d{8})\s+(\d+)\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)\s*\n\s*(.*?)\n",
        re.DOTALL
    )

    for match in regex_outras.finditer(full_text):
        m = match.groups()
        dados.append({
            "Nº Guia": n_guia_usuario, "Tipo Guia": "Outras Despesas", "Nome do Paciente": paciente,
            "Data de Atendimento": "", "Tabela": m[0], "Código": m[1],
            "Descrição": m[5].strip(), "Quantidade": m[2],
            "Valor Unitário (R$)": m[3], "Valor Total (R$)": m[4],
            "Valor Glosado": "", "Motivo de Glosa": "", "Valor Pago": ""
        })

    return sanitize_df(pd.DataFrame(dados))

# ========= 5. INTERFACE E EXECUÇÃO STREAMLIT =========
st.set_page_config(page_title="AMHP Automação", layout="wide")
st.title("🏥 AMHP: Automação de Faturamento")

with st.sidebar:
    st.header("Pesquisa")
    guia_input = st.text_input("🔢 Número da Guia (AMHPTISS)")
    st.info("Credenciais carregadas via st.secrets")

if st.button("🚀 Iniciar Captura Automatizada"):
    if not guia_input:
        st.error("Digite o número da guia.")
    else:
        # Carregando Secrets
        try:
            USER = st.secrets["credentials"]["usuario"]
            PASS = st.secrets["credentials"]["senha"]
        except Exception:
            st.error("Configure [credentials] no secrets.toml")
            st.stop()

        preparar_pasta_downloads()
        
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--window-size=1920,1080")
        if os.environ.get("CHROME_BINARY"):
            opts.binary_location = os.environ.get("CHROME_BINARY")
            
        prefs = {"download.default_directory": DOWNLOAD_TEMPORARIO, "plugins.always_open_pdf_externally": True}
        opts.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=opts)
        wait = WebDriverWait(driver, 40)
        
        try:
            with st.status("Executando Automação...", expanded=True) as status:
                # Login
                st.write("🔑 Login...")
                driver.get("https://portal.amhp.com.br/")
                wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(USER)
                driver.find_element(By.ID, "input-12").send_keys(PASS + Keys.ENTER)
                
                # Sistema TISS
                st.write("🔄 Acessando AMHPTISS...")
                btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
                driver.execute_script("arguments[0].click();", btn_tiss)
                time.sleep(5)
                driver.switch_to.window(driver.window_handles[-1])
                
                # Pesquisa
                st.write(f"🔍 Filtrando Guia: {guia_input}")
                driver.get("https://arhptiss.amhp.com.br/AtendimentosRealizados.aspx")
                input_f = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_rdgAtendimentosRealizados_ctl00_ctl02_ctl02_FilterTextBox_NrGuia")))
                input_f.send_keys(guia_input + Keys.ENTER)
                time.sleep(3)
                
                # Impressão/Download
                st.write("📥 Baixando Guia...")
                js_safe_click(driver, By.XPATH, "//input[contains(@id, 'btnImprimir')]")
                time.sleep(15) # Espera o download

                arquivos = [os.path.join(DOWNLOAD_TEMPORARIO, f) for f in os.listdir(DOWNLOAD_TEMPORARIO) if f.lower().endswith(".pdf")]
                
                if arquivos:
                    st.write("📄 Processando PDF...")
                    # Pega o arquivo mais recente
                    recente = max(arquivos, key=os.path.getctime)
                    df_final = processar_pdf_faturamento(recente, guia_input)
                    
                    if not df_final.empty:
                        st.subheader("📋 Conferência de Faturamento")
                        st.dataframe(df_final, use_container_width=True)
                        
                        csv = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                        st.download_button("💾 Baixar CSV", csv, f"guia_{guia_input}.csv", "text/csv")
                        status.update(label="✅ Sucesso!", state="complete")
                    else:
                        st.warning("Texto não extraído. Verifique se o PDF é uma imagem.")
                else:
                    st.error("Erro: PDF não encontrado.")

        except Exception as e:
            st.error(f"Erro Crítico: {e}")
        finally:
            driver.quit()
