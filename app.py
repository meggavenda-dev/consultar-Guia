# -*- coding: utf-8 -*-
import os, re, time, shutil
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

# ========= CONFIGURAÇÕES E PASTAS =========
DOWNLOAD_TEMPORARIO = os.path.join(os.getcwd(), "temp_downloads")
if not os.path.exists(DOWNLOAD_TEMPORARIO):
    os.makedirs(DOWNLOAD_TEMPORARIO)

st.set_page_config(page_title="AMHP Automação", layout="wide")
st.title("🏥 AMHP: Automação de Downloads e Faturamento")

# ========= LÓGICA DE EXTRAÇÃO DE PDF (SAD + OUTRAS) =========

def extrair_dados_guias(caminho_pdf, n_guia_usuario):
    dados_faturamento = []
    full_text = ""
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            texto_pag = page.extract_text()
            if texto_pag:
                # Normalização de espaços e quebras para a Regex não falhar
                texto_pag = re.sub(r"[ \t]+", " ", texto_pag)
                full_text += texto_pag + "\n"
    
    if not full_text.strip():
        return pd.DataFrame()

    # 1. Captura de Dados do Cabeçalho (Paciente e Data)
    # Tenta localizar o nome após o campo 10-Nome
    paciente_match = re.search(r"10-Nome\s*\n(.*?)\n", full_text)
    paciente = paciente_match.group(1).strip() if paciente_match else "Não Identificado"
    
    # 2. Regex para SADT (Procedimentos)
    # Padrão: Data | Tab | Código | Descrição | Qtd | Valor Unit | Valor Total
    regex_sadt = re.compile(
        r"(\d{2}/\d{2}/\d{4}).*?(\d{2}).*?([\d\.]+-\d).*?\n?(.*?)\s+(\d+)\s+.*?([\d,.]+)\s+([\d,.]+)",
        re.DOTALL
    )

    for match in regex_sadt.finditer(full_text):
        res = match.groups()
        dados_faturamento.append({
            "Nº Guia": n_guia_usuario,
            "Tipo Guia": "SP/SADT",
            "Nome do Paciente": paciente,
            "Data de Atendimento": res[0],
            "Tabela": res[1],
            "Código": res[2],
            "Descrição": res[3].replace("\n", " ").strip(),
            "Quantidade": res[4],
            "Valor Unitário (R$)": res[5],
            "Valor Total (R$)": res[6],
            "Valor Glosado": "",
            "Motivo de Glosa": "",
            "Valor Pago": ""
        })

    # 3. Regex para Outras Despesas (Materiais)
    # Padrão simplificado para capturar a tabela e o código do item
    regex_outras = re.compile(
        r"(\d{2})\s+(\d{8})\s+(\d+)\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)\s*\n\s*(.*?)\n",
        re.DOTALL
    )

    for match in regex_outras.finditer(full_text):
        res = match.groups()
        dados_faturamento.append({
            "Nº Guia": n_guia_usuario,
            "Tipo Guia": "Outras Despesas",
            "Nome do Paciente": paciente,
            "Data de Atendimento": "", # Geralmente atrelado à guia principal
            "Tabela": res[0],
            "Código": res[1],
            "Descrição": res[5].strip(),
            "Quantidade": res[2],
            "Valor Unitário (R$)": res[3],
            "Valor Total (R$)": res[4],
            "Valor Glosado": "",
            "Motivo de Glosa": "",
            "Valor Pago": ""
        })

    return pd.DataFrame(dados_faturamento)

# ========= AUTOMAÇÃO SELENIUM =========

def configurar_driver():
    opts = Options()
    opts.add_argument("--headless") # Mude para False se quiser ver o navegador
    opts.add_argument("--no-sandbox")
    opts.add_argument("--window-size=1920,1080")
    
    prefs = {
        "download.default_directory": DOWNLOAD_TEMPORARIO,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    opts.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=opts)

# ========= INTERFACE PRINCIPAL =========

with st.sidebar:
    st.header("Acesso AMHP")
    user_input = st.text_input("Usuário")
    pass_input = st.text_input("Senha", type="password")
    guia_input = st.text_input("Número da Guia para Pesquisa")

if st.button("🚀 Iniciar Processo"):
    if not all([user_input, pass_input, guia_input]):
        st.error("Por favor, preencha todos os campos no menu lateral.")
    else:
        driver = configurar_driver()
        wait = WebDriverWait(driver, 30)
        
        try:
            with st.status("Executando automação...", expanded=True) as status:
                # 1. Login
                st.write("🔑 Acessando portal...")
                driver.get("https://portal.amhp.com.br/")
                wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(user_input)
                driver.find_element(By.ID, "input-12").send_keys(pass_input + Keys.ENTER)
                
                # 2. Navegação para AMHPTISS
                st.write("🔄 Entrando no sistema TISS...")
                btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
                driver.execute_script("arguments[0].click();", btn_tiss)
                time.sleep(5)
                driver.switch_to.window(driver.window_handles[-1])
                
                # 3. Pesquisa de Guia
                st.write(f"🔍 Pesquisando guia {guia_input}...")
                driver.get("https://arhptiss.amhp.com.br/AtendimentosRealizados.aspx") # URL direta para agilizar
                
                input_filtro = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_rdgAtendimentosRealizados_ctl00_ctl02_ctl02_FilterTextBox_NrGuia")))
                input_filtro.send_keys(guia_input + Keys.ENTER)
                time.sleep(3)
                
                # 4. Download do PDF
                st.write("📥 Gerando e baixando PDF...")
                btn_print = driver.find_element(By.XPATH, "//input[contains(@id, 'btnImprimir')]")
                driver.execute_script("arguments[0].click();", btn_print)
                time.sleep(10) # Tempo para o download concluir

                # 5. Processamento dos Arquivos
                st.write("📄 Lendo dados do faturamento...")
                arquivos = [os.path.join(DOWNLOAD_TEMPORARIO, f) for f in os.listdir(DOWNLOAD_TEMPORARIO) if f.endswith(".pdf")]
                
                if arquivos:
                    df_final = pd.DataFrame()
                    for arq in arquivos:
                        df_temp = extrair_dados_guias(arq, guia_input)
                        df_final = pd.concat([df_final, df_temp], ignore_index=True)
                    
                    status.update(label="✅ Processo Concluído!", state="complete")
                    
                    st.subheader("📊 Planilha de Faturamento Gerada")
                    st.dataframe(df_final, use_container_width=True)
                    
                    # Download do CSV para o usuário
                    csv = df_final.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                    st.download_button("💾 Baixar Planilha (CSV)", csv, f"faturamento_guia_{guia_input}.csv", "text/csv")
                else:
                    st.error("Nenhum PDF foi encontrado na pasta de downloads.")

        except Exception as e:
            st.error(f"Erro durante a execução: {e}")
        finally:
            driver.quit()
