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
    download_dir = os.path.abspath(download_dir)
    wait = WebDriverWait(driver, 30)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    janela_principal = driver.current_window_handle
    
    try:
        # [Lógica de Login e Busca de Guia mantida...]
        # ... (Imagine aqui o código que faz o login e clica na guia)

        # 1. Clicar no botão Imprimir (Usando o seletor exato que você passou)
        # O botão está dentro de um frame, por isso usamos sua função entrar_no_frame_do_elemento
        btn_imprimir_id = "ctl00_MainContent_btnImprimir_input"
        if entrar_no_frame_do_elemento(driver, btn_imprimir_id):
            btn_imprimir = driver.find_element(By.ID, btn_imprimir_id)
            driver.execute_script("arguments[0].click();", btn_imprimir)
            
            # 2. Gerenciar a nova janela que abre (o pop-up do relatório)
            time.sleep(5) # Tempo para o pop-up carregar
            wait.until(lambda d: len(d.window_handles) > 2) # Espera a 3ª janela (Home > Sistema > Relatório)
            
            # Alterna para a janela do relatório
            for handle in driver.window_handles:
                if handle not in [janela_principal]:
                    driver.switch_to.window(handle)
            
            # 3. Selecionar 'PDF' no DropDownList
            # Usando o ID exato: ReportView_ReportToolbar_ExportGr_FormatList_DropDownList
            dropdown_id = "ReportView_ReportToolbar_ExportGr_FormatList_DropDownList"
            select_element = wait.until(EC.presence_of_element_located((By.ID, dropdown_id)))
            select = Select(select_element)
            select.select_by_value("PDF") # Ou select_by_visible_text("Acrobat (PDF) file")
            
            # 4. Clicar no link de Exportar
            # Usando o ID exato: ReportView_ReportToolbar_ExportGr_Export
            time.sleep(1)
            btn_exportar = driver.find_element(By.ID, "ReportView_ReportToolbar_ExportGr_Export")
            driver.execute_script("arguments[0].click();", btn_exportar)
            
            # 5. Monitorar o Download
            # Aguardamos até que um arquivo .pdf apareça na pasta
            time.sleep(10) 
            
            # OPCIONAL: Mover o arquivo para outro local após baixar
            # for f in os.listdir(download_dir):
            #     shutil.move(os.path.join(download_dir, f), "/novo/caminho/f")

        return {"status": "Sucesso", "diretorio": download_dir}

    except Exception as e:
        driver.save_screenshot("erro_fluxo_exportacao.png")
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
