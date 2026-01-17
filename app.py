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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# === CONFIGURAÇÃO DO AMBIENTE ===
def configurar_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    
    # Localização de binários (Streamlit Cloud vs Local)
    chrome_bin = os.environ.get("CHROME_BINARY", "/usr/bin/chromium")
    if os.path.exists(chrome_bin):
        opts.binary_location = chrome_bin

    try:
        driver = webdriver.Chrome(options=opts)
    except:
        # Fallback para caminhos comuns em servidores Linux
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)
    return driver

# === NAVEGAÇÃO ENTRE FRAMES ===
def entrar_no_frame_do_elemento(driver, element_id):
    """Percorre todos os frames do site até achar o ID desejado."""
    driver.switch_to.default_content()
    try:
        driver.find_element(By.ID, element_id)
        return True # Já está no root
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
    
    try:
        # 1. Login
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # 2. Acesso ao Módulo TISS (Aguarda carregamento da Dashboard)
        time.sleep(6)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        # Troca para a nova aba que o portal abre
        time.sleep(5)
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        # 3. Navegação Direta
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(4)

        # 4. A SOLUÇÃO SISTEMÁTICA PARA O RADINPUT
        input_id = "ctl00_MainContent_rtbNumeroAtendimento"
        state_id = "ctl00_MainContent_rtbNumeroAtendimento_ClientState"
        
        if not entrar_no_frame_do_elemento(driver, input_id):
            raise Exception("Não foi possível encontrar o campo de busca em nenhum frame.")

        # Preparação do JSON de Estado (Essencial para Telerik)
        client_state = json.dumps({
            "enabled": True,
            "emptyMessage": "",
            "validationText": valor_solicitado,
            "valueAsString": valor_solicitado,
            "lastSetTextBoxValue": valor_solicitado
        })

        # Injeção via JS: sincroniza o campo visível com o motor do site
        driver.execute_script("""
            var el = document.getElementById(arguments[0]);
            var state = document.getElementById(arguments[1]);
            var val = arguments[2];
            var json = arguments[3];
            
            if(el) {
                el.value = val;
                if(state) state.value = json;
                // Dispara eventos de validação
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
        """, input_id, state_id, valor_solicitado, client_state)

        # 5. Clique no Botão Buscar
        btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
        driver.execute_script("arguments[0].click();", btn_buscar)
        
        # 6. Coleta de Resultados
        time.sleep(4)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rgMasterTable")))
        
        # Clica no link que contém exatamente o número da guia
        link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
        driver.execute_script("arguments[0].click();", link_guia)
        
        time.sleep(3)
        # Extração de dados da tela de detalhes
        paciente = driver.find_element(By.ID, "ctl00_MainContent_txtNomeBeneficiario").get_attribute("value")
        data_atd = driver.find_element(By.ID, "ctl00_MainContent_dtDataAtendimento_dateInput").get_attribute("value")
        
        return {"paciente": paciente, "data": data_atd, "status": "Sucesso"}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === INTERFACE DO PROGRAMA ===
st.set_page_config(page_title="GABMA - Consulta AMHP", page_icon="🏥")
st.title("🏥 Consulta de Atendimento AMHP")

if "credentials" not in st.secrets:
    st.error("Erro: Credenciais não encontradas. Configure o 'Secrets' no Streamlit.")
else:
    guia = st.text_input("Número do Atendimento (AMHPTISS):", placeholder="Ex: 61789641")
    
    if st.button("🔍 Pesquisar no Portal"):
        if not guia:
            st.warning("Por favor, digite o número da guia.")
        else:
            with st.spinner("Conectando ao portal AMHP e injetando dados..."):
                res = extrair_detalhes_site_amhp(guia)
                
                if "erro" in res:
                    st.error(f"Erro na consulta: {res['erro']}")
                    if os.path.exists("erro_amhptiss.png"):
                        st.image("erro_amhptiss.png", caption="Momento do erro")
                else:
                    st.success("Dados recuperados com sucesso!")
                    col1, col2 = st.columns(2)
                    col1.metric("Paciente", res['paciente'])
                    col2.metric("Data", res['data'])
