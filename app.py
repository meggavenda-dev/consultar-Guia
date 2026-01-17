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

# === CONFIGURAÇÃO DO SELENIUM ===
def configurar_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    
    # Tenta localizar binários no Streamlit Cloud ou Local
    chrome_bin = os.environ.get("CHROME_BINARY", "/usr/bin/chromium")
    if os.path.exists(chrome_bin):
        opts.binary_location = chrome_bin

    try:
        driver = webdriver.Chrome(options=opts)
    except:
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)
    return driver

# === FUNÇÃO DE BUSCA NOS IFRAMES (SISTEMÁTICA) ===
def find_element_in_frames(driver, element_id):
    """Varre o documento e todos os iframes pelo ID."""
    driver.switch_to.default_content()
    try:
        return driver.find_element(By.ID, element_id)
    except:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for i, frame in enumerate(iframes):
            driver.switch_to.default_content()
            driver.switch_to.frame(i)
            try:
                return driver.find_element(By.ID, element_id)
            except:
                continue
    return None

# === FUNÇÃO PRINCIPAL DA AMHP ===
def extrair_detalhes_site_amhp(numero_guia):
    driver = configurar_driver()
    wait = WebDriverWait(driver, 25)
    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())
    
    try:
        # 1. Login
        driver.get("https://portal.amhp.com.br/")
        wait.until(EC.presence_of_element_located((By.ID, "input-9"))).send_keys(st.secrets["credentials"]["usuario"])
        driver.find_element(By.ID, "input-12").send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

        # 2. Entrar no módulo TISS
        time.sleep(5)
        btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
        driver.execute_script("arguments[0].click();", btn_tiss)
        
        time.sleep(5)
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        # 3. Ir para Atendimentos
        driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
        time.sleep(3)

        # 4. TRATAMENTO PROFUNDO DO RADINPUT (O CORAÇÃO DA SOLUÇÃO)
        input_id = "ctl00_MainContent_rtbNumeroAtendimento"
        state_id = "ctl00_MainContent_rtbNumeroAtendimento_ClientState"
        
        campo = find_element_in_frames(driver, input_id)
        if not campo:
            raise Exception("Campo de busca não encontrado nos frames.")

        # Criamos o JSON que o servidor AMHP espera no campo oculto
        client_state = json.dumps({
            "enabled": True, "emptyMessage": "", "validationText": valor_solicitado,
            "valueAsString": valor_solicitado, "lastSetTextBoxValue": valor_solicitado
        })

        # Injeção via JS: preenche o visível e o oculto simultaneamente
        driver.execute_script("""
            var val = arguments[0];
            var json = arguments[1];
            var el = document.getElementById(arguments[2]);
            var state = document.getElementById(arguments[3]);
            
            if(el) {
                el.value = val;
                if(state) state.value = json;
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
        """, valor_solicitado, client_state, input_id, state_id)

        # 5. Clique no Buscar
        btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
        driver.execute_script("arguments[0].click();", btn_buscar)
        
        # 6. Coleta dos Resultados
        time.sleep(4)
        link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
        driver.execute_script("arguments[0].click();", link_guia)
        
        time.sleep(3)
        # Extração de exemplo
        paciente = driver.find_element(By.ID, "ctl00_MainContent_txtNomeBeneficiario").get_attribute("value")
        
        return {"paciente": paciente, "guia": valor_solicitado, "status": "Sucesso"}

    except Exception as e:
        driver.save_screenshot("erro_amhptiss.png")
        return {"erro": str(e)}
    finally:
        driver.quit()

# === INTERFACE STREAMLIT ===
st.title("Automação Portal AMHP")

if "credentials" not in st.secrets:
    st.error("Configure o arquivo .streamlit/secrets.toml com as chaves [credentials] usuario e senha.")
else:
    guia_input = st.text_input("Número da Guia AMHP:", value="61789641")
    if st.button("Consultar no Site"):
        with st.spinner("Acessando portal e injetando dados..."):
            resultado = extrair_detalhes_site_amhp(guia_input)
            if "erro" in resultado:
                st.error(f"Erro: {resultado['erro']}")
                if os.path.exists("erro_amhptiss.png"):
                    st.image("erro_amhptiss.png")
            else:
                st.success(f"Guia localizada! Paciente: {resultado['paciente']}")
