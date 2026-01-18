
# app_amhp.py
# -*- coding: utf-8 -*-

import os
import re
import io
import time
import shutil
import json
import pandas as pd
import streamlit as st

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import pdfplumber
from pdf2image import convert_from_path
from pytesseract import image_to_string

# ============= CONFIGS GERAIS =============
APP_TITLE = "🏥 Inteligência de Faturamento AMHP"
PAGE_ICON = "🏥"
TIMEOUT_PADRAO = 30  # segundos
TIMEOUT_DOWNLOAD = 90

# ============= FUNÇÕES DE SUPORTE (CHROME/SELENIUM) =============

def habilitar_download_headless(driver, download_dir):
    """Libera downloads em modo headless via CDP (se suportado)."""
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(download_dir)
        })
    except Exception:
        pass

def configurar_driver(headless=True, download_dir=None, chrome_binary_env="CHROME_BINARY"):
    """Cria e configura o driver do Chrome/Chromium com diretório de download controlado."""
    if download_dir is None:
        download_dir = os.path.join(os.getcwd(), "temp_pdfs")

    # Limpeza preventiva
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir)
    os.makedirs(download_dir, exist_ok=True)

    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0
    }
    opts.add_experimental_option("prefs", prefs)

    # Permite apontar um binário específico (ex.: contêiner com Chromium)
    chrome_bin = os.environ.get(chrome_binary_env, "/usr/bin/chromium")
    if os.path.exists(chrome_bin):
        opts.binary_location = chrome_bin

    try:
        driver = webdriver.Chrome(options=opts)
    except Exception:
        # Fallback para path fixo do chromedriver; ajuste se necessário.
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=opts)

    habilitar_download_headless(driver, download_dir)
    return driver, download_dir

def switch_to_default_and_return(driver):
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

def switch_to_frame_contendo(driver, by, locator, max_depth=6):
    """
    Busca recursivamente por iframes até encontrar um elemento (by, locator).
    Mantém o driver DENTRO do frame onde o elemento foi encontrado.
    Retorna True/False.
    """
    switch_to_default_and_return(driver)

    def _search_in_frame(depth=0):
        if depth > max_depth:
            return False
        # Tenta no contexto atual
        try:
            driver.find_element(by, locator)
            return True
        except Exception:
            pass

        # Percorre iframes deste nível
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for idx in range(len(frames)):
            # Sempre recaptura (DOM pode mudar)
            switch_to_default_and_return(driver)
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            try:
                driver.switch_to.frame(frames[idx])
                if _search_in_frame(depth + 1):
                    return True
            except Exception:
                continue
        return False

    return _search_in_frame()

def entrar_no_iframe_reportviewer_por_toolbar(driver, wait, max_depth=6):
    """
    Tenta entrar recursivamente no iframe que contenha a toolbar do ReportViewer
    (dropdown de formato e/ou botão Export).
    """
    # Procura por um seletor genérico dentro do iframe (select/button/input com 'Export')
    xp_toolbar = "//*[self::select or self::button or self::input][contains(@id,'Export') or contains(@title,'Export') or contains(@aria-label,'Export')]"
    achou = switch_to_frame_contendo(driver, By.XPATH, xp_toolbar, max_depth=max_depth)
    return achou

def mudar_para_contexto_relatorio(driver, wait, janela_principal, janela_sistema):
    """
    Depois de clicar em 'Imprimir' ou 'Outras Despesas':
      1) Se abrir nova janela/aba -> troca pra ela (modo 'window').
      2) Senão, tenta achar o iframe do ReportViewer na própria aba (modo 'iframe').
    Retorna dict: { "modo": "window"|"iframe"|None, "success": bool }.
    """
    # 1) tenta nova janela/aba
    try:
        wait.until(lambda d: len(d.window_handles) > 2)
        for h in driver.window_handles:
            if h not in [janela_principal, janela_sistema]:
                driver.switch_to.window(h)
                return {"modo": "window", "success": True}
    except Exception:
        pass

    # 2) tenta localizar ReportViewer embutido via iframe
    try:
        driver.switch_to.window(janela_sistema)
        if entrar_no_iframe_reportviewer_por_toolbar(driver, wait, max_depth=6):
            return {"modo": "iframe", "success": True}
    except Exception:
        pass

    return {"modo": None, "success": False}

def esperar_pdf_baixar(download_dir, timeout=TIMEOUT_DOWNLOAD):
    """
    Aguarda até existir pelo menos 1 PDF completo (sem .crdownload) no diretório.
    Considera estabilização de tamanho para evitar pegar arquivo incompleto.
    """
    t0 = time.time()
    ultimo_total = -1
    while time.time() - t0 < timeout:
        arquivos = os.listdir(download_dir)
        pdfs = [a for a in arquivos if a.lower().endswith(".pdf")]
        crds = [a for a in arquivos if a.lower().endswith(".crdownload")]
        if pdfs and not crds:
            total = sum(os.path.getsize(os.path.join(download_dir, p)) for p in pdfs)
            if total == ultimo_total and total > 0:
                return True
            ultimo_total = total
        time.sleep(1.2)
    raise TimeoutException("Tempo excedido aguardando download do PDF.")

def exportar_pdf_reportviewer(driver, wait, download_dir):
    """
    Interage com a toolbar do ReportViewer:
      - Seleciona 'PDF' no dropdown
      - Clica no botão 'Export'
      - Aguarda download finalizar
    """
    dropdown_xpaths = [
        "//select[contains(@id,'Export') or contains(@title,'Export') or contains(@aria-label,'Export')]",
        "//select[contains(@id,'FormatList') or contains(@name,'FormatList')]"
    ]
    export_btn_xpaths = [
        "//*[@id and (contains(@id,'Export') or contains(@title,'Export')) and (self::button or self::input)]",
        "//input[@type='submit' and (contains(@id,'Export') or contains(@title,'Export'))]"
    ]

    dropdown = None
    for xp in dropdown_xpaths:
        try:
            dropdown = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            break
        except Exception:
            continue
    if dropdown is None:
        raise NoSuchElementException("Dropdown de formato do ReportViewer não encontrado.")

    # Seleciona por texto visível e, se falhar, por value
    try:
        Select(dropdown).select_by_visible_text("PDF")
    except Exception:
        try:
            Select(dropdown).select_by_value("PDF")
        except Exception:
            # Tenta com casos localizados
            try:
                Select(dropdown).select_by_visible_text("Pdf")
            except Exception:
                pass

    export_btn = None
    for xp in export_btn_xpaths:
        try:
            export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            break
        except Exception:
            continue
    if export_btn is None:
        raise NoSuchElementException("Botão Export do ReportViewer não encontrado.")

    driver.execute_script("arguments[0].click();", export_btn)
    esperar_pdf_baixar(download_dir, timeout=TIMEOUT_DOWNLOAD)

# ============= OCR / PDF =============

def extrair_texto_pdf(caminho_pdf, forcar_ocr=False, lang="por"):
    """
    Tenta extrair texto com pdfplumber; se for imagem ou forçar OCR, usa pytesseract.
    """
    texto_full = ""
    if not forcar_ocr:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        texto_full += t + "\n"
        except Exception as e:
            st.error(f"Erro ao ler PDF nativo: {e}")

    if forcar_ocr or len(texto_full.strip()) < 50:
        try:
            paginas_img = convert_from_path(caminho_pdf, dpi=200)
            for img in paginas_img:
                texto_full += image_to_string(img, lang=lang) + "\n"
        except Exception as e:
            st.error(f"Erro no OCR (verifique se tesseract/poppler estão instalados): {e}")

    return texto_full

def processar_arquivos_baixados(diretorio, numero_guia, forcar_ocr=False):
    """
    Percorre PDFs do diretório, extrai texto e aplica regex para linhas de faturamento.
    """
    dados_lista = []
    # Regex flexível para Data, Código TUSS, Descrição, Qtd, Unit, Total
    padrao = re.compile(
        r"(\d{2}/\d{2}/\d{4})"        # Data
        r".*?"                         # salto não guloso
        r"(\d[\d\.\-]{5,15})"          # Código TUSS
        r"\s+(.*?)\s+"                 # Descrição
        r"(\d+)\s+"                    # Qtd
        r"([\d\.\,]+)\s+"              # Valor Unit
        r"([\d\.\,]+)",                # Valor Total
        re.DOTALL
    )

    for arquivo in os.listdir(diretorio):
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(diretorio, arquivo)
            texto = extrair_texto_pdf(caminho, forcar_ocr=forcar_ocr)
            # Normaliza espaços e quebras
            texto_limpo = re.sub(r"[ \t]+", " ", texto)
            matches = padrao.findall(texto_limpo)

            for m in matches:
                dados_lista.append({
                    "Guia": str(numero_guia),
                    "Data": m[0],
                    "Código": m[1],
                    "Descrição": m[2].replace("\n", " ").strip(),
                    "Qtd": m[3],
                    "Valor Unit": m[4],
                    "Valor Total": m[5],
                    "Arquivo Origem": arquivo
                })

    return pd.DataFrame(dados_lista)

# ============= AUTOMAÇÃO AMHP/AMHPTISS =============

def _login_portal(driver, wait, status_cb=None):
    driver.get("https://portal.amhp.com.br/")
    if status_cb:
        status_cb("🔐 Fazendo login no portal…")

    # Tenta por IDs conhecidos e fallback genérico
    try:
        u = wait.until(EC.presence_of_element_located((By.ID, "input-9")))
        p = driver.find_element(By.ID, "input-12")
    except Exception:
        # fallback: primeiro input texto e depois input senha
        u = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        p = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))

    u.clear()
    u.send_keys(st.secrets["credentials"]["usuario"])
    p.clear()
    p.send_keys(st.secrets["credentials"]["senha"] + Keys.ENTER)

def _entrar_amhptiss(driver, wait, status_cb=None):
    # Aguarda UI carregar e clica em AMHPTISS
    if status_cb:
        status_cb("➡️ Abrindo AMHPTISS…")
    time.sleep(6)
    btn_tiss = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'AMHPTISS')]")))
    driver.execute_script("arguments[0].click();", btn_tiss)

    # Troca para nova janela
    wait.until(lambda d: len(d.window_handles) > 1)
    janela_principal = driver.window_handle
    janela_principal = driver.current_window_handle
    for h in driver.window_handles:
        if h != janela_principal:
            driver.switch_to.window(h)
            break
    janela_sistema = driver.current_window_handle
    return janela_principal, janela_sistema

def _buscar_atendimento(driver, wait, valor_solicitado, status_cb=None):
    if status_cb:
        status_cb(f"🔎 Acessando tela de Atendimentos e consultando: {valor_solicitado}…")
    driver.get("https://amhptiss.amhp.com.br/AtendimentosRealizados.aspx")
    time.sleep(3)

    input_atendimento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_rtbNumeroAtendimento")))
    driver.execute_script("arguments[0].value = arguments[1];", input_atendimento, valor_solicitado)

    btn_buscar = driver.find_element(By.ID, "ctl00_MainContent_btnBuscar_input")
    driver.execute_script("arguments[0].click();", btn_buscar)

    # Abre detalhamento da guia
    time.sleep(4)
    link_guia = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{valor_solicitado}')]")))
    driver.execute_script("arguments[0].click();", link_guia)
    time.sleep(2)

def extrair_detalhes_site_amhp(numero_guia, headless=True, preferir_outras=True, forcar_ocr=False, debug=False):
    """
    Fluxo completo: login, buscar guia, abrir relatório, exportar PDF, ler PDFs e extrair dados.
    Retorna: dict {status, dados (DataFrame), diretorio (downloads)} ou {erro}
    """
    driver, download_dir = configurar_driver(headless=headless)
    download_dir = os.path.abspath(download_dir)
    wait = WebDriverWait(driver, TIMEOUT_PADRAO)

    valor_solicitado = re.sub(r"\D+", "", str(numero_guia).strip())

    status_msg = {"msg": ""}

    def set_status(m):
        status_msg["msg"] = m

    # Para UI (externo) poder ver status
    if "status_container" not in st.session_state:
        st.session_state["status_container"] = st.empty()

    def update_ui_status():
        st.session_state["status_container"].info(status_msg["msg"])

    try:
        _login_portal(driver, wait, status_cb=set_status)
        update_ui_status()

        janela_principal, janela_sistema = _entrar_amhptiss(driver, wait, status_cb=set_status)
        update_ui_status()

        _buscar_atendimento(driver, wait, valor_solicitado, status_cb=set_status)
        update_ui_status()

        # Tenta os botões em ordem: preferir Outras Despesas ou Imprimir
        botoes = [
            "ctl00_MainContent_rbtOutrasDespesas_input",
            "ctl00_MainContent_btnImprimir_input",
        ] if preferir_outras else [
            "ctl00_MainContent_btnImprimir_input",
            "ctl00_MainContent_rbtOutrasDespesas_input",
        ]

        click_results = []
        for id_btn in botoes:
            try:
                set_status(f"🖨️ Procurando botão no DOM e clicando: {id_btn}…")
                update_ui_status()

                driver.switch_to.window(janela_sistema)
                achou = switch_to_frame_contendo(driver, By.ID, id_btn, max_depth=6)
                if not achou:
                    click_results.append(f"Não achei {id_btn} em nenhum iframe.")
                    continue

                btn = driver.find_element(By.ID, id_btn)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", btn)

                # Mudar para contexto do relatório (nova aba ou iframe)
                driver.switch_to.window(janela_sistema)
                set_status("🧭 Aguardando relatório (janela/iframe)…")
                update_ui_status()

                ctx = mudar_para_contexto_relatorio(driver, wait, janela_principal, janela_sistema)
                if not ctx["success"]:
                    click_results.append(f"Cliquei em {id_btn}, mas não localizei relatório (nova janela/iframe).")
                    continue

                # Exporta para PDF
                set_status("📄 Exportando PDF no ReportViewer…")
                update_ui_status()
                exportar_pdf_reportviewer(driver, wait, download_dir)

                # Fecha a janela do relatório, se foi popup
                if ctx["modo"] == "window":
                    driver.close()
                    driver.switch_to.window(janela_sistema)
                else:
                    driver.switch_to.default_content()

                click_results.append(f"Sucesso com {id_btn}")
                break

            except Exception as e:
                try:
                    driver.save_screenshot("erro_download.png")
                except Exception:
                    pass
                click_results.append(f"Falha ao usar {id_btn}: {e}")
                continue

        # Extração
        set_status("🔎 Lendo PDFs baixados e extraindo dados…")
        update_ui_status()
        df_final = processar_arquivos_baixados(download_dir, valor_solicitado, forcar_ocr=forcar_ocr)

        resp = {
            "status": "Sucesso",
            "dados": df_final,
            "diretorio": download_dir,
            "cliques": click_results
        }

        # Debug opcional: salvar page_source
        if debug:
            try:
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                resp["page_source"] = "page_source.html"
            except Exception:
                pass

        return resp

    except Exception as e:
        try:
            driver.save_screenshot("erro_download.png")
        except Exception:
            pass
        return {"erro": str(e)}
    finally:
        driver.quit()

# ============= UI STREAMLIT =============

st.set_page_config(page_title=APP_TITLE, page_icon=PAGE_ICON, layout="wide")
st.title(APP_TITLE)

if "credentials" not in st.secrets:
    st.error("Configure as credenciais em Secrets (credentials.usuario / credentials.senha).")
    st.stop()

with st.sidebar:
    st.header("⚙️ Configurações")
    headless = st.checkbox("Rodar headless (recomendado em servidor)", value=True)
    preferir_outras = st.checkbox("Preferir 'Outras Despesas' antes de 'Imprimir'", value=True)
    forcar_ocr = st.checkbox("Forçar OCR (mesmo que o PDF tenha texto)", value=False)
    debug = st.checkbox("Debug: salvar page_source.html", value=False)

guia = st.text_input("Número do Atendimento:", value="", help="Digite apenas números; caracteres não numéricos serão ignorados.")

# Status placeholder
st.session_state["status_container"] = st.empty()

if st.button("🚀 Processar e Analisar"):
    if not guia:
        st.warning("Informe a guia.")
    else:
        with st.spinner("Navegando no portal e baixando documentos..."):
            res = extrair_detalhes_site_amhp(
                guia,
                headless=headless,
                preferir_outras=preferir_outras,
                forcar_ocr=forcar_ocr,
                debug=debug
            )

        if "erro" in res:
            st.error(f"Erro: {res['erro']}")
            if os.path.exists("erro_download.png"):
                st.image("erro_download.png", caption="Screenshot do Erro")
        else:
            st.success("Automação concluída!")
            if "cliques" in res:
                with st.expander("🧪 Log de tentativas de clique (diagnóstico)"):
                    for item in res["cliques"]:
                        st.write("•", item)

            # Conferência de arquivos baixados
            with st.expander("📂 Conferência de Arquivos Baixados"):
                arquivos = sorted(os.listdir(res["diretorio"]))
                if arquivos:
                    for arq in arquivos:
                        caminho = os.path.join(res["diretorio"], arq)
                        tamanho_kb = os.path.getsize(caminho) / 1024
                        st.write(f"📄 {arq} ({tamanho_kb:.1f} KB)")
                        with open(caminho, "rb") as f:
                            st.download_button(f"📥 Baixar {arq}", f, file_name=arq)
                else:
                    st.warning("Nenhum arquivo encontrado na pasta de download.")

            # Exibição dos dados extraídos
            df = res["dados"]
            if not df.empty:
                st.subheader("📋 Dados Extraídos")
                st.dataframe(df, use_container_width=True)
                csv = df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "📥 Baixar Planilha de Resultados",
                    csv, "faturamento.csv", "text/csv"
                )
            else:
                st.info("Os arquivos foram baixados, mas a extração não encontrou o padrão esperado.\n"
                        "Sugestões: habilite 'Forçar OCR' e/ou ajuste a Regex no código.")

            # Debug
            if debug and os.path.exists("page_source.html"):
                with open("page_source.html", "rb") as f:
                    st.download_button("🪛 Baixar page_source.html (debug)", f, file_name="page_source.html")
