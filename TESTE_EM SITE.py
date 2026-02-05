# --------------------------------------------------------------------------------
# Copyright (c) 2026 Luiz Tiago da Silva - PROGEN S.A.
# Descrição: Robô PRÉ-APR PRO - Versão WEB Cloud com Fix de Conexão e Agente
# --------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import time
import os
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Progen S.A. - Portal APR", layout="wide")

# Estilo para o Log Rosa Profissional e Botões
st.markdown("""
    <style>
    .log-container {
        background-color: #ffffff;
        color: #FF007F;
        font-family: 'Consolas', monospace;
        padding: 15px;
        border: 2px solid #f0f2f6;
        border-radius: 10px;
        height: 450px;
        overflow-y: auto;
        white-space: pre-wrap;
        font-weight: bold;
        font-size: 13px;
    }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def calcular_validade(data_texto):
    try:
        data_obj = datetime.strptime(data_texto.split(" ")[0], "%d/%m/%Y")
        return data_obj + timedelta(days=45)
    except: return None

# --- SIDEBAR (CONFIGURAÇÕES E BOTÃO INICIAR) ---
with st.sidebar:
    st.image("https://progen.com.br/wp-content/uploads/2021/05/logo-progen.png", width=180)
    st.title("⚙️ Configuração")
    
    reg = st.text_input("Registro (Usuário)")
    senha = st.text_input("Senha") 
    loc = st.selectbox("Localidade", ["Campos", "Lagos", "Macaé", "Magé", "Niterói", "São Gonçalo", "Serrana", "Sul", "Noroeste"], index=4)
    
    st.divider()
    btn_iniciar = st.button("🚀 INICIAR PROCESSAMENTO")
    
    st.divider()
    uploaded_file = st.file_uploader("Upload da Planilha (.xlsx)", type=["xlsx"])

# --- ÁREA PRINCIPAL ---
st.title("🤖 Robô Pré-APR PRO - Web")
st.write(f"Status atual: **Aguardando comando para {loc}**")

log_placeholder = st.empty()
log_history = []

def update_log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_history.append(f"[{timestamp}] {msg}")
    log_html = f'<div class="log-container">{"<br>".join(log_history)}</div>'
    log_placeholder.markdown(log_html, unsafe_allow_html=True)

# --- EXECUÇÃO DO ROBÔ ---
if btn_iniciar:
    if not reg or not senha or not uploaded_file:
        st.error("⚠️ Erro: Preencha Registro, Senha e suba a Planilha antes de iniciar.")
    else:
        # Configurações do Navegador (Disfarçado para evitar bloqueios de rede)
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        # Identificando-se como um navegador Windows real
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
        chrome_options.binary_location = "/usr/bin/chromium" 

        try:
            update_log("Iniciando motor do navegador no servidor...")
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 45) # Aumentei o tempo de espera para rede lenta
            actions = ActionChains(driver)

            # Lendo planilha
            df_input = pd.read_excel(uploaded_file)
            novos_resultados = []
            hoje = datetime.now()

            # URL DIRETA DE LOGIN
            url_site = "https://intapr.enel.com/login"
            update_log(f"Conectando ao Portal Enel ({url_site})...")
            driver.get(url_site)
            
            time.sleep(8) # Espera para estabilizar a conexão
            update_log("Realizando login no sistema...")
            
            # LOGIN
            wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys(reg)
            driver.find_element(By.ID, "password").send_keys(senha)
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[@type='submit']"))
            
            time.sleep(15) 
            update_log(f"Filtros: {loc}")

            # SELEÇÃO DE LOCALIDADE
            def selecionar_web(id_campo, texto):
                campo = wait.until(EC.element_to_be_clickable((By.ID, id_campo)))
                actions.move_to_element(campo).click().perform()
                time.sleep(2)
                actions.send_keys(texto).perform()
                time.sleep(2)
                actions.send_keys(Keys.ENTER).perform()
                time.sleep(1)
                actions.send_keys(Keys.ESCAPE).perform()

            selecionar_web("select-state", "Rio de Janeiro")
            selecionar_web("select-distribuidora", loc)

            # PROCESSAMENTO
            col_nome = next((c for c in df_input.columns if "PROJETOS.PY" in str(c).upper()), None)
            
            if col_nome:
                projetos = df_input[col_nome].dropna().unique()
                for projeto in projetos:
                    proj_id = str(projeto).strip().replace(".0", "").replace(".", "").zfill(10)
                    update_log(f"Consultando Projeto: {proj_id}")

                    try:
                        campo_busca = wait.until(EC.element_to_be_clickable((By.ID, "standard-basic")))
                        campo_busca.send_keys(Keys.CONTROL + "a", Keys.DELETE)
                        campo_busca.send_keys(proj_id)
                        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'Filtrar')]"))
                        time.sleep(7)

                        linhas = driver.find_elements(By.CSS_SELECTOR, "div.MuiDataGrid-row")
                        for linha in linhas:
                            def get_js(f): return driver.execute_script(f"return arguments[0].querySelector('[data-field=\"{f}\"]').innerText;", linha).strip()
                            val_obj = calcular_validade(get_js("createdAt"))
                            id_limpo = get_js("id").replace(".", "")

                            if val_obj and val_obj > hoje:
                                update_log(f"   ID {id_limpo}: ATIVO (EXPIRA: {val_obj.strftime('%d/%m/%Y')})")
                                novos_resultados.append({
                                    "Ordem": get_js("orderNumber"), "ID_tabela": id_limpo,
                                    "Empresa": get_js("companyName"), "Data_Criacao": get_js("createdAt"),
                                    "Status": get_js("statusName"), "Validade_APR": val_obj.strftime("%d/%m/%Y"),
                                    "Retirado_em": datetime.now().strftime("%d/%m/%Y %H:%M")
                                })
                    except: continue

            # FINALIZAÇÃO
            if novos_resultados:
                update_log(f"Fim: {len(novos_resultados)} ativos.")
                df_final = pd.DataFrame(novos_resultados)
                
                towrite = BytesIO()
                df_final.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                
                st.balloons()
                st.download_button(
                    label="✅ BAIXAR PLANILHA",
                    data=towrite,
                    file_name=f"APR_{loc}_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                update_log("Nenhum ativo encontrado.")

        except Exception as e:
            st.error(f"Erro Crítico: {e}")
            update_log(f"ERRO DE CONEXÃO: Verifique se o portal está acessível.")
        finally:
            if 'driver' in locals():
                driver.quit()
