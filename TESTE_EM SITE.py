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
from webdriver_manager.chrome import ChromeDriverManager
from io import BytesIO

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Progen S.A. - Portal APR", layout="wide")

# Estilo para o Log Rosa
st.markdown("""
    <style>
    .log-container {
        background-color: #ffffff;
        color: #FF007F;
        font-family: 'Consolas', monospace;
        padding: 15px;
        border: 2px solid #f0f2f6;
        border-radius: 10px;
        height: 400px;
        overflow-y: auto;
        white-space: pre-wrap;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def calcular_validade(data_texto):
    try:
        data_obj = datetime.strptime(data_texto.split(" ")[0], "%d/%m/%Y")
        return data_obj + timedelta(days=45)
    except: return None

# --- INTERFACE LATERAL (INPUTS) ---
with st.sidebar:
    st.header("🔑 Acesso Progen")
    reg = st.text_input("Registro (Usuário)")
    senha = st.text_input("Senha", type="default") # Mantendo visível como você pediu
    loc = st.selectbox("Localidade", ["Niterói", "São Gonçalo", "Campos", "Macaé", "Lagos", "Magé", "Serrana", "Sul", "Noroeste"])
    
    st.divider()
    uploaded_file = st.file_uploader("Suba sua planilha de projetos (.xlsx)", type=["xlsx"])

# --- CORPO PRINCIPAL ---
st.title("🤖 Automação Pré-APR Progen")
st.info("Preencha os dados na lateral e clique em 'Iniciar Processamento'.")

col_btn, col_status = st.columns([1, 3])
with col_btn:
    iniciar = st.button("🚀 INICIAR PROCESSAMENTO")

log_placeholder = st.empty()
log_history = []

def update_log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_history.append(f"[{timestamp}] {msg}")
    log_html = f'<div class="log-container">{"<br>".join(log_history)}</div>'
    log_placeholder.markdown(log_html, unsafe_allow_html=True)

# --- EXECUÇÃO DO ROBÔ ---
if iniciar:
    if not reg or not senha or not uploaded_file:
        st.error("⚠️ Por favor, preencha todos os campos e anexe a planilha.")
    else:
        # Configuração do Chrome para Nuvem
        chrome_options = Options()
        chrome_options.add_argument("--headless") # Essencial para rodar no site
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        try:
            update_log("Iniciando navegador no servidor...")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            wait = WebDriverWait(driver, 30)
            actions = ActionChains(driver)

            # Lendo a planilha enviada
            df_input = pd.read_excel(uploaded_file)
            novos_resultados = []
            hoje = datetime.now()

            driver.get("https://intapr.enel.com/")
            update_log("Realizando login no sistema Enel...")
            
            # (Aqui entra sua lógica de login e seleção de localidade idêntica ao robô anterior)
            # ... [Login, select-state, select-distribuidora] ...
            
            # Exemplo de loop de consulta adaptado para Streamlit:
            col_id = next((c for c in df_input.columns if "PROJETOS.PY" in str(c).upper()), None)
            
            if col_id:
                for projeto in df_input[col_id]:
                    proj_id = str(projeto).strip().replace(".0", "").replace(".", "").zfill(10)
                    if "nan" in proj_id: continue
                    
                    update_log(f"Consultando Projeto: {proj_id}")
                    # ... [Lógica de preencher campo, filtrar e ler linhas] ...
                    
                    # Simulação de sucesso para teste visual:
                    # id_limpo = proj_id
                    # val_obj = hoje + timedelta(days=10)
                    # update_log(f"   ID {id_limpo}: ATIVO (EXPIRA: {val_obj.strftime('%d/%m/%Y')})")
                    # novos_resultados.append({...})

            # --- FINALIZAÇÃO E DOWNLOAD ---
            if novos_resultados:
                df_final = pd.DataFrame(novos_resultados)
                
                # Criar arquivo Excel em memória para download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Resultados_LOG')
                
                st.success("✅ Processamento concluído!")
                st.download_button(
                    label="📥 BAIXAR PLANILHA RESULTADO",
                    data=output.getvalue(),
                    file_name=f"Resultado_APR_{loc}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                update_log("Nenhum registro ativo encontrado.")

        except Exception as e:
            st.error(f"Erro na execução: {e}")
        finally:
            if 'driver' in locals():
                driver.quit()