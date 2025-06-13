# ==============================================================================
# IMPORTAÇÃO DAS BIBLIOTECAS
# ==============================================================================
import os
import json
import pandas as pd
import gspread
import smtplib
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ==============================================================================
# CARREGAMENTO DAS CONFIGURAÇÕES E CREDENCIAIS (VIA GITHUB SECRETS)
# ==============================================================================
print("[ROBÔ] Carregando configurações e credenciais...")
# Credenciais da Intranet
usuario_intranet = os.getenv('INTRANET_USER')
senha_intranet = os.getenv('INTRANET_PASSWORD')

# Credenciais do Google
url_da_planilha = os.getenv('GOOGLE_SHEET_URL')
email_remetente = os.getenv('GMAIL_USER')
senha_app_gmail = os.getenv('GMAIL_APP_PASSWORD')
google_credentials_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
google_credentials_dict = json.loads(google_credentials_json)

# Autenticação para Google Sheets e Drive
gc = gspread.service_account_from_dict(google_credentials_dict)
drive_service = build('drive', 'v3', credentials=gc.auth)
print("[ROBÔ] Credenciais carregadas.")

# ==============================================================================
# ATENÇÃO: PERSONALIZE AS INFORMAÇÕES DA INTRANET AQUI
# Use a ferramenta "Inspecionar" (F12) do seu navegador para encontrar os IDs
# ==============================================================================
url_login = "https://intranet.beneficiofacil.com.br/moumo3.asp"
id_campo_usuario = "usuario"      # Troque "usuario" pelo ID real do campo de usuário
id_campo_senha = "senha"          # Troque "senha" pelo ID real do campo de senha
id_botao_login = "btnEntrar"      # Troque "btnEntrar" pelo ID real do botão de login
url_pagina_da_tabela = "https://intranet.beneficiofacil.com.br/outra_pagina.asp" # URL da página após o login
id_da_tabela = "id_da_sua_tabela" # Troque pelo ID real da tabela que você quer copiar

# ==============================================================================
# FUNÇÕES AUXILIARES (GOOGLE DRIVE E GMAIL)
# ==============================================================================

def encontrar_ou_criar_pasta(nome_da_pasta, id_da_pasta_pai=None):
    """Busca uma pasta pelo nome no Drive. Se não encontrar, cria uma nova."""
    query = f"name='{nome_da_pasta}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    if response.get('files', []):
        folder_id = response.get('files', [])[0].get('id')
        print(f"[DRIVE] Pasta '{nome_da_pasta}' encontrada com ID: {folder_id}")
        return folder_id
    else:
        file_metadata = {'name': nome_da_pasta, 'mimeType': 'application/vnd.google-apps.folder'}
        if id_da_pasta_pai: file_metadata['parents'] = [id_da_pasta_pai]
        folder = drive_service.files().create(body=file_metadata, fields='id').execute()
        folder_id = folder.get('id')
        print(f"[DRIVE] Pasta '{nome_da_pasta}' criada com ID: {folder_id}")
        return folder_id

def mover_arquivo_para_pasta(id_do_arquivo, id_da_pasta_destino):
    """Move um arquivo para uma nova pasta no Drive."""
    try:
        file = drive_service.files().get(fileId=id_do_arquivo, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))
        drive_service.files().update(fileId=id_do_arquivo, addParents=id_da_pasta_destino, removeParents=previous_parents, fields='id, parents').execute()
        print(f"[DRIVE] Arquivo movido com sucesso.")
    except HttpError as error:
        print(f"[DRIVE] Erro ao mover o arquivo: {error}")

def enviar_email_confirmacao(sucesso=True, erro_msg=""):
    """Envia um e-mail de confirmação (sucesso ou falha)."""
    if not email_remetente or not senha_app_gmail:
        print("[GMAIL] Credenciais de e-mail não configuradas. Pulando envio.")
        return

    msg = EmailMessage()
    if sucesso:
        msg['Subject'] = "✅ Automação de Relatórios Concluída com Sucesso"
        msg.set_content(f"Olá William,\n\nO robô executou a rotina de atualização da planilha com sucesso.\n\nData: {pd.Timestamp.now(tz='America/Sao_Paulo').strftime('%d/%m/%Y %H:%M')}")
    else:
        msg['Subject'] = "❌ FALHA na Automação de Relatórios"
        msg.set_content(f"Olá William,\n\nO robô encontrou um erro ao tentar executar a automação.\n\nErro: {erro_msg}\n\nData: {pd.Timestamp.now(tz='America/Sao_Paulo').strftime('%d/%m/%Y %H:%M')}")

    msg['From'] = email_remetente
    msg['To'] = email_remetente # Enviando para você mesmo
    try:
        print("[GMAIL] Enviando e-mail de notificação...")
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(email_remetente, senha_app_gmail)
            smtp.send_message(msg)
        print("[GMAIL] E-mail enviado com sucesso!")
    except Exception as e:
        print(f"[GMAIL] Erro ao enviar o e-mail: {e}")

# ==============================================================================
# LÓGICA PRINCIPAL DO ROBÔ
# ==============================================================================
print("[ROBÔ] Iniciando a execução principal...")
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
service = Service()
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.implicitly_wait(10) # Espera implícita para todos os elementos

try:
    # 1. LOGIN
    print(f"[SELENIUM] Acessando {url_login}...")
    driver.get(url_login)
    driver.find_element(By.ID, id_campo_usuario).send_keys(usuario_intranet)
    driver.find_element(By.ID, id_campo_senha).send_keys(senha_intranet)
    driver.find_element(By.ID, id_botao_login).click()
    print("[SELENIUM] Login realizado com sucesso.")

    # 2. EXTRAÇÃO DA TABELA
    print(f"[SELENIUM] Navegando para a página da tabela...")
    driver.get(url_pagina_da_tabela)
    print("[SELENIUM] Extraindo tabela HTML...")
    tabela_html = driver.find_element(By.ID, id_da_tabela).get_attribute('outerHTML')
    df = pd.read_html(tabela_html)[0]
    print("[PANDAS] Tabela extraída e convertida para DataFrame:")
    print(df.head())

    # 3. ATUALIZAÇÃO DO GOOGLE SHEETS
    print("[GSHEETS] Conectando ao Google Sheets...")
    spreadsheet = gc.open_by_url(url_da_planilha)
    worksheet = spreadsheet.worksheet("Página1") # Mude o nome da aba se for diferente
    print("[GSHEETS] Limpando dados antigos e inserindo novos...")
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    print("[GSHEETS] Planilha atualizada com sucesso.")

    # 4. ORGANIZAÇÃO NO GOOGLE DRIVE
    id_pasta_principal = encontrar_ou_criar_pasta("Robô-Automação")
    mover_arquivo_para_pasta(spreadsheet.id, id_pasta_principal)

    # 5. NOTIFICAÇÃO POR E-MAIL
    enviar_email_confirmacao(sucesso=True)

except Exception as e:
    print(f"[ERRO] Ocorreu uma falha grave na automação: {e}")
    enviar_email_confirmacao(sucesso=False, erro_msg=str(e))

finally:
    driver.quit()
    print("[ROBÔ] Navegador fechado. Fim da execução.")
