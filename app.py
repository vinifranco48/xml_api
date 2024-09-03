import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import xml.etree.ElementTree as ET
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
from dotenv import load_dotenv
import os
# Configurações da página Streamlit
st.set_page_config(
    page_title="Processador de NFe",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="auto"
)
load_dotenv()
GOOGLE_PROJECT_ID = os.getenv("GOOGLE_PROJECT_ID")
GOOGLE_PRIVATE_KEY_ID = os.getenv("GOOGLE_PRIVATE_KEY_ID")
GOOGLE_PRIVATE_KEY = os.getenv("GOOGLE_PRIVATE_KEY")
GOOGLE_CLIENT_EMAIL = os.getenv("GOOGLE_CLIENT_EMAIL")
GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_CERT_URL = os.getenv("GOOGLE_CLIENT_CERT_URL")
# Função para autenticar e conectar ao Google Sheets
@st.cache_resource
def connect_to_gsheet(sheet_name):
    credentials_info = {
        "type": "service_account",
        "project_id": GOOGLE_PROJECT_ID,
        "private_key_id": GOOGLE_PRIVATE_KEY_ID,
        "private_key": GOOGLE_PRIVATE_KEY,
        "client_email": GOOGLE_CLIENT_EMAIL,
        "client_id": GOOGLE_CLIENT_ID,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": GOOGLE_CLIENT_CERT_URL,
    }
    credentials = Credentials.from_service_account_info(credentials_info, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ])
    client = gspread.authorize(credentials)
    
    try:
        sheet = client.open(sheet_name).sheet1
        return sheet
    except Exception as e:
        st.error(f"Erro ao acessar a planilha: {e}")
        return None

def criar_colunas(sheet):
    colunas = [
        'Data', 'Placa', 'Modelo', 'KM', 'Nota', 'Fornecedor', 'Tipo de Serviço', 'Item da Nota', 'Valor', 'Quantidade', 'Peça', 'Observações'
    ]
    existing_headers = sheet.row_values(1)
    if not existing_headers or set(colunas) != set(existing_headers):
        sheet.insert_row(colunas, 1)
        st.success('Cabeçalhos criados na planilha!')

def recuperar_fornecedores(sheet):
    try:
        fornecedores = sheet.col_values(6)[1:]  # Coluna "Fornecedor", ignorando o cabeçalho
        st.session_state['fornecedores'] = list(set(fornecedores))  # Remover duplicatas
    except Exception as e:
        st.error(f"Erro ao recuperar fornecedores: {e}")

def initialize_session_state():
    if 'carros' not in st.session_state:
        st.session_state['carros'] = [
            {'placa': 'PSK9760', 'modelo': 'S10 ESTREITO'},
            {'placa': 'PTA8229', 'modelo': 'S10'},
            {'placa': 'PTP4215', 'modelo': 'CAMINHÃO VOLVO'},
            {'placa': 'PTQ9932', 'modelo': 'HILUX'},
            {'placa': 'PTS8I32', 'modelo': 'S10'},
            {'placa': 'ROC0A68', 'modelo': 'SAVEIRO'},
            {'placa': 'SNJ8I23', 'modelo': 'FIAT TORO'},
            {'placa': 'SMP1C48', 'modelo': 'CAMINHÃO SPRINTER'}
        ]
    
    if 'fornecedores' not in st.session_state:
        st.session_state['fornecedores'] = []

def setup_driver(download_dir):
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    return webdriver.Chrome(options=options)

def download_xml(driver, url, chave_de_acesso):
    try:
        driver.get(url)
        
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Digite a CHAVE DE ACESSO']"))
        )
        input_field.send_keys(chave_de_acesso)
        
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Buscar DANFE/XML')]"))
        )
        search_button.click()
        
        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'bg-contrast-blue') and contains(text(), 'Baixar XML')]"))
        )
        download_button.click()
        
        # Esperar pelo download
        time.sleep(10)  # Ajuste conforme necessário
        
        return True
    except Exception as e:
        st.error(f"Erro ao baixar XML: {e}")
        return False

def process_nfe_xml(xml_content):
    root = ET.fromstring(xml_content)
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    
    nfe_data = {
        'numero_nota': root.find('.//nfe:nNF', ns).text,
        'chave_acesso': root.find('.//nfe:chNFe', ns).text,
        'data_emissao': datetime.strptime(root.find('.//nfe:dhEmi', ns).text.split('T')[0], '%Y-%m-%d').strftime('%d/%m/%Y'),
        'fornecedor': root.find('.//nfe:emit/nfe:xNome', ns).text,
        'cnpj_fornecedor': root.find('.//nfe:emit/nfe:CNPJ', ns).text,
        'valor_total': float(root.find('.//nfe:vNF', ns).text),
        'itens': []
    }
    
    for det in root.findall('.//nfe:det', ns):
        item = {
            'numero': det.get('nItem'),
            'descricao': det.find('.//nfe:xProd', ns).text,
            'quantidade': float(det.find('.//nfe:qCom', ns).text),
            'unidade': det.find('.//nfe:uCom', ns).text,
            'valor_unitario': float(det.find('.//nfe:vUnCom', ns).text),
            'valor_total': float(det.find('.//nfe:vProd', ns).text),
            'ncm': det.find('.//nfe:NCM', ns).text
        }
        nfe_data['itens'].append(item)
    
    return nfe_data

def registrar_nota_from_nfe(sheet, nfe_data, placa_carro, km):
    modelo_carro = next(carro['modelo'] for carro in st.session_state['carros'] if carro['placa'] == placa_carro)
    
    for item in nfe_data['itens']:
        registro = [
            nfe_data['data_emissao'],
            placa_carro,
            modelo_carro,
            km,
            nfe_data['numero_nota'],
            nfe_data['fornecedor'],
            'Peça/Serviço',  # Você pode ajustar isso conforme necessário
            item['descricao'],
            item['valor_unitario'],
            item['quantidade'],
            item['descricao'],  # Usando a descrição como nome da peça
            f"NCM: {item['ncm']}"  # Usando o NCM como observação
        ]
        sheet.append_row(registro)
    
    st.success(f"Nota {nfe_data['numero_nota']} registrada com sucesso!")

def main():
    st.title('Processador de NFe e Registro de Serviços')

    initialize_session_state()

    sheet = connect_to_gsheet('Controle_Frota')
    if sheet:
        criar_colunas(sheet)
        recuperar_fornecedores(sheet)
        
        chave_acesso = st.text_input("Insira a chave de acesso da NFe")
        placa_carro = st.selectbox('Placa do Carro', [carro['placa'] for carro in st.session_state['carros']])
        km = st.number_input('Quilometragem', min_value=0)
        
        if st.button("Processar NFe e Registrar"):
            if chave_acesso and placa_carro:
                download_dir = os.path.join(os.getcwd(), "downloads")
                os.makedirs(download_dir, exist_ok=True)
                
                driver = setup_driver(download_dir)
                try:
                    if download_xml(driver, "https://meudanfe.com.br/", chave_acesso):
                        xml_file = os.path.join(download_dir, f"{chave_acesso}.xml")
                        if os.path.exists(xml_file):
                            with open(xml_file, 'r', encoding='utf-8') as file:
                                xml_content = file.read()
                            nfe_data = process_nfe_xml(xml_content)
                            if nfe_data:
                                st.write("Dados da NFe:")
                                st.json(nfe_data)
                                registrar_nota_from_nfe(sheet, nfe_data, placa_carro, km)
                            else:
                                st.error("Falha ao processar o XML da NFe.")
                        else:
                            st.error("Arquivo XML não encontrado após o download.")
                    else:
                        st.error("Falha ao baixar o XML da NFe.")
                finally:
                    driver.quit()
            else:
                st.error("Por favor, insira a chave de acesso da NFe e selecione a placa do carro.")

if __name__ == "__main__":
    print("GOOGLE_PROJECT_ID:", GOOGLE_PROJECT_ID)
    print("GOOGLE_PRIVATE_KEY_ID:", GOOGLE_PRIVATE_KEY_ID)
    print("GOOGLE_PRIVATE_KEY:", GOOGLE_PRIVATE_KEY)
    print("GOOGLE_CLIENT_EMAIL:", GOOGLE_CLIENT_EMAIL)
    print("GOOGLE_CLIENT_ID:", GOOGLE_CLIENT_ID)
    print("GOOGLE_CLIENT_CERT_URL:", GOOGLE_CLIENT_CERT_URL)
    main()