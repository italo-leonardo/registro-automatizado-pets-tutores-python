from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Caminho para o arquivo XLSX contendo os dados dos clientes
caminho_XLSX = "testep.xlsx"

# Ler arquivo Excel com Pandas, garantindo que algumas colunas sejam tratadas como strings
df = pd.read_excel(caminho_XLSX, dtype={"RG": str, "CEP": str})

# Configuração do Selenium e do WebDriver
caminho_drive = 'chromedriver.exe'
service = Service(caminho_drive)
navegador = webdriver.Chrome(service=service)

def login():
    """Realiza o login no sistema."""
    navegador.maximize_window()
    print("Realizando login...")
    navegador.get("URL_SITE")

    # Esperar pelo campo de e-mail e inserir credenciais
    WebDriverWait(navegador, 60).until(
        EC.presence_of_element_located((By.NAME, "l_usu_var_email"))
    ).send_keys("SEU_EMAIL")

    # Inserir senha e pressionar Enter para login
    navegador.find_element(By.NAME, "l_usu_var_senha").send_keys("SUA_SENHA", Keys.RETURN)

    # Esperar o carregamento da página principal
    WebDriverWait(navegador, 60).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    WebDriverWait(navegador, 60).until(EC.presence_of_element_located((By.ID, "p__btn_adicionar")))
    print("Login realizado com sucesso.")

def clicar_botao_adicionar():
    """Clica no botão 'Adicionar Cliente'."""
    print("Clicando no botão Adicionar...")
    WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.ID, "p__btn_adicionar"))
    ).click()
    print("Botão Adicionar clicado.")

def preencher_formulario(row):
    """Preenche os dados básicos do formulário do cliente."""
    try:
        print(f"Preenchendo formulário para o cliente: {row['Nome']}")

        # Campos do formulário e seus valores correspondentes
        campos = {
            '//*[@id="pes_var_nome"]': row["Nome"],
            '//*[@id="pes_var_cpf"]': str(row["CPF"]).zfill(11),
            '//*[@id="pes_var_rg"]': row["RG"],
            '//input[contains(@name, "con_var_contato[]")]': row["Celular"]
        }

        # Preenchimento de cada campo com pequenas pausas para evitar erros
        for xpath, valor in campos.items():
            campo = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            campo.clear()
            for char in str(valor):
                campo.send_keys(char)
                time.sleep(0.1)

        # Preencher Sexo e WhatsApp
        preencher_sexo(row)
        selecionar_whatsapp(row)

        print(f"Formulário preenchido para o cliente: {row['Nome']}")
    except Exception as e:
        print(f"Erro ao preencher formulário do cliente '{row['Nome']}': {e}")

def preencher_sexo(row):
    """Seleciona o sexo do cliente no formulário."""
    try:
        print(f"Selecionando o Sexo para o cliente: {row['Nome']}")
        campo_sexo = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="pes_cha_sexo"]'))
        )
        campo_sexo.send_keys(row["Sexo"], Keys.RETURN)
    except Exception as e:
        print(f"Erro ao preencher o campo Sexo para '{row['Nome']}': {e}")

def selecionar_whatsapp(row):
    """Define se o telefone é WhatsApp ou não."""
    try:
        print(f"Definindo WhatsApp para o cliente: {row['Nome']}")
        campo_whatsapp = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//select[contains(@name, "con_cha_whatsapp[]")]'))
        )
        campo_whatsapp.send_keys("É WhatsApp" if row["WhatsApp"] == "S" else "Não é WhatsApp")
    except Exception as e:
        print(f"Erro ao definir WhatsApp para '{row['Nome']}': {e}")

def clicar_aba_endereco():
    """Acessa a aba de Endereço no formulário."""
    try:
        print("Acessando aba Endereço...")
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@href="#tabEndereco"]'))
        ).click()
        print("Aba Endereço acessada.")
    except Exception as e:
        print(f"Erro ao acessar a aba Endereço: {e}")

def preencher_campos_endereco(row):
    """Preenche os campos de endereço do cliente."""
    try:
        print(f"Preenchendo os campos de endereço para {row['Nome']}")

        # Campos do endereço e seus valores correspondentes
        endereco_campos = {
            "end_var_cep": row["CEP"],
            "end_var_endereco": row["Endereco"],
            "end_var_numero": row["Numero"],
            "end_var_complemento": row["Complemento"],
            "end_var_bairro": row["Bairro"],
            "end_var_municipio": row["Cidade"]
        }

        # Preenchimento de cada campo
        for campo_id, valor in endereco_campos.items():
            campo = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.ID, campo_id))
            )
            campo.clear()
            campo.send_keys(valor)
        
        print(f"Endereço preenchido para {row['Nome']}")
    except Exception as e:
        print(f"Erro ao preencher endereço para '{row['Nome']}': {e}")

def clicar_botao_salvar():
    """Clica no botão Salvar para cadastrar o cliente."""
    try:
        print("Salvando formulário...")
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="f__btn_salvar"]'))
        ).click()
        print("Formulário salvo com sucesso.")
    except Exception as e:
        print(f"Erro ao clicar no botão Salvar: {e}")

def clicar_botao_fechar():
    """Fecha a tela de cadastro do cliente."""
    try:
        print("Fechando tela de cadastro...")
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="v__btn_fechar_topo"]'))
        ).click()
        print("Tela de cadastro fechada.")
    except Exception as e:
        print(f"Erro ao clicar no botão Fechar: {e}")

def processar_cadastro():
    """Executa o fluxo completo de cadastro para cada cliente na planilha."""
    for _, row in df.iterrows():
        try:
            clicar_botao_adicionar()
            preencher_formulario(row)
            clicar_aba_endereco()
            preencher_campos_endereco(row)
            clicar_botao_salvar()
            clicar_botao_fechar()
        except Exception as e:
            print(f"Erro ao processar cadastro de '{row['Nome']}': {e}")
            continue

# Fluxo principal de execução
try:
    login()
    processar_cadastro()
except Exception as e:
    print(f"Erro geral: {e}")
finally:
    print("Execução finalizada.")
    navegador.quit()
