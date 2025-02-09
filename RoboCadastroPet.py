from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Caminho do arquivo Excel contendo os dados dos animais
caminho_XLSX = "teste.xlsx"

# Carregar os dados da planilha, garantindo que algumas colunas sejam lidas como string
df = pd.read_excel(caminho_XLSX, dtype={"Cpf": str, "aniversario": str})

# Configuração do Selenium e WebDriver
caminho_drive = 'chromedriver.exe'
service = Service(caminho_drive)
navegador = webdriver.Chrome(service=service)

def login():
    """Realiza login no sistema."""
    navegador.maximize_window()
    print("Realizando login...")
    navegador.get("URL_SITE")

    # Aguardar campo de e-mail e inserir credenciais
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

def procura_cliente(row):
    """Busca o cliente pelo CPF no sistema."""
    try:
        print(f"Procurando cliente: {row['Cpf']}")

        barra_pesquisa = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.ID, "p__pes_var_nome"))
        )
        barra_pesquisa.clear()
        barra_pesquisa.send_keys(row["Cpf"])
        print(f"CPF inserido na pesquisa: {row['Cpf']}")

        botao_pesquisar = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.ID, "p__btn_filtrar"))
        )
        botao_pesquisar.click()
        print(f"Cliente pesquisado: {row['Cpf']}")

        # Aguardar carregamento da lista e clicar no cliente
        cliente = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//tr/td'))
        )
        cliente.click()
        print(f"Cliente encontrado e clicado: {row['Cpf']}")
        time.sleep(1)

    except Exception as e:
        raise Exception(f"Erro ao procurar cliente '{row['Cpf']}': {e}")

def adicionar_animal():
    """Clica no botão para adicionar um novo animal ao cliente."""
    try:
        botao_adicionar = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_add_animal"]'))
        )
        botao_adicionar.click()
        print("Botão 'Adicionar Animal' clicado.")
        time.sleep(1)

    except Exception as e:
        raise Exception(f"Erro ao adicionar animal: {e}")

def preencher_dados_animal(row):
    """Preenche os dados do animal no formulário."""
    try:
        print(f"Preenchendo dados do animal para o cliente: {row['Cpf']}")

        # Campo "Animal"
        campo_animal = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="ani_var_nome"]'))
        )
        campo_animal.clear()
        campo_animal.send_keys(row["nome"])
        print(f"Animal preenchido: {row['nome']}")

        # Campo "Sexo"
        campo_sexo = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ani_cha_sexo"]'))
        )
        campo_sexo.click()
        campo_sexo.send_keys(row["sexo"])
        campo_sexo.send_keys(Keys.RETURN)
        print(f"Sexo preenchido: {row['sexo']}")
        time.sleep(1)

        # Campo "Nascimento" (se disponível)
        if pd.notna(row["aniversario"]):
            data_formatada = pd.to_datetime(row["aniversario"]).strftime("%d/%m/%Y")
            
            campo_nascimento = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="ani_dat_nascimento"]'))
            )
            campo_nascimento.clear()

            for char in data_formatada:
                campo_nascimento.send_keys(char)
                time.sleep(0.2)  # Pequena pausa para simular digitação
            print(f"Nascimento preenchido: {data_formatada}")
        else:
            print("Aniversário não informado, campo deixado em branco.")

        # Campo "Espécie"
        seletor_especie = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="s2id_esp_int_codigo"]/a'))
        )
        seletor_especie.click()  # Clica para abrir o dropdown
        print("Dropdown de Espécie aberto.")

        input_especie = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="select2-drop"]/div/input'))
        )
        input_especie.send_keys(row["especie"])  # Digita a espécie
        input_especie.send_keys(Keys.RETURN)  # Pressiona Enter para selecionar
        print(f"Espécie preenchida: {row['especie']}")

        # Botão "Salvar"
        botao_salvar = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_salvar_ani"]'))
        )
        botao_salvar.click()
        print("Botão 'Salvar Animal' clicado.")

    except Exception as e:
        raise Exception(f"Erro ao preencher dados do animal para '{row['Cpf']}': {e}")

def botao_fecha_animal():
    """Fecha o formulário de cadastro de um animal."""
    try:
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="v__btn_fechar_ani"]'))
        ).click()
        print("Botão 'Fechar Animal' clicado.")
    except Exception as e:
        raise Exception("Erro ao clicar no botão 'Fechar Animal':", e)

def botao_fecha_cadastro():
    """Fecha o cadastro completo do cliente."""
    try:
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="v__btn_fechar_pro"]'))
        ).click()
        print("Botão 'Fechar Cadastro' clicado.")
    except Exception as e:
        raise Exception("Erro ao clicar no botão 'Fechar Cadastro':", e)

# Processo principal de cadastro de animais
try:
    login()
    for index, row in df.iterrows():
        try:
            print(f"Processando registro {index + 1}/{len(df)}: {row['Cpf']}")
            procura_cliente(row)
            adicionar_animal()
            preencher_dados_animal(row)
            botao_fecha_animal()
            time.sleep(0.5)

            # Verificar se o próximo CPF é o mesmo
            if index + 1 < len(df) and df.iloc[index + 1]["Cpf"] == row["Cpf"]:
                print(f"Mesmo CPF {row['Cpf']} encontrado, adicionando novo animal.")
            else:
                botao_fecha_cadastro()
                print(f"Cadastro do CPF {row['Cpf']} finalizado.")

        except Exception as registro_erro:
            print(f"Erro no registro {index + 1} (CPF {row['Cpf']}): {registro_erro}")
            continue

except Exception as e:
    print(f"Erro geral: {e}")
finally:
    print("Execução finalizada.")
    navegador.quit()
