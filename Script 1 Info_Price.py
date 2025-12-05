import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import pyperclip
import os

# ================= CONFIGURAÇÕES =================
USUARIO = "SEU_EMAIL_AQUI"
SENHA = "SUA_SENHA_AQUI"
URL_LOGIN = "https://painel.infoprice.com.br/login"
PASTA_DOWNLOAD = r"C:\Downloads_InfoPrice"
ARQUIVO_CURVA_A = "C:\\Users\\Leonardo.Galdino\\Desktop\\scripts\\InfoPrice\\Curva A Compradores 01-01 á 25-11.xlsx"
# =================================================

# --- 1. PREPARAR OS CÓDIGOS ---
print(">>> 1. Preparando lista de produtos...")


def carregar_planilha_robusta(arquivo):
    try:
        return pd.read_csv(arquivo, sep=None, engine='python', encoding='latin1')
    except:
        pass
    try:
        return pd.read_csv(arquivo, sep=',', engine='python', encoding='utf-8')
    except:
        pass
    try:
        return pd.read_csv(arquivo, sep=';', engine='python', encoding='latin1')
    except:
        pass
    try:
        return pd.read_excel(arquivo)
    except:
        pass
    return None


try:
    df = carregar_planilha_robusta(ARQUIVO_CURVA_A)
    if df is None:
        print("❌ ERRO: Não leu o arquivo.")
        exit()

    if 'Código Acesso' not in df.columns:
        coluna_codigo = df.columns[4]
    else:
        coluna_codigo = 'Código Acesso'

    lista_eans = df[coluna_codigo].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique()
    lista_eans = [x for x in lista_eans if len(x) > 6 and x.isdigit()]

    texto_eans = ", ".join(lista_eans)
    pyperclip.copy(texto_eans)
    print(f"✓ SUCESSO! {len(lista_eans)} produtos carregados.")

except Exception as e:
    print(f"ERRO GERAL: {e}")
    exit()

# --- 2. INICIAR ROBÔ (MODO ANTI-PROXY) ---
print(">>> 2. Iniciando Microsoft Edge (Bypass Proxy)...")

options = EdgeOptions()
prefs = {"download.default_directory": PASTA_DOWNLOAD}
options.add_experimental_option("prefs", prefs)

# --- COMANDOS PARA DRIBLAR O BLOQUEIO DE REDE ---
options.add_argument("--no-proxy-server")  # <--- OBRIGA A NÃO USAR PROXY
options.add_argument("--remote-debugging-port=9222")  # <--- USA PORTA PADRÃO DE DEV
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--start-maximized")
options.add_argument("--ignore-certificate-errors")

# Verifica Driver
caminho_driver = "msedgedriver.exe"
if not os.path.exists(caminho_driver):
    print("❌ ERRO: O arquivo 'msedgedriver.exe' não está na pasta!")
    exit()

try:
    # Inicia o Serviço
    service = Service(executable_path=caminho_driver)
    driver = webdriver.Edge(service=service, options=options)

    print(">>> Navegador aberto! Acessando site...")
    driver.get(URL_LOGIN)

    # LOGIN
    print(">>> Fazendo Login...")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "email")))
    driver.find_element(By.NAME, "email").send_keys(USUARIO)
    driver.find_element(By.NAME, "password").send_keys(SENHA)

    btn_login = driver.find_element(By.XPATH, "//button[contains(text(), 'Entrar') or @type='submit']")
    btn_login.click()

    print(">>> Aguardando dashboard (15s)...")
    time.sleep(15)

    # --- 3. ABRIR FILTRO ---
    print(">>> 3. Abrindo filtro...")
    xpath_botao_filtro = '//*[@id="product-filter-desktop"]'

    try:
        driver.find_element(By.XPATH, xpath_botao_filtro).click()
    except:
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_botao_filtro))).click()

    time.sleep(2)

    # --- 4. COLAR CÓDIGOS ---
    print(">>> 4. Inserindo códigos...")
    xpath_input = '/html/body/div[9]/div[1]/input'

    try:
        campo_busca = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_input)))
        campo_busca.click()
        campo_busca.clear()
        campo_busca.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(5)
    except:
        print("⚠️ Falha na colagem automática. COLE MANUALMENTE (CTRL+V).")
        time.sleep(10)

    # --- 5. CHECKBOX ---
    print(">>> 5. Selecionando Checkbox...")
    xpath_checkbox = '/html/body/div[9]/span/div[2]/ul/li[1]/div/div/label/span'

    try:
        checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_checkbox)))
        checkbox.click()
        time.sleep(1)
        # Fechar modal
        try:
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except:
            pass
    except:
        print("⚠️ Checkbox não encontrado ou já fechado.")

    time.sleep(2)

    # --- 6. ATUALIZAR ---
    print(">>> 6. Atualizando resultados...")
    xpath_atualizar = '/html/body/div[3]/div/div/div[4]/div[1]/div[3]/button'

    try:
        driver.find_element(By.XPATH, xpath_atualizar).click()
    except:
        driver.execute_script("window.scrollBy(0, -200);")
        try:
            driver.find_element(By.XPATH, xpath_atualizar).click()
        except:
            print("⚠️ CLIQUE EM 'ATUALIZAR RESULTADOS' MANUALMENTE AGORA!")

    print(">>> Carregando tabela (25s)...")
    time.sleep(25)

    # --- 7. DOWNLOAD ---
    print(">>> 7. Baixando...")
    xpath_download = '//*[@id="download-btn"]/button'

    try:
        driver.find_element(By.XPATH, xpath_download).click()
        print(">>> Download iniciado! Aguardando 30s...")
        time.sleep(30)
    except:
        print("❌ Não achei botão de download. Tente baixar manualmente.")
        time.sleep(60)

except Exception as e:
    print(f"❌ ERRO: {e}")

finally:
    print(">>> Script finalizado.")
    # driver.quit()