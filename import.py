import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
from dotenv import load_dotenv
import os
import keyboard  # Precisamos instalar: pip install keyboard
import pyautogui
import warnings

# Variável global para controlar pausa
pausado = False

# Variável global para controlar o modo
IS_HEADLESS = False

def toggle_pausa(e):
    global pausado
    pausado = not pausado
    print("Script " + ("PAUSADO" if pausado else "DESPAUSADO"))

# Carregar variáveis do arquivo .env
load_dotenv()

def preencher_faixas_cep():
    # Suprimir mensagens de warning
    warnings.filterwarnings("ignore")
    
    modo = input("Escolha o modo de execução (1 para normal, 2 para minimizado): ")
    
    # Configurações do Chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-software-rasterizer')
    chrome_options.add_argument('--disable-features=VizDisplayCompositor')
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    if modo == "2":
        print("Executando em modo minimizado...")
        chrome_options.add_argument('--window-state=minimized')
    else:
        print("Executando em modo normal...")
        chrome_options.add_argument('--start-maximized')
    
    driver = webdriver.Chrome(options=chrome_options)
    
    # Registrar o atalho de teclado 'p' para pausar/despausar
    keyboard.on_press_key('p', toggle_pausa)
    
    # Pegar credenciais do arquivo .env
    email = os.getenv('LI_EMAIL')
    senha = os.getenv('LI_SENHA')
    
    # Caminho do arquivo Excel (corrigido)
    caminho_arquivo = r'C:\Users\natal\OneDrive\Área de Trabalho\Thalles\Preencer-Motoboy-Li\ABRANGENCIA__TRANSMOTO_RJ_09122024_173376087486006.xls'
    
    # Verificar se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        print(f"Erro: O arquivo não foi encontrado em: {caminho_arquivo}")
        return
        
    # Ler a planilha Excel
    df = pd.read_excel(caminho_arquivo)
    
    # Maximizar a janela do navegador
    driver.maximize_window()
    
    # Ir para página correta de login
    try:
        driver.get('https://app.lojaintegrada.com.br/painel')
    except TimeoutException:
        print("Timeout ao carregar página. Tentando refresh...")
        driver.refresh()
        time.sleep(5)
    
    # Fazer login
    wait = WebDriverWait(driver, 20)
    
    try:
        # Login...
        try:
            email_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
        except:
            email_field = wait.until(EC.presence_of_element_located((By.NAME, "email")))
        email_field.send_keys(email)
        
        try:
            senha_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
        except:
            senha_field = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        senha_field.send_keys(senha)
        
        try:
            login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
        except:
            try:
                login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entrar')]")))
            except:
                login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn-login")))
        
        login_btn.click()
        time.sleep(10)
        
        # Clicar no botão Aceitar após o login
        try:
            aceitar_btn = wait.until(EC.element_to_be_clickable((By.ID, "hs-eu-confirmation-button")))
            aceitar_btn.click()
            time.sleep(2)
        except Exception as e:
            print("Botão Aceitar não encontrado ou já clicado:", str(e))
        
        # Vai para a página de faixas de CEP
        try:
            driver.get('https://app.lojaintegrada.com.br/painel/configuracao/envio/194/motoboy')
        except TimeoutException:
            print("Timeout ao carregar página de motoboy. Tentando refresh...")
            driver.refresh()
            time.sleep(5)
        time.sleep(5)
        
        # Rolar a página uma distância menor
        driver.execute_script("window.scrollBy(0, 300);")  # Rolagem inicial mais curta (300 pixels)
        time.sleep(2)
        
        index = 6453  # Começar do índice desejado
        while index < len(df):
            try:
                row = df.iloc[index]
                print(f"Processando linha {index}")
                
                # Verificar se está pausado
                while pausado:
                    time.sleep(1)  # Esperar enquanto estiver pausado
                
                # Usar o seletor CSS atualizado do botão
                wait = WebDriverWait(driver, 10)
                adicionar_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, 
                    "button[data-type-range='zipcode'].btnAddZipcode.toggleRange.range-zipcode"
                )))
                adicionar_btn.click()
                time.sleep(2)
                
                # Verificar se está pausado
                while pausado:
                    time.sleep(1)
                
                driver.execute_script("window.scrollBy(0, 300);")  # Rolagem inicial mais curta (300 pixels)
                time.sleep(2)
                
                # Preencher os dados
                regiao = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Ex.: São Paulo Capital']")))
                regiao.clear()
                regiao.send_keys(str(row['Cidade']))
                
                # Verificar se está pausado
                while pausado:
                    time.sleep(1)
                
                # Garantir que os CEPs mantenham exatamente o formato da planilha
                cep_inicial = str(row['Cep Inicial']).strip()  # Remove espaços
                cep_final = str(row['Cep Final']).strip()  # Remove espaços
                
                # Preencher CEP inicial dígito por dígito
                campo_cep_inicial = driver.find_element(By.CSS_SELECTOR, "input[placeholder='_____-___']")
                for digito in cep_inicial:
                    # Verificar se está pausado
                    while pausado:
                        time.sleep(1)
                    campo_cep_inicial.send_keys(digito)
                    time.sleep(0.07)
                
                # Preencher CEP final dígito por dígito
                campo_cep_final = driver.find_elements(By.CSS_SELECTOR, "input[placeholder='_____-___']")[1]
                for digito in cep_final:
                    # Verificar se está pausado
                    while pausado:
                        time.sleep(1)
                    campo_cep_final.send_keys(digito)
                    time.sleep(0.07)
                
                # Verificar e ajustar o prazo de entrega
                if row['Prazo'] == 1:  # Se o prazo for 1 na planilha
                    # Clicar no seletor de prazo
                    prazo_select = wait.until(EC.element_to_be_clickable((
                        By.CSS_SELECTOR, 
                        "span.select2-selection__rendered[id='select2-id_prazo_entrega-container']"
                    )))
                    prazo_select.click()
                    time.sleep(1)
                    
                    # Selecionar 2 dias úteis
                    opcao_2dias = wait.until(EC.element_to_be_clickable((
                        By.XPATH, 
                        "//li[contains(text(), '2 dias úteis')]"
                    )))
                    opcao_2dias.click()
                    time.sleep(1)
                
                valor = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='0,00']")))
                valor.clear()
                valor.send_keys("15,90")
                
                try:
                    criar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnCreateRangeZipcode")))
                except:
                    try:
                        criar_btn = wait.until(EC.element_to_be_clickable((
                            By.XPATH, "//button[@type='submit' and @class='button--primary' and @id='btnCreateRangeZipcode']"
                        )))
                    except:
                        criar_btn = wait.until(EC.element_to_be_clickable((
                            By.XPATH, "//button[contains(text(), 'Criar faixa de CEP')]"
                        )))
                
                try:
                    criar_btn.click()
                except:
                    try:
                        driver.execute_script("arguments[0].click();", criar_btn)
                    except:
                        actions = webdriver.ActionChains(driver)
                        actions.move_to_element(criar_btn).click().perform()
                
                time.sleep(2)
                
                # Rolar até uma distância menor após criar faixa
                driver.execute_script("window.scrollBy(0, 200);")  # Rolagem mais curta após criar faixa (200 pixels)
                time.sleep(1)
                
                # Preencher os campos
                preencher_campo_inicial(cep_inicial)
                preencher_campo_final(cep_final)
                clicar_botao_salvar()
                
                # Aguardar um momento para o salvamento ser concluído
                time.sleep(2)
                
                # Atualizar a página usando o Selenium
                driver.refresh()
                
                # Aguardar o recarregamento da página
                time.sleep(3)
                
                index += 1  # Só avança para o próximo se não houver erro
                
            except TimeoutException:
                print(f"Timeout detectado na linha {index}. Tentando refresh...")
                driver.refresh()
                time.sleep(5)
                continue  # Mantém o mesmo índice
            except Exception as e:
                print(f"Erro ao processar linha {index}: {str(e)}")
                index += 1  # Avança para o próximo em caso de outros erros
                continue
                
    except Exception as e:
        print(f"Erro durante o login: {str(e)}")
    finally:
        driver.quit()

if __name__ == "__main__":
    preencher_faixas_cep()