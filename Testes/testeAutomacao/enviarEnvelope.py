from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
from dotenv import load_dotenv # Segurança das senhas   

inicio = time.time()
load_dotenv() # Carrega o .env

navegar = True

organizar = False

#///////////////////////////// UTILIZAR O MODELO CRIADO PARA PREENCHER EMAILS /////////////////////////////

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
options.add_argument(r'--profile-directory=Default')

if navegar:
    driver = webdriver.Chrome(options=options)

    driver.get("https://apps.docusign.com/send/documents")

    try:
        element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
            EC.presence_of_element_located((By.XPATH, "//button[@type='submit']"))
        )
        if element.text  == "NEXT":
            element.click()
            try:
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//button[@type='submit']"))
                )
                element.click()
            except:
                print("erro no botão de login")
        else:
            element.click()
    except:
        print("não foi possível efetuar o login")
    else:
        try:
            time.sleep(5)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='header-TEMPLATES-tab-button']"))
            )
            element.click()
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@data-qa='templates-main-header-form-input']"))
            )
            element.send_keys("Modelo de Contrato - CC")
            element.send_keys(Keys.ENTER)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='templates-main-list-row-b4024bba-b898-4c4d-a1c5-3ef9b408ee3c-actions-use']"))
            )
            element.click()
        except:
            print("erro no butau")
            c = input("a: ")
        else:
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "windows-drag-handler-wrapper"))
                )
                element.find_element(By.TAG_NAME, "input").send_keys(r"C:\Users\DESOUR10\Downloads\CONTRATOS RIO CLARO ADM 01.02.pdf")
            except:
                print("envou nau")
            else:
                try:
                    element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@data-qa='prepare-subject']"))
                    )
                    element.send_keys(Keys.CONTROL+"a")
                    element.send_keys(Keys.DELETE)
                    element.send_keys("Contrato teste")
                    element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//textarea[@data-qa='prepare-message']"))
                    )
                    element.send_keys(Keys.CONTROL+"a")
                    element.send_keys(Keys.DELETE)
                    element.send_keys("mensagem teste")
                except:
                    print("não aplicou o modelo")
                else:
                    try:
                        element = WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='footer-add-fields-link']"))
                        )
                        time.sleep(2)
                        element.click()
                    except:
                        print("Achou n")
                    else:
                        try:
                            element = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-button']"))
                            )
                            element.click()
                            element = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-level-75']"))
                            )
                            element.click()
                        except:
                            print("Achou o zoom n")
                        else:
                            tam = driver.find_element(By.XPATH, "//span[@data-qa='uploaded-file-pages']").text
                            tam = int(tam[-2:])
                            rubrica = driver.find_element(By.XPATH, "//button[@data-qa='Initial']")
                            assinatura = driver.find_element(By.XPATH, "//button[@data-qa='Signature']")
                            try:
                                element = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, "//div[@data-qa='document-accordion-region']"))
                                )
                                paginas = element.find_elements(By.XPATH, "//button[@data-qa='tagger-documents']")
                                for i, pagina in enumerate(paginas):
                                    rubrica.click()
                                    try:
                                        element = WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='tagger-documents']"))
                                        )
                                    except:
                                        print("não achou as páginas")
                                    else:
                                        pagina.click()
                                        if i % 2 == 0 or i == 0:
                                            ActionChains(driver)\
                                                .click_and_hold(rubrica)\
                                                .move_by_offset(450, 370)\
                                                .release()\
                                                .perform()
                                        else:
                                            ActionChains(driver)\
                                                .click_and_hold(assinatura)\
                                                .move_by_offset(335, 325)\
                                                .release()\
                                                .perform()
                                            C = input("A: ")
                            except:
                                print("deu ruim o final")
                            finally:
                                C = input("A: ")

if organizar:
    folder_path = "C://Users//DESOUR10//Downloads//" #escolhe a pasta dos arquivos
    arquivos = os.listdir(folder_path) #lista todos os arquivos dentro da pasta escolhida
    nome = 'Aviso_Férias'
    try: 
        os.mkdir(folder_path+nome) # Cria uma pasta no destino
    except:
        print("Pasta já existe")
    finally:
        for arquivo in arquivos: #para cada arquivo dentro da lista de arquivos
            if 'pdf' in arquivo: #caso tenha pdf no arquivo
                if nome in arquivo: # e caso tenha Compras no nome
                    source_path = os.path.join(folder_path, arquivo) #pega o arquivo que se encaixa
                    if not '(1)' in arquivo: 
                        destination_path = os.path.join(folder_path, nome, arquivo) #cria uma variável com o destino do arquivo
                        os.rename(source_path, destination_path) #coloca o arquivo dentro da pasta de destino
                    else: ## Caso o arquivo esteja duplicado
                        os.remove(source_path) ## O arquivo é excluído