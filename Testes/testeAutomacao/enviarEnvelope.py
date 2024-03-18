from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import time

import os
from dotenv import load_dotenv
load_dotenv()

navegar = True

j=0
iniciog = time.time()
    
#///////////////////////////// UTILIZAR O MODELO CRIADO PARA PREENCHER EMAILS /////////////////////////////

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
options.add_argument(r'--profile-directory=Default')

while navegar:
    iniciop = time.time()
    driver = webdriver.Chrome(options=options)
    driver.get("https://apps.docusign.com/send/documents")
    j+=1
    
    try:
        element = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
        )
        if element.get_dom_attribute("data-qa")  == "submit-username":
            try:
                element = WebDriverWait(driver, 55).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@data-qa='username']"))
                )
                element.send_keys(Keys.CONTROL+"a")
                element.send_keys(Keys.DELETE)
                element.send_keys(os.getenv('LOGIN'))
                element = WebDriverWait(driver, 55).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='submit-username']"))
                )
                element.click()
            except:
                print("erro no botão de NEXT")
                navegar=False
        else:
            element = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, "//input[@data-qa='password']"))
            )
            element.send_keys(Keys.CONTROL+"a")
            element.send_keys(Keys.DELETE)
            element.send_keys(os.getenv('SENHA'))
            element = WebDriverWait(driver, 50).until( #espera até 10s para que o elemento especificado apareça
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='submit-password']"))
            )
            element.click()
    except:
        print("Erro no botão de Log In")
        navegar=False
    else:
        try:
            element = WebDriverWait(driver, 50).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='header-TEMPLATES-tab-button']"))
            )
            time.sleep(4)
            element.click()
            element = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.XPATH, "//input[@data-qa='templates-main-header-form-input']"))
            )
            element.send_keys("Modelo de Contrato - CC")
            element.send_keys(Keys.ENTER)
            element = WebDriverWait(driver, 50).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='templates-main-list-row-b4024bba-b898-4c4d-a1c5-3ef9b408ee3c-actions-use']"))
            )
            element.click()
        except:
            print("Erro na seleção de Modelo")
            navegar=False
        else:
            try:
                element = WebDriverWait(driver, 50).until(
                    EC.presence_of_element_located((By.ID, "windows-drag-handler-wrapper"))
                )
                element.find_element(By.TAG_NAME, "input").send_keys(r"C:\Users\DESOUR10\Downloads\CONTRATOS RIO CLARO ADM 01.02.pdf")
            except:
                print("Erro no envio do arquivo para assinatura")
                navegar=False
            else:
                try:
                    element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@data-qa='prepare-subject']"))
                    )
                    element.send_keys(Keys.CONTROL+"a")
                    element.send_keys(Keys.DELETE)
                    element.send_keys("Contrato teste")
                    element = WebDriverWait(driver, 55).until(
                        EC.element_to_be_clickable((By.XPATH, "//textarea[@data-qa='prepare-message']"))
                    )
                    element.send_keys(Keys.CONTROL+"a")
                    element.send_keys(Keys.DELETE)
                    element.send_keys("mensagem teste")
                except:
                    print("Não foi possível modificar o título e a mensagem do envio")
                    navegar=False
                else:
                    try:
                        element = WebDriverWait(driver, 55).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='footer-add-fields-link']"))
                        )
                        element.click()
                    except:
                        print("Não foi possível encontrar o botão de Avançar para assinatura")
                        navegar=False
                    else:
                        try:
                            element = WebDriverWait(driver, 55).until(
                                EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-button']"))
                            )
                            element.click()
                            element = WebDriverWait(driver, 55).until(
                                EC.presence_of_element_located((By.XPATH, "//button[@data-qa='zoom-level-75']"))
                            )
                            element.click()
                        except:
                            print("Achou o zoom n")
                            navegar=False
                        else:
                            element = WebDriverWait(driver, 50).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@data-qa='document-accordion-region']"))
                            )
                            rubrica = driver.find_element(By.XPATH, "//button[@data-qa='Initial']")
                            assinatura = driver.find_element(By.XPATH, "//button[@data-qa='Signature']")
                            paginas = element.find_elements(By.XPATH, "//button[@data-qa='tagger-documents']")
                            try:
                                for i, pagina in enumerate(paginas):
                                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                                    if "tem campos nesta página" not in alt_value:
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
                                                .move_by_offset(332, 325)\
                                                .release()\
                                                .perform()
                                        rubrica.click()
                            except ElementClickInterceptedException:
                                element = WebDriverWait(driver, 55).until(
                                    EC.presence_of_element_located((By.XPATH, "//button[@data-qa='modal-cancel-btn']"))
                                ).click()
                            finally:
                                element = WebDriverWait(driver, 50).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='tagger-documents']"))
                                )
                                for i, pagina in enumerate(paginas):
                                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                                    if "tem campos nesta página" not in alt_value:
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
                                                .move_by_offset(332, 325)\
                                                .release()\
                                                .perform()
                                        rubrica.click()
                                for i, pagina in enumerate(paginas):
                                    alt_value = pagina.find_element(By.TAG_NAME, "img").get_dom_attribute("alt")
                                    print(f"{i}: {alt_value}")
    finally:
        driver.quit()
        print(f"Tentativa: {j}")
        fim = time.time()
        print(f'O código finalizou em {round(fim - iniciop,2)} segundos')
        
fim = time.time()
print(f"O Programa quebrou depois de {j} tentativas\nTempo médio de execução de {round(fim - iniciog,2)/(j)} segundos")