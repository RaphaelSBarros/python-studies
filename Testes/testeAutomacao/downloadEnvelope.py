from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
from dotenv import load_dotenv # Segurança das senhas

inicio = time.time()
load_dotenv() # Carrega o .env

continuar = False

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\Test Data")
options.add_argument(r'--profile-directory=Default')
driver = webdriver.Chrome(options=options)

driver.get("https://apps.docusign.com/send/documents")

try:
    element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
        EC.presence_of_element_located((By.XPATH, "//button[@type='submit']"))
    )
except:
    print("não achou o botão")
else:
    entrarBtn = driver.find_element(By.XPATH, "//button[@type='submit']")
    entrarBtn.click()
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@title='Concluído']"))
        )
    except:
        print("Não achou o concluído")
    else:
        concluidoBtn = driver.find_element(By.XPATH, "//button[@title='Concluído']") 
        concluidoBtn.click()
        time.sleep(2)
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//table"))
            )
        except:
            print("não achou a lista")
        else:
            tabela = driver.find_element(By.TAG_NAME, "tbody")
            linhas = tabela.find_elements(By.TAG_NAME, "tr")
            for linha in linhas:
                print(linha.text)
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "td"))
                    )
                except:
                    print("não achou a tabela")
                else:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    ultimaCol = len(colunas) - 1
                    if 'CONTRATO DE TRABALHO - CARLOS EDUARDO DAMACENO - 10231590' in linha.text:
                        colunas[ultimaCol].find_element(By.TAG_NAME, "button").click()
                        try:
                            element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//label/span/span[text()='Combinar todos os PDFs em um único arquivo']"))
                            )
                        except:
                            print("não achou a lista")
                        else:
                            combinarPDF = driver.find_element(By.XPATH, "//label/span/span[text()='Combinar todos os PDFs em um único arquivo']")
                            combinarPDF.click()
                            try:
                                element = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, "//div[3]/button"))
                                )
                                element.click()
                            except:
                                print("inclicável")
                                driver.quit()
                            time.sleep(3)

if continuar:
    folder_path = "C://Users//DESOUR10//Downloads//" #escolhe a pasta dos arquivos
    arquivos = os.listdir(folder_path) #lista todos os arquivos dentro da pasta escolhida

    for arquivo in arquivos: #para cada arquivo dentro da lista de arquivos
        if 'pdf' in arquivo: #caso tenha pdf no arquivo
            if 'Aviso_Férias' in arquivo: # e caso tenha Compras no nome
                source_path = os.path.join(folder_path, arquivo) #pega o arquivo que se encaixa
                destination_path = os.path.join(folder_path, "férias", arquivo) #cria uma variável com o destino do arquivo C://Users//DESOUR10//Downloads//07-16 - Copiar e Mover Arquivos//Compras//
                os.rename(source_path, destination_path) #coloca o arquivo dentro da pasta de destino C://Users//DESOUR10//Downloads//07-16 - Copiar e Mover Arquivos//Compras//arquivo