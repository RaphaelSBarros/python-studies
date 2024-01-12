##Tentando automatizar o ronda com o Pyautogui

# Etapas do processo
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
import subprocess
import pyautogui
from dotenv import load_dotenv # Segurança das senhas

inicio = time.time()
load_dotenv() # Carrega o .env

### Acessar o link -> https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
options.add_argument(r'--profile-directory=Default')
driver = webdriver.Chrome(options=options)

driver.get("https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html")

### Informar o login 
try:
    element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
        EC.presence_of_element_located((By.XPATH, "//input[@id='login']"))
    )
except Exception as e:
    print(f"Erro ao tentar executar código: {e}")
else:
    loginSenior = driver.find_element(By.XPATH, "//input[@id='login']")
    loginSenior.send_keys(os.getenv('LOGIN'))
    
    ### Clicar em Log On
    logOnBtn = driver.find_element(By.ID, "loginBtn")
    logOnBtn.click()
    
    ### Digitar a senha 
    try:
        element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
            EC.presence_of_element_located((By.XPATH, "//input[@id='passwd']"))
        )
    except Exception as e:
        print(f"Erro ao tentar executar código: {e}")
    else:
        passwdSenior = driver.find_element(By.XPATH, "//input[@id='passwd']")
        passwdSenior.send_keys(os.getenv('SENHA')) # Puxa a senha armazenada no .env
        
        ### Clicar em Log On novamente
        logOnBtn = driver.find_element(By.ID, "loginBtn")
        logOnBtn.click()
        
        ### Ao logar Clicar em Acesso e Segurança - Produção
        try:
            element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
                EC.presence_of_element_located((By.CLASS_NAME, "storeapp-details-link"))
            )
        except Exception as e:
            print(f"Erro ao tentar executar código: {e}")
        else:
            acesso = driver.find_element(By.CLASS_NAME, "storeapp-details-link")
            acesso.click()
            
            ### Ao clicar, irá baixar um arquivo. Esse arquivo precisa ser clicado para abrir o sistema
            time.sleep(2) # Espera o arquivo ser baixado          
            lista_arquivos= glob.glob("C:/Users/DESOUR10/Downloads/*") #pegando todos os arquivos dentro da pasta downloads
            ultimo_arquivo=max(lista_arquivos, key=os.path.getmtime) #armazenando quem tem a data mais recente dentro da pasta
            driver.quit()
            try:
                # Executa o arquivo usando subprocess
                subprocess.run(ultimo_arquivo, shell=True)
            except Exception as e:
                print(f"Erro ao tentar executar código: {e}")
            else:
                time.sleep(60)
                pyautogui.PAUSE = 1
                pyautogui.click(640, 409) # Clicar na a área de login 
                pyautogui.write(os.getenv('USER'))
                pyautogui.click(640, 433) # Clicar na a área de senha 
                pyautogui.write(os.getenv('PSSWRD'))
                pyautogui.click(896, 336) # Clicar no Ok
                time.sleep(20) # Esperar para logar no sistema
                REs = ['106481','74740','70087','7975','60035','54257','54325','127626','31217'] ### Dados do usuário
                dataI = '01012024' ### Dados do usuário
                for count, RE in enumerate(REs):
                    if count == 0:
                        pyautogui.doubleClick(100, 100) # Clicar em Consultar Acessos
                        time.sleep(5) # Esperar o carregamento da consulta
                    pyautogui.click(405, 90) ## Clicar em Período (405, 90)
                    pyautogui.write(dataI) # Digitar a data de início 01022023
                    pyautogui.doubleClick(405, 115) ## Clicar em Local Físico (405, 115)
                    pyautogui.write('1') # Digitar 1 para Brazil
                    pyautogui.doubleClick(1310, 90) ## Clicar em Seleção (1310, 90)
                    time.sleep(5) # Esperar o carregamento da consulta
                    pyautogui.doubleClick(341, 46) # Em empresa inserir o código 100 (231, 46)
                    pyautogui.write('100')
                    pyautogui.doubleClick(232, 95) # Em Colaborador inserir o RE (232, 95)
                    pyautogui.write(RE)
                    pyautogui.doubleClick(1155, 50) # Clicar em Ok (1155, 50)
                    time.sleep(2) # Esperar o carregamento da consulta
                    pyautogui.doubleClick(1310, 120) ## Clicar em Mostrar (1310, 120)
                    time.sleep(3)
                    ## Salvando Dados
                    print(count, RE)
                    pyautogui.doubleClick(652, 272) # Baixar .xlsxW
                    pyautogui.rightClick() #botão direito na planilha
                    pyautogui.click(736, 309) #clica em exportar planilha
                    pyautogui.doubleClick(913, 368) #fecha a mensagem de erro
                    pyautogui.doubleClick(610, 385) #Seleciona o nome do arquivo
                    pyautogui.write(RE) # Escreve o nome do arquivo
                    if count == 0:
                        pyautogui.doubleClick(610, 385) #Seleciona o nome do arquivo
                        pyautogui.write(RE) # Escreve o nome do arquivo
                        pyautogui.doubleClick(600, 310) # Clica em Meu Computador
                        pyautogui.press('l') #Procura a o Disco Local
                        pyautogui.doubleClick(580, 250) 
                        pyautogui.press('u') #Procura a pasta Users
                        pyautogui.doubleClick(530, 315) #Clica na pasta
                        pyautogui.press('d') #Procura o usuário
                        pyautogui.doubleClick(540, 280) #Clica na pasta do usuário
                        time.sleep(2)
                        pyautogui.press('d') #Procura a pasta documentos
                        pyautogui.doubleClick(550, 280) #Clica na pasta Documentos
                    ## Caso não seja só salvar e confirmar
                    pyautogui.doubleClick(840, 385) # Clica em salvar
                    pyautogui.doubleClick(685, 420) # Clica em ok
finally:
    fim = time.time() 
    print(f'O código finalizou em {fim - inicio} segundos')