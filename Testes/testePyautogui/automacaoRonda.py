##Tentando automatizar o ronda com o Pyautogui

# Etapas do processo
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
import subprocess
import pyautogui
from dotenv import load_dotenv # Segurança das senhas

load_dotenv() # Carrega o .env

### Acessar o link -> https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
options.add_argument(r'--profile-directory=Default')
driver = webdriver.Chrome(options=options)
erro = False
while not erro:
    driver.get("https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html")

### Informar o login 

    try:
        element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
            EC.presence_of_element_located((By.XPATH, "//input[@id='login']"))
        )
    except:
        print("Deu no login 1")
        erro=True
        raise
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
        except:
            print("Deu na senha 1")
            erro=True
            raise
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
            except:
                print("erro no acesso")
                erro=True
                raise
            else:
                acesso = driver.find_element(By.CLASS_NAME, "storeapp-details-link")
                acesso.click()
                
                ### Ao clicar, irá baixar um arquivo. Esse arquivo precisa ser clicado para abrir o sistema
                time.sleep(2) # Espera o arquivo ser baixado          
                lista_arquivos= glob.glob("C:/Users/DESOUR10/Downloads/*") #pegando todos os arquivos dentro da pasta downloads
                ultimo_arquivo=max(lista_arquivos, key=os.path.getmtime) #armazenando quem tem a data mais recente dentro da pasta
                try:
                    # Executa o arquivo usando subprocess
                    subprocess.run(ultimo_arquivo, shell=True)
                except Exception as e:
                    print(f"Erro ao tentar executar o arquivo: {e}")
                    erro=True
                else:
                    time.sleep(60)
                    pyautogui.PAUSE = 1
                    pyautogui.moveTo(640, 409) # Mover para a área de login 
                    pyautogui.click()
                    pyautogui.write(os.getenv('USER'))
                    pyautogui.moveTo(640, 433) # Mover para a área de senha 
                    pyautogui.click()
                    pyautogui.write(os.getenv('PWD'))
                    pyautogui.moveTo(896, 336) # Mover para o Ok
                    pyautogui.click()
                    time.sleep(20) # Esperar para logar no sistema
                    pyautogui.moveTo(100, 100) # Clicar em Consultar Acessos (100, 100)
                    pyautogui.doubleClick()
                    time.sleep(10) # Esperar o carregamento da consulta
                    pyautogui.moveTo(405, 90) ## Clicar em Período (405, 90)
                    pyautogui.click()
                    pyautogui.write('15122023') # Digitar a data de início 01022023
                    pyautogui.moveTo(405, 115) ## Clicar em Local Físico (405, 115)
                    pyautogui.doubleClick()
                    pyautogui.write('1') # Digitar 1 para Brazil
                    pyautogui.press('enter') # Apertar Enter para exibir
                    pyautogui.moveTo(1310, 90) ## Clicar em Seleção (1310, 90)
                    pyautogui.doubleClick()
                    time.sleep(5) # Esperar o carregamento da consulta
                    pyautogui.moveTo(341, 46) # Em empresa inserir o código 100 (232, 45)
                    pyautogui.doubleClick()
                    pyautogui.write('100')
                    pyautogui.moveTo(232, 95) # Em Colaborador inserir o RE (232, 95)
                    pyautogui.doubleClick()
                    pyautogui.write('929358')
                    pyautogui.moveTo(1155, 50) # Clicar em Ok (1155, 50)
                    pyautogui.doubleClick()
                    time.sleep(2) # Esperar o carregamento da consulta
                    pyautogui.moveTo(1310, 120) ## Clicar em Mostrar (1310, 120)
                    pyautogui.doubleClick()
                    
                    pyautogui.hotkey('alt', 'f4')