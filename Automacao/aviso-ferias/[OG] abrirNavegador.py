### [ INICIO ]Abrir o Chrome em perfil específico

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

#Create chromeoptions instance
options = webdriver.ChromeOptions()
#Provide location where chrome stores profiles
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\User Data")
#Provide the profile name with which we want to open browser
options.add_argument(r'--profile-directory=Default')
#Specify where your chrome driver present in your pc
driver = webdriver.Chrome(options=options)
#Provide website url here
driver.get("https://apps.docusign.com/send/documents")
#Aguarda a página carregar
time.sleep(10)

### [ FIM ] Abrir o Chrome em perfil específico ###


## [ INICIO ] Preenche o campo email ##

#Variável contendo o email
email = "datamanagementbrazil@whirlpool.com"

#Lista os elementos que propuserem entrada de email
elementos_email = driver.find_elements("xpath","(//input[@placeholder='Enter email'])[1]")
#Conta quantos elementos foram encontrados em elementos_email
quantidade_elementos_email = len(elementos_email)

# SE a quantidade de elementos_email for MAIOR que 0
if quantidade_elementos_email > 0:
    # ENTÃO
    # Envia a variável email para o campo de texto
    driver.find_element("xpath","(//input[@placeholder='Enter email'])[1]").send_keys(email)
    # Envia um comando de clique para a caixa "NEXT"
    driver.find_element("xpath","(//button[@type='submit'])[1]").click()
else:
    print("Aviso: O elemento de email não foi encontrado ou requirido")

## [ FIM ] Preenche o campo email ##


## [ INICIO ] Preenche o campo senha ##
    
#Variável contendo o senha
senha = "Rh2020"

#Lista os elementos que propuserem entrada de senha
elementos_senha = driver.find_elements("xpath","(//input[@placeholder='Enter password'])[1]")
#Conta quantos elementos foram encontrados em elementos_senha
quantidade_elementos_senha = len(elementos_senha)

# SE a quantidade de elementos_senha for MAIOR que 0
if quantidade_elementos_senha > 0:
    # ENTÃO
    # Envia a variável senha para o campo de texto
    driver.find_element("xpath","(//input[@placeholder='Enter password'])[1]").send_keys(senha)
    # Envia um comando de clique para a caixa "NEXT"
    driver.find_element("xpath","(//button[@type='submit'])[1]").click()
else:
    print("Aviso: O elemento de senha não foi encontrado ou requirido")

## [ FIM ] Preenche o campo senha ##
    

## [ INICIO ] Preenche o campo código de verificação ##
    
#Variável contendo o código de autenticação (auth_code)
auth_code = ""

#Lista os elementos que propuserem entrada de código de autenticação
elementos_auth = driver.find_elements("xpath","(//input[@placeholder='Enter password'])[1]")
#Conta quantos elementos foram encontrados em elementos_auth
quantidade_elementos_auth = len(elementos_auth)

# SE a quantidade de elementos_auth for MAIOR que 0
if quantidade_elementos_auth > 0:
    # ENTÃO
    # Envia a variável auth para o campo de texto
    driver.find_element("xpath","(//input[@placeholder='Enter code'])[1]").send_keys(auth_code)
    # Envia um comando de clique para a caixa "NEXT"
    driver.find_element("xpath","(//button[@type='submit'])[1]").click()
else:
    print("Aviso: O elemento de auth não foi encontrado ou requirido")

## [ FIM ] Preenche o campo auth ##
    

