from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

email = "datamanagementbrazil@whirlpool.com"
senha = "Rh2020"

driver = webdriver.Chrome()
driver.get("https://apps.docusign.com/send/documents?type=envelopes")

wait = WebDriverWait(driver, 10)

#Teste de login e senha
try:
    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter email']"))
    )
finally:
    driver.find_element(By.NAME, "email").send_keys(email, Keys.RETURN)
    try:
        wait.until(
            EC.presence_of_element_located((By.NAME, "password"))
        )
    finally:
        driver.find_element(By.NAME, "password").send_keys(senha, Keys.RETURN)
        try:
            wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter code']"))
            )
        finally:
            code = input("Insira o código: ")
            if code != "":
                driver.find_element(By.CSS_SELECTOR, "input[placeholder='Enter code']").send_keys(code, Keys.RETURN)
            else:
                print("Código não informado. Encerrando o processo")
                quit()

    
arquivo = 'C:/Users/DESOUR10/Downloads/EnviarAvisosDeFérias - Layout.xlsm'
