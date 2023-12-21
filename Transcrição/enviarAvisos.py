from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os

chrome_options = Options()
chrome_options.add_argument("user-data-dir=C:/Users/DESOUR10/AppData/Local/Google/Chrome/User Data")

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
                print("Código não informado. Encerrando o processo")
                quit()
            driver.refresh()