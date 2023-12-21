from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

driver = webdriver.Chrome()
driver.get("https://the-internet.herokuapp.com/download")

elementos = driver.find_elements(By.XPATH, "//div[@class='example']/a")
arquivos = []
i = 0
for elemento in elementos:
    if "2023" in elemento.text:
        i+=1
        print("Baixando "+elemento.text+"...")
        elemento.click()
        arquivos.append(elemento.text)
        time.sleep(1)
        
driver.get("https://demo.automationtesting.in/FileUpload.html#google_vignette")

path = "C:/Users/DESOUR10/Downloads/"

for arquivo in arquivos:
    driver.find_element(By.CSS_SELECTOR, "input[type='file']").send_keys(path + arquivo)
    time.sleep(1)
    
driver.find_element(By.XPATH, "/html[1]/body[1]/section[1]/div[1]/div[1]/div[1]/div[1]/button[3]/span[1]").click()