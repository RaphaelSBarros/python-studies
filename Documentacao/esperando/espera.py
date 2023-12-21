from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC #importa todas as condições esperadas

driver = webdriver.Chrome()
driver.get("http://somedomain/url_that_delays_loading")

##Espera Explícita

try:
    element = WebDriverWait(driver, 10).until( #espera até 10s para que o elemento especificado apareça
        EC.presence_of_element_located((By.ID, "myDynamicElement"))
    )
finally:
    driver.quit()
    

##Espera Implícita

driver.implicitly_wait(10)
driver.get("http://somedomain/url_that_delays_loading")
myDynamicelement = driver.find_element(By.ID("myDynamicElement"))