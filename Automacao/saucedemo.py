from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome()
driver.get("https://www.saucedemo.com/")

wait = WebDriverWait(driver, 10)
login = []

try:
    wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "login_credentials_wrap-inner"))
    )
finally:
    elementos = driver.find_element(By.XPATH, "//div[@class='login_credentials']").text.split()
    i=0
    print(elementos)

