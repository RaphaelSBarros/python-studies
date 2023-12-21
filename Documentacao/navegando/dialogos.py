from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
driver.get("")

alert = driver.switch_to.alert #retorna o alerta ativo

#comandos de voltar e avançar do navegador
driver.forward()
driver.back()