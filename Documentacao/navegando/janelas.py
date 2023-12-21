from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
driver.get("")

#Encontrando o nome da página
#<a href="somewhere.html" target="windowName">Click here to open a new window</a>
driver.switch_to.window("windowName") #usando o nome para interagir

#Dá pra navegar por todas as páginas abertas
for handle in driver.window_handles:
    driver.switch_to.window(handle)

#Navegando por frames e Iframes
driver.switch_to.frame("frameName") #ou
driver.switch_to.frame("frameName.0.child") #acessando subframes
driver.switch_to.default_content() #voltando para o frame principal
