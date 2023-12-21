# Testando o selenium no python

from selenium import webdriver #provém todas as implementações do WebDriver
from selenium.webdriver.common.keys import Keys #teclas do teclado
from selenium.webdriver.common.by import By #usada para localizar elementos dentro de um documento

driver = webdriver.Chrome() #Cria a instância do navegador
driver.get("https://www.python.org") #Navega até a página informada passada por URL
assert "Python" in driver.title #Confirma se a palavra Python está dentro do título da página

elem = driver.find_element(By.NAME, "q") #Localiza algum elemento NAME com o valor 'q'
elem.clear() #Limpa o campo
elem.send_keys("pycon") #Digita no campo
elem.send_keys(Keys.RETURN) #Aperta enter
assert "No results found." not in driver.page_source
driver.close()