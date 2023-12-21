from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()
driver.get("")

#Vai pro site
driver.get("http://www.example.com")

# Seta os cookies pro site
cookie = {'name' : 'foo', 'value' : 'bar'}
driver.add_cookie(cookie)

# Pega todos os cookies disponíveis para a URL atual
driver.get_cookies()