import unittest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

class PythonOrgSearch(unittest.TestCase): #Classe de testes herdando de unittest.TestCase que herda da classe TestCase
    def setUp(self): #parte da inicialização dos testes
        self.driver = webdriver.Chrome() #nesse caso criando a instancia do navegador
    
    def test_SearchInPythonOrg(self): #O teste em si. Deve sempre começar com *test*
        driver = self.driver #Cria uma referencia local para self.driver no metodo setUp
        driver.get("https://www.python.org") #navegação
        self.assertIn("Python", driver.title) #Confirma se o título da página tem Python escrito
        elem = driver.find_element(By.NAME, "q") #Localiza o elemento Name com valor q
        elem.send_keys("pycon") #escreve no elemento name='q'
        elem.send_keys(Keys.RETURN) #enter
        self.assertNotIn("No results found.", driver.page_source) #resultados encontrados
    
    def tearDown(self): #método para limpar todas as ações anteriores
        self.driver.close()
        
if __name__ == "__main__": #Não faço ideia. Deixa aí
    unittest.main()