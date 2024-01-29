from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import workbook, load_workbook
import subprocess
import datetime
import time
from tkinter import filedialog

file_path = filedialog.askopenfilename() # Escolhe o diretório do arquivo
planilha = load_workbook(filename=file_path, data_only= True) # Data_only evita o armazenamento das fórmulas
aba_ativa = planilha.active

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\Selenium Data")
options.add_argument(r'--profile-directory=Default')

#driver = webdriver.Chrome(options=options)

#driver.get("https://performancemanager8.successfactors.com/sf/customExternalModule?bplte_company=whirlpoolc&_s.crb=EsOgEIHB37k6ucolE3YkiE4vImirlYAXlhbuqs2HjaY%253d")

# PERCORRE TODAS AS LINHA DA COLUNA C E CASO O VALOR FOR IGUAL A RC002-01 TROCA O VALOR DA LINHA CORRESPONDENTE NA COLUNA E PARA 'NOK'
for celula in aba_ativa['C'][2:]:
    if celula.value is not None:
        linha = celula.row       
        id = aba_ativa[f'A{linha}'].value 
        ## Formatando a data
        dataVerificar = aba_ativa[f'B{3}'].value
        dataVerificar = str(dataVerificar)
        dataVerificar = f'{dataVerificar[8:10]}/{dataVerificar[5:7]}/{dataVerificar[:4]}'
        dataAlvo = (aba_ativa[f'D{linha}'].value)
        
        ## SELENIUM ##
        
        
planilha.save("TesteOpenpy.xlsx")
#subprocess.run('TesteOpenpy.xlsx', shell=True)