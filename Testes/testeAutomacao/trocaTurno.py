from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import workbook, load_workbook
import subprocess
import datetime
import time

file_path = 'C:/Users/DESOUR10/Downloads/Macro Workforce.xlsm' # Escolhe o diretório do arquivo

planilha = load_workbook(file_path)
aba_ativa = planilha.active

# PERCORRE TODAS AS LINHA DA COLUNA C E CASO O VALOR FOR IGUAL A RC002-01 TROCA O VALOR DA LINHA CORRESPONDENTE NA COLUNA E PARA 'NOK'
for celula in aba_ativa['G']:
    linha = celula.row
    aba_ativa[f'E{linha}'] = datetime.time(12,00,00) # Resolvendo problema na hora das fórmulas lerem os valores.
for celula in aba_ativa['H']:
    linha = celula.row
    aba_ativa[f'F{linha}'] = datetime.time(18,15,00) # Caso seja inserido as horas em formato de texto as fórmulas não conseguem ler
    

planilha.save("TesteOpenpy.xlsx")
subprocess.run('TesteOpenpy.xlsx', shell=True)