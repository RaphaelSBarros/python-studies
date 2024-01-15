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

planilha = load_workbook(file_path)
aba_ativa = planilha.active

# PERCORRE TODAS AS LINHA DA COLUNA C E CASO O VALOR FOR IGUAL A RC002-01 TROCA O VALOR DA LINHA CORRESPONDENTE NA COLUNA E PARA 'NOK'
for celula in aba_ativa['C'][2:]:
    if celula.value != None:
        linha = celula.row
        aba_ativa[f'E{linha}'] = datetime.time(12,00,00) # Resolvendo problema na hora das fórmulas lerem os valores.
        aba_ativa[f'F{linha}'] = datetime.time(18,15,00)

planilha.save("TesteOpenpy.xlsx")
subprocess.run('TesteOpenpy.xlsx', shell=True)