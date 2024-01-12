from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

file_path = 'C:/Users/DESOUR10/Downloads/Macro Workforce.xlsm' # Escolhe o diretório do arquivo
df = pd.read_excel(
    io=file_path,
    sheet_name="Formulário"
) # Lê o arquivo

print(df) # Exibe a tabela