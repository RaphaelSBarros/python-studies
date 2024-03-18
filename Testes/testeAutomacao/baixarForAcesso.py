from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

import datetime
import time
import os
import pandas as pd
from dotenv import load_dotenv
load_dotenv()

class ForAcesso:
    def __init__(self, email, senha):
        self.SITE_LINK = {
            "home": "http://awn-wappp17/ForacessoWeb/foracesso.aspx",
        }
        self.SITE_MAP = {
            "buttons": {
                "popup_btn": {
                    "xpath": "//input[@id='popAlerta_btnConfirmaPopup']"
                },
                "login_btn": {
                    "xpath": "//input[@id='linkBtnOK']"
                },
                "consultas_btn": {
                    "xpath": "//li[@title='Consultas']"
                },
                "executar_consultas_btn": {
                    "xpath": "//li[@id='WucMenuPrincipal1_LI_06_05']"
                }
            },
            "inputs": {
                "username": {
                    "xpath": "//input[@id='txtUsuario']",
                    "keys": email
                },
                "password": {
                    "xpath": "//input[@id='txtSenha']",
                    "keys": senha
                },
                "query": {
                    "checkbox": {
                        "xpath": "//span[@title='Informe se deve gravar o resultado da consulta em TXT']"
                    },
                    "code":{
                        "xpath": "//input[@id='ctl00_cphCadastro_txtCodigo_I']",
                        "keys": "001"
                    },
                    "name":{
                        "xpath": "//input[@id='ctl00_cphCadastro_txtNomeTxt_I']",
                        "keys": "ForAcesso"
                    }
                }
            }
        }
        
        self.options = webdriver.ChromeOptions()
        self.options.add_argument(r"--user-data-dir=C:\Users\DESOUR10\AppData\Local\Google\Chrome\Test Data")
        self.options.add_argument(r'--profile-directory=Default')
        
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.maximize_window()
        
        self.wait = WebDriverWait(self.driver, 30)
    
    def open_website(self):
        self.driver.get(self.SITE_LINK['home'])
        
    def login(self):
        try:
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['popup_btn']['xpath']))
            ).click()
        except:
            pass
        finally:
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['username']['xpath']))
            ).send_keys(self.SITE_MAP['inputs']['username']['keys'])
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['password']['xpath']))
            ).send_keys(self.SITE_MAP['inputs']['password']['keys'])
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['login_btn']['xpath']))
            ).click()

    def click_consultas(self):
        self.driver.switch_to.default_content()
    
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['buttons']['consultas_btn']['xpath']))
        ).click()

    def execute_query(self):
        element = self.wait.until(
            EC.presence_of_element_located((By.XPATH, self.SITE_MAP['buttons']['executar_consultas_btn']['xpath']))
        )
        element.find_element(By.TAG_NAME, "a").click()
        
        self.driver.switch_to.frame('ifrmPrincipal')
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['query']['checkbox']['xpath']))
        ).click()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['query']['name']['xpath']))
        ).send_keys(self.SITE_MAP['inputs']['query']['name']['keys'])
        
        element = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, self.SITE_MAP['inputs']['query']['code']['xpath']))
        )
        element.send_keys(self.SITE_MAP['inputs']['query']['code']['keys'])
        element.send_keys(Keys.ENTER)
        a = datetime.datetime.now()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='ctl00_popAlerta_btnConfirmaPopup']"))
        ).click()
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='ctl00_btnCancelar']"))
        ).click()
        return a

    def download_query(self):
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//li[@id='WucMenuPrincipal1_LI_06_15']"))
        ).click()
        
        self.driver.switch_to.frame('ifrmPrincipal')
        
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='ctl00_cphConsulta_grdConsulta_cell0_0_btnDownload']"))
        ).click()

# processo = ForAcesso(os.getenv('foracessoU'), os.getenv('foracessoS'))
# processo.open_website()
# processo.login()
# processo.click_consultas()
# a = processo.execute_query()
# processo.click_consultas()
# processo.download_query()
# print(a)

from_path = r"C:\Users\DESOUR10\Downloads"
to_path = r"C:\Users\DESOUR10\Downloads\Conferência"
folder = os.listdir(from_path)

for file in folder:
    if 'ForAcesso' in file:
        file_path = os.path.join(from_path, file)
        df = pd.read_csv(file_path, sep=';')
        df.to_excel((os.path.join(to_path, 'ForAcesso.xlsx')), index=False)
        break
print('foi')
