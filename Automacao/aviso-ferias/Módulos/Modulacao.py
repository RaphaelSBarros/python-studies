from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time

class AutenticarAcesso():
    def __init__(self):
        # Create chromeoptions instance
        self.options = webdriver.ChromeOptions()
        # Provide location where chrome stores profiles
        self.options.add_argument(r"--user-data-dir=C:\Users\CAMBIA3\AppData\Local\Google\Chrome\User Data - Selenium")
        # Provide the profile name with which we want to open the browser
        self.options.add_argument(r'--profile-directory=Profile 7')
        # Specify where your chrome driver is present on your PC
        self.driver = webdriver.Chrome(options=self.options)
    def abrir_chrome_com_perfil(self):
        ### [ INICIO ] Abrir o Chrome em perfil específico ###
        # Provide website URL here
        self.driver.get("https://apps.docusign.com/send/documents")
        # Maximiza a tela para exibir todos os elementos
        self.driver.maximize_window()
        # Aguarda a página carregar
        time.sleep(5)

        ### [ FIM ] Abrir o Chrome em perfil específico ###
        return self.driver

    def verificar_autenticação_email(self):
        ## [ INICIO ] Preenche o campo email ##

        #Variável contendo o email
        email = "datamanagementbrazil@whirlpool.com"

        #Lista os elementos que propuserem entrada de email
        elementos_email = self.driver.find_elements("xpath","(//input[@placeholder='Enter email'])[1]")
        #Conta quantos elementos foram encontrados em elementos_email
        quantidade_elementos_email = len(elementos_email)

        # SE a quantidade de elementos_email for MAIOR que 0
        if quantidade_elementos_email > 0:
            # ENTÃO
            # Envia a variável email para o campo de texto
            self.driver.find_element("xpath","(//input[@placeholder='Enter email'])[1]").send_keys(email)
            # Envia um comando de clique para a caixa "NEXT"
            self.driver.find_element("xpath","(//button[@type='submit'])[1]").click()
        else:
            print("Aviso: O elemento de email não foi encontrado ou requirido")
        
        ## [ FIM ] Preenche o campo email ##

    def verificar_autenticação_senha(self, s):
        ## [ INICIO ] Preenche o campo senha ##

        # Obtem a senha da instância de ChamarUI
        self.senha = s

        #Lista os elementos que propuserem entrada de senha
        elementos_senha = self.driver.find_elements("xpath","(//input[@placeholder='Enter password'])[1]")
        #Conta quantos elementos foram encontrados em elementos_senha
        quantidade_elementos_senha = len(elementos_senha)

        # SE a quantidade de elementos_senha for MAIOR que 0
        if quantidade_elementos_senha > 0:
            # ENTÃO
            # Envia a variável senha para o campo de texto
            self.driver.find_element("xpath","(//input[@placeholder='Enter password'])[1]").send_keys(self.senha)
            # Envia um comando de clique para a caixa "NEXT"
            self.driver.find_element("xpath","(//button[@type='submit'])[1]").click()
        else:
            print("Aviso: O elemento de senha não foi encontrado ou requirido")

        ## [ FIM ] Preenche o campo senha ##
            
    def contar_elementos_auth(self):   
        ## [ INICIO ] Preenche o campo código de verificação ##

        #Lista os elementos que propuserem entrada de código de autenticação
        self.elementos_auth = self.driver.find_elements("xpath","(//input[@placeholder='Enter code'])[1]")
        #Conta quantos elementos foram encontrados em elementos_auth
        self.quantidade_elementos_auth = len(self.elementos_auth)
            
        ## [ FIM ] Preenche o campo código de verificação ##

    def verificar_autenticação_auth(self, c):   
        ## [ INICIO ] Preenche o campo código de verificação ##

        #Define a variável auth_code
        self.auth_code = c

        # SE a quantidade de elementos_auth for MAIOR que 0
        if self.quantidade_elementos_auth > 0:
            # ENTÃO
            # Envia a variável auth para o campo de texto
            self.driver.find_element("xpath","(//input[@placeholder='Enter code'])[1]").send_keys(self.auth_code)
            # Envia um comando de clique para a caixa "NEXT"
            self.driver.find_element("xpath","(//button[@type='submit'])[1]").click()
        else:
            print("Aviso: O elemento de auth não foi encontrado ou requirido")
            return self.quantidade_elementos_auth

        ## [ FIM ] Preenche o campo auth ##

    def abrir_envelope(self):

        self.driver.find_element("xpath","(//button[@class='olv-button olv-ignore-transform css-17ozirp'])[1]").click()
        self.driver.find_element("xpath","(//span[normalize-space()='Enviar um envelope'])[1]").click()

    def anexar_PDF(self, c):
        self.driver.find_element(By.CLASS_NAME,"css-89bprp").send_keys(c)

    def contar_elementos_template(self):   
        ## [ INICIO ] Conta a quantidade de campos de template ##

        #Lista os elementos que propuserem escolha de template
        self.elementos_template = self.driver.find_elements("xpath","(//span[normalize-space()='Selecione os modelos correspondentes'])[1]")
        #Conta quantos elementos foram encontrados em elementos_template
        self.quantidade_elementos_template = len(self.elementos_template)
        print("Aviso: Quantidade de elementos_template encontrados: "+ str(self.quantidade_elementos_template))
            
        ## [ FIM ] Conta a quantidade de campos de template ##

    def fechar_template(self):   
        ## [ INICIO ] Verifica a quantidade de elementos_template ##

        tentativas_fechar_template = 20
        # ENQUANTO a quantidade de quantidade_elementos_template for IGUAL A 0  (E)  a quantidade de tentativas_fechar_template for DIFERENTE DE 0
        while self.quantidade_elementos_template == 0 or tentativas_fechar_template != 0:
        # SE a quantidade de elementos_template for MAIOR que 0
            if self.quantidade_elementos_template > 0:
                # ENTÃO
                # Envia a variável auth para o campo de texto
                self.driver.find_element("xpath", "(//button[@class='olv-button olv-ignore-transform css-mtra3x'])[1]").click()
                print("Aviso: Template fechado com sucesso")
                tentativas_fechar_template = 0
                break
            else:
                tentativas_fechar_template = tentativas_fechar_template - 1
                print("Aviso " + str(tentativas_fechar_template) + " : O elemento de template não foi encontrado")
                self.contar_elementos_template()
                return self.quantidade_elementos_template

        ## [ FIM ] Verifica a quantidade de elementos_template ##
            
    def preencher_destinatários(self, n, e): 

        ## [ INICIO ] Preencher destinatários ##

        # Preenche o campo "Nome" com o nome recolhido em data_list
        self.driver.find_element("xpath", "(//input[@class='css-1ez4hss'])[1]").send_keys(n)
        # Preenche o campo "Email" com o email recolhido em data_list
        self.driver.find_element("xpath", "(//input[contains(@role,'combobox')])[2]").send_keys(e)

        ## [ FIM ] Preencher destinatários ##
    
    def preencher_id_de_venda(self, re): 

        ## [ INICIO ] Preencher "ID de Venda" ##

        # Preenche o campo "ID de Venda" com o nome recolhido em data_list
        self.driver.find_element("xpath", "(//input[@data-qa='label-input-ID de venda'])[1]").send_keys(re)

        # Clica no campo "Tipo de Documento"
        dropdowntipodedocumento = self.driver.find_element("xpath", "(//select[@class='css-12ihcxq'])[1]") # Instancia num objeto o elemento dropdown
        select = Select(dropdowntipodedocumento) # Seleciona o dropdown como uma caixa de elementos
        select.select_by_visible_text("Aditivo Contratual") # Procura pelo elemeto Aditivo Contratual e seleciona

        ## [ FIM ] Preencher "ID de Venda" ##
    
    def preencher_texto_email(self):

        ## [ INICIO ] Preencher "Mensagem do Email" ##    

        # Formatação do texto padrão
        texto = "Prezado(a)\n\n\nO período de férias foi programado conforme solicitado.\nSegue o aviso para assinatura\n\nObs: Não é necessário a entrega do aviso físico.\n\n\nAtenciosamente\nRH Whirlpool."
        # Envio do texto para o campo "Mensagem de Email"
        self.driver.find_element("xpath", "(//textarea[@placeholder='Inserir mensagem'])[1]").send_keys(texto)

        ## [ FIM ] Preencher "Mensagem do Email" ##  
    def clicar_em_avançar(self):

        ## [ INICIO ] Clicar no botão "Avançar" ##    

        # Clica no botão "Avançar"
        self.driver.find_element("xpath","(//button[@class='olv-button olv-ignore-transform css-1n4kelk'])[1]").click()  

        ## [ FIM ] Clicar no botão "Avançar" ##

    def contar_elementos_aviso_assinado(self):   
        ## [ INICIO ] Conta a quantidade de campos de template avisando que o pdf já está assinado ##

        #Lista os elementos que tragam a mensagem de que o aviso já está assinado
        self.elementos_aviso_assinado = self.driver.find_elements("xpath","(//span[normalize-space()='Gerenciar os dados do campo de formulário em PDF'])[1]")
        #Conta quantos elementos foram encontrados em elementos_aviso_assinado
        self.quantidade_elementos_aviso_assinado = len(self.elementos_aviso_assinado)
        print("Aviso: Quantidade de elementos_aviso_assinado encontrados: "+ str(self.quantidade_elementos_aviso_assinado))
            
        ## [ FIM ] Conta a quantidade de campos de template avisando que o pdf já está assinado ##

    def fechar_template_aviso_assinado(self):   
        ## [ INICIO ] Verifica a quantidade de elementos_aviso_assinado ##

        self.validaçãoavisoassinado = "" 

        tentativas_fechar_aviso_assinado = 5
        # ENQUANTO a quantidade de quantidade_elementos_aviso_assinado for IGUAL A 0  (E)  a quantidade de tentativas_fechar_aviso_assinado for DIFERENTE DE 0
        while self.quantidade_elementos_aviso_assinado == 0 or tentativas_fechar_aviso_assinado != 0:
        # SE a quantidade de elementos_aviso_assinado for MAIOR que 0
            if self.quantidade_elementos_aviso_assinado > 0:
                # ENTÃO
                # Preenche a variável validaçãoavisoassinado com TRUE
                self.validaçãoavisoassinado = "TRUE"
                tentativas_fechar_aviso_assinado = 0
                break
            else:
                # Preenche a variável validaçãoavisoassinado com FALSE
                tentativas_fechar_aviso_assinado = tentativas_fechar_aviso_assinado - 1
                print("Aviso " + str(tentativas_fechar_aviso_assinado) + " : O elemento de aviso já assinado não foi encontrado")
                self.contar_elementos_aviso_assinado()
                if self.quantidade_elementos_aviso_assinado == 0 and tentativas_fechar_aviso_assinado != 0:
                    self.validaçãoavisoassinado = "FALSE"
                    return self.quantidade_elementos_template

        ## [ FIM ] Verifica a quantidade de elementos_aviso_assinado ##