import time
from Módulos.Modulacao import AutenticarAcesso
from GUI.PrimeiraGUI import ChamarPrimeiraUI
from GUI.SegundaGUI import ChamarSegundaUI
from testeExcel.dialogbox4 import App
import tkinter as tk

class AcessarDocusign():
    def acessar_docusign(self):

        ### [ INICIO ] Verificando necessidade de login ###qwiuifisd 
        #Chama uma GUI para solicitar a senha
        self.chamar_módulo1 = ChamarPrimeiraUI()
        #Define a senha  informada na GUI para uma variável
        senha = self.chamar_módulo1.senha

        # Cria uma instância da classe AutenticarAcesso
        self.chamar_módulo0 = AutenticarAcesso()

        # Chama o método para verificar necessidade de login
        self.chamar_módulo0.abrir_chrome_com_perfil()
        # Chama o método para verificar email
        self.chamar_módulo0.verificar_autenticação_email()
        time.sleep(2)
        # Chama o método para verificar senha
        self.chamar_módulo0.verificar_autenticação_senha(senha)
        time.sleep(2)
        # Chama o método para verificar se o código auth é necessário
        self.chamar_módulo0.contar_elementos_auth()
        if self.chamar_módulo0.quantidade_elementos_auth > 0:
            chamar_módulo2 = ChamarSegundaUI()
            auth_code = chamar_módulo2.auth
            # Chama o método para verificar auth
            self.chamar_módulo0.verificar_autenticação_auth(auth_code)
        else:
            print("Aviso: O elemento de auth não foi encontrado ou requirido")
            pass
        ### [ FIM ] Verificando necessidade de login ###

        ### [ INICIO ] Criando envelope

        if __name__ == "__main__":
            root = tk.Tk()  # Cria a instância principal da interface gráfica
            app_instance = App(root)  # Cria a instância da classe App
            root.mainloop()  # Inicia o loop principal da interface gráfica

        # Supondo que `app.data_list` já tenha sido preenchida no código anterior
        for data in app_instance.data_list:
            # A primeira linha contém os nomes das colunas, então você pode pular essa linha
            if "Path" not in data:
                continue
            # Acessa a página para criar envelopes ( parte inicial do loop )
            self.chamar_módulo0.driver.get("https://apps.docusign.com/send/documents")
            # Obtém o caminho do PDF associado à linha
            caminho_pdf = data["Path"]

            # Criando envelope
            time.sleep(10)
            self.chamar_módulo0.abrir_envelope()

            # Realizando upload dos PDFs
            time.sleep(10)
            self.chamar_módulo0.anexar_PDF(caminho_pdf)
            time.sleep(10)
            self.chamar_módulo0.contar_elementos_template()
            self.chamar_módulo0.fechar_template()

            # Adiciona a mensagem de confirmação à última coluna do dicionário
            data["Status"] = "Aviso Enviado"
            print(app_instance.data_list)

        ### [ FIM ] Criando envelope
            
obj = AcessarDocusign()
obj.acessar_docusign()