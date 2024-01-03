import time
from Módulos.Modulacao import AutenticarAcesso
from GUI.PrimeiraGUI import ChamarPrimeiraUI
from GUI.SegundaGUI import ChamarSegundaUI
from testeExcel.dialogbox4 import App
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()  # Cria a instância principal da interface gráfica
    app_instance = App(root)  # Cria a instância da classe App
    root.mainloop()  # Inicia o loop principal da interface gráfica

### [ INICIO ] Verificando necessidade de login ###

#Chama uma GUI para solicitar a senha
chamar_módulo1 = ChamarPrimeiraUI()
#Define a senha  informada na GUI para uma variável
senha = chamar_módulo1.senha

# Cria uma instância da classe AutenticarAcesso
chamar_módulo0 = AutenticarAcesso()

# Chama o método para verificar necessidade de login
chamar_módulo0.abrir_chrome_com_perfil()
# Chama o método para verificar email
chamar_módulo0.verificar_autenticação_email()
time.sleep(2)
# Chama o método para verificar senha
chamar_módulo0.verificar_autenticação_senha(senha)
time.sleep(2)
# Chama o método para verificar se o código auth é necessário
chamar_módulo0.contar_elementos_auth()
if chamar_módulo0.quantidade_elementos_auth > 0:
    chamar_módulo2 = ChamarSegundaUI()
    auth_code = chamar_módulo2.auth
    # Chama o método para verificar auth
    chamar_módulo0.verificar_autenticação_auth(auth_code)
else:
    print("Aviso: O elemento de auth não foi encontrado ou requirido")
    pass
### [ FIM ] Verificando necessidade de login ###

### [ INICIO ] Criando envelope

índice = 0

# Supondo que `app.data_list` já tenha sido preenchida no código anterior
for data in app_instance.data_list:

    # A primeira linha contém os nomes das colunas, então você pode pular essa linha
    if "Path" not in data:
        continue
    
    # Acessa a página para criar envelopes ( parte inicial do loop )
    chamar_módulo0.driver.get("https://apps.docusign.com/send/documents")
    # Obtém o caminho do PDF associado à linha
    caminho_pdf = data["Path"]

    # Criando envelope
    time.sleep(5)
    chamar_módulo0.abrir_envelope()

    # Realizando upload dos PDFs
    time.sleep(5)
    chamar_módulo0.anexar_PDF(caminho_pdf)
    time.sleep(10)

    # Verificando se o pop-up de templates irá abrir
    chamar_módulo0.contar_elementos_template()

    # Caso o pop-up abra, fecha o pop-up
    chamar_módulo0.fechar_template()

    # Preencher Destinatário
    chamar_módulo0.preencher_destinatários(data["Nome Destinatário"], data["Email Destinatário"])

    # Preencher ID de Venda e Tipo de Documento
    chamar_módulo0.preencher_id_de_venda(data["RE Colaborador"])

    # Preencher "Mensagem de Email"
    chamar_módulo0.preencher_texto_email()

    # Clica no botão "Avançar"
    chamar_módulo0.clicar_em_avançar()

    # Verificando se o pop-up de templates irá abrir
    chamar_módulo0.contar_elementos_aviso_assinado()
    time.sleep(3)

    # Caso o pop-up abra, fecha o pop-up
    chamar_módulo0.fechar_template_aviso_assinado()

    # Trazer a validação de email já assinado
    validaçãoavisoassinado = chamar_módulo0.validaçãoavisoassinado
    # Se a validação for TRUE
    if validaçãoavisoassinado == "TRUE":
        print(validaçãoavisoassinado)
        # Adiciona a mensagem de confirmação à última coluna do dicionário
        data["Status"] = "Erro: Aviso de Férias em anexo já assinado, o aviso não foi enviado."
        print(str(índice))
        print(app_instance.data_list[índice])
        índice = índice + 1
    # Adiciona a mensagem de confirmação à última coluna do dicionário
    data["Status"] = "Aviso Enviado"
    
    