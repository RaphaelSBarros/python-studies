from tkinter import Tk, Label, Entry, Button, messagebox

class ChamarSegundaUI():
    def __init__(self):

        #Esquema de cores Drácula 
        self.backgroundcolor = "#282a36"
        self.foregroundcolor = "#f8f8f2"
        self.buttoncolor = "#50fa7b"

        self.menu_inicial = Tk()
        self.menu_inicial.title("Código de Verificação solicitado")
        self.menu_inicial['bg'] = self.backgroundcolor

        # Adiciona um rótulo para instruções
        label_auth = Label(self.menu_inicial, 
                            text="O Código de verificação via email foi solicitado!\nDigite o código recebido no email na caixa a seguir e clique em Continuar",
                            font="Arial 10",
                            fg=self.foregroundcolor,
                            bg=self.backgroundcolor,
                            anchor="w")
        label_auth.pack()
        
        # Adiciona uma caixa de entrada de texto para o código
        self.entry_auth = Entry(self.menu_inicial, 
                                 show='*',
                                 font="Arial 10")
        self.entry_auth.pack()

        # Adiciona um botão para acionar a função que obtém o auth
        Button(self.menu_inicial, 
               text="Continuar", 
               command=self.obter_credenciais,
               font="Arial 10",
               fg="black",
               bg=self.buttoncolor).pack(side="bottom", anchor="se")

        # Inicia o loop principal da interface gráfica
        self.menu_inicial.mainloop()

    def obter_credenciais(self):
        # Obtém o auth da caixa de entrada
        self.auth = self.entry_auth.get()

        # Verifica se o campo de auth está vazio
        if not self.auth:
            # Mostra uma janela de erro
            messagebox.showerror("Erro", "Por favor, insira o código.")
        else:
            def obter_senha(self):
                # Retorna a senha da caixa de entrada
                return self.auth
        # Encerra a aplicação
        self.menu_inicial.destroy()