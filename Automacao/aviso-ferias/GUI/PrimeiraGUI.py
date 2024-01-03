from tkinter import Tk, Label, Entry, Button, messagebox

class ChamarPrimeiraUI():
    def __init__(self):

        #Esquema de cores Drácula 
        self.backgroundcolor = "#282a36"
        self.foregroundcolor = "#f8f8f2"
        self.buttoncolor = "#50fa7b"

        self.menu_inicial = Tk()
        self.menu_inicial.title("Acesso ao Docusign")
        self.menu_inicial['bg'] = self.backgroundcolor

        # Adiciona um rótulo para instruções
        label_email = Label(self.menu_inicial,
                            text="Confirme os dados de login!\nEmail padrão: datamanagementbrazil@whirlpool.com\n",
                            font="Arial 10",
                            fg=self.foregroundcolor,
                            bg=self.backgroundcolor,
                            anchor="w")
        label_email.pack()

        # Adiciona um rótulo para instruções
        label_senha = Label(self.menu_inicial, 
                            text="Digite a senha",
                            font="Arial 10",
                            fg=self.foregroundcolor,
                            bg=self.backgroundcolor,
                            anchor="w")
        label_senha.pack()
        
        # Adiciona uma caixa de entrada de texto para a senha
        self.entry_senha = Entry(self.menu_inicial, 
                                 show='*',
                                 font="Arial 10")
        self.entry_senha.pack()

        # Adiciona um botão para acionar a função que obtém a senha
        Button(self.menu_inicial, 
               text="Continuar", 
               command=self.obter_credenciais,
               font="Arial 10",
               fg="black",
               bg=self.buttoncolor).pack(side="bottom", anchor="se")

        # Inicia o loop principal da interface gráfica
        self.menu_inicial.mainloop()

    def obter_credenciais(self):
        # Obtém a senha da caixa de entrada
        self.senha = self.entry_senha.get()

        # Verifica se o campo de senha está vazio
        if not self.senha:
            # Mostra uma janela de erro
            messagebox.showerror("Erro", "Por favor, insira a senha.")
        else:
            def obter_senha(self):
                # Retorna a senha da caixa de entrada
                return self.senha
        # Encerra a aplicação
        self.menu_inicial.destroy()