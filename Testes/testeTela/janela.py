import customtkinter

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

janela = customtkinter.CTk()
janela.geometry("500x300")

def Clique():
    print("Logando...")
    
texto = customtkinter.CTkLabel(janela, text="Login")
texto.pack(padx=10, pady=10)

email = customtkinter.CTkEntry(janela, placeholder_text="Seu e-mail")
email.pack(padx=10, pady=10)

senha = customtkinter.CTkEntry(janela, placeholder_text="Sua senha", show="*")
senha.pack(padx=10, pady=10)

checkbox = customtkinter.CTkCheckBox(janela, text="Manter login")
checkbox.pack(padx=10, pady=10)

botao = customtkinter.CTkButton(janela, text="Login", command=Clique)
botao.pack(padx=10, pady=10)

janela.mainloop()