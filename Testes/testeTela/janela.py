import customtkinter

# def button_callback():
#    print("button pressed")

# app = customtkinter.CTk()
# app.title = "my app"
# app.geometry("400x150") #define o tamanho da tela
# app.grid_columnconfigure((0, 1), weight=1) #Define como vai ser as posições dos itens na janela

# button = customtkinter.CTkButton(app, text="my button", command=button_callback)
# button.grid(row=0, column=0, padx=20, pady=20, sticky="ew", columnspan=2) #melhor do que usar o .pack por ser mais fácil de criar responsividade nas interfaces

# checkbox_1 = customtkinter.CTkCheckBox(app, text="checkbox 1")
# checkbox_1.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")
# checkbox_2 = customtkinter.CTkCheckBox(app, text="checkbox 2")
# checkbox_2.grid(row=1, column=1, padx=20, pady=(0, 20), sticky="w")

# app.mainloop()

class MyCheckboxFrame(customtkinter.CTkFrame):
    def __init__(self, master, values):
        super().__init__(master)
        self.values = values
        self.checkboxes = []
        
        for i, value in enumerate(self.values):
            checkbox = customtkinter.CTkCheckBox(self, text=value)
            checkbox.grid(row=i, column=0, padx=10, pady=(10, 0), sticky="w")
            self.checkboxes.append(checkbox)
    
    def get(self):
        checked_checkboxes = []
        for checkbox in self.checkboxes:
            if checkbox.get() == 1:
                checked_checkboxes.append(checkbox.cget("text"))
        return checked_checkboxes

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("my app")
        self.geometry("400x180")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.checkbox_frame = MyCheckboxFrame(self, values=["value 1", "value 2", "value 3"])
        self.checkbox_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="nsw")

        self.button = customtkinter.CTkButton(self, text="my button", command=self.button_callback)
        self.button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

    def button_callback(self):
        print("checked checkboxes: ", self.checkbox_frame.get())

app = App()
app.mainloop()