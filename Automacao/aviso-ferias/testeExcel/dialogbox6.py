import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk

class App0:
    def __init__(self, root):
        self.root = root
        self.root.title("Enviar & Recolher Avisos de Férias")

        self.file_path_label = tk.Label(root, text="Qual processo você deseja realizar?")
        self.file_path_label.pack(pady=10)

        self.browse_button = tk.Button(root, text="Enviar Avisos", command=self.enviar_avisos)
        self.browse_button.pack(pady=10)

    def enviar_avisos(self):
        print("!!!!!!!!!!111")

root = tk.Tk()
app_instance = App0(root)
root.mainloop()