import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to PDF Finder")

        self.file_path_label = tk.Label(root, text="Selecione o arquivo CSV:")
        self.file_path_label.pack(pady=10)

        self.file_path_entry = tk.Entry(root, width=50)
        self.file_path_entry.pack(pady=10)

        self.browse_button = tk.Button(root, text="Procurar", command=self.browse_csv)
        self.browse_button.pack(pady=10)

        self.pdf_folder_label = tk.Label(root, text="Selecione a pasta com os PDFs:")
        self.pdf_folder_label.pack(pady=10)

        self.pdf_folder_entry = tk.Entry(root, width=50)
        self.pdf_folder_entry.pack(pady=10)

        self.browse_pdf_button = tk.Button(root, text="Procurar", command=self.browse_pdf_folder)
        self.browse_pdf_button.pack(pady=10)

        self.process_button = tk.Button(root, text="Processar", command=self.process_data)
        self.process_button.pack(pady=10)

        self.tree = ttk.Treeview(root)
        self.tree["columns"] = ("Path",)
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Path", anchor=tk.W, width=400)
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading("Path", text="Path")
        self.tree.pack(pady=10)

        # Lista para armazenar as informações do CSV e os caminhos dos arquivos PDF
        self.data_list = []

    def browse_csv(self):
        self.root.after(1, self.ask_csv_filename)

    def ask_csv_filename(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.file_path_entry.delete(0, tk.END)
        self.file_path_entry.insert(0, file_path)

    def browse_pdf_folder(self):
        self.root.after(1, self.ask_pdf_folder)

    def ask_pdf_folder(self):
        pdf_folder = filedialog.askdirectory()
        self.pdf_folder_entry.delete(0, tk.END)
        self.pdf_folder_entry.insert(0, pdf_folder)

    def process_data(self):
        csv_path = self.file_path_entry.get()
        pdf_folder = self.pdf_folder_entry.get()

        if not csv_path or not pdf_folder:
            return

        df = pd.read_csv(csv_path)

        column_names = list(df.columns)

        self.tree["columns"] = tuple(column_names + ["Path"])
        self.tree.delete(*self.tree.get_children())

        for column_name in column_names:
            self.tree.column(column_name, anchor=tk.W, width=150)
            self.tree.heading(column_name, text=column_name)

        self.tree.column("Path", anchor=tk.W, width=400)
        self.tree.heading("Path", text="Path")


        self.root.update_idletasks()

        # Limpa a lista antes de adicionar novos dados
        self.data_list = []

        for _, row in df.iterrows():
            values = tuple(str(row[column_name]) for column_name in column_names)
            self.pdf_file = self.find_pdf(pdf_folder, str(row[column_names[0]]))
            self.tree.insert("", "end", values=values + (self.pdf_file,))
            
            # Adiciona as informações à lista
            data_dict = dict(zip(column_names, values))
            data_dict["Path"] = self.pdf_file
            self.data_list.append(data_dict)
            print(self.data_list)
            self.root.update_idletasks()

    def find_pdf(self, folder, info):
        for file in os.listdir(folder):
            if info in file:
                return os.path.join(folder, file)
        return "PDF não encontrado"