from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import pandas as pd
import datetime
import os

# Precisa automatizar isso?

file_path = filedialog.askopenfilename()
df = pd.read_csv(file_path, delimiter=";")
xlsx_path = f"data/convfile.xlsx"
df.to_excel(xlsx_path, index=False)
relatorio = load_workbook(xlsx_path, data_only=True).active
os.remove(xlsx_path)

final_wb = Workbook()
int_final_ws = final_wb.active
int_final_ws.title = "suporte_cliente_interno"
final_wb.create_sheet("suporte_cliente_externo")
ext_final_ws = final_wb["suporte_cliente_externo"]

int_titles =["Dia", "Chamado", "Sistema", "Area", "Categoria", "Status", "Atendente", "Descrição", "Data de Fechamento", "SLA Primeiro Atendimento"]
ext_titles =["Dia", "Chamado", "Sistema", "Cliente", "Categoria", "Status", "Descrição", "Atendente", "Data de Fechamento", "SLA Primeiro Atendimento"]

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

for x in range(len(int_titles)):
    int_final_ws.cell(row=1, column=x+1, value=int_titles[x])

for x in range(len(int_titles)):
    ext_final_ws.cell(row=1, column=x+1, value=ext_titles[x])

for cliente in relatorio["C"][1:]:
    new_row = int_final_ws.max_row+1
    if cliente.value == "SZ Soluções":
        int_final_ws.cell(row=new_row, column=1, value=yesterday)
        int_final_ws.cell(row=new_row, column=2, value= relatorio[f"A{cliente.row}"].value)
        int_final_ws.cell(row=new_row, column=3, value= cliente.value)
        int_final_ws.cell(row=new_row, column=4, value= "algo")
        int_final_ws.cell(row=new_row, column=5, value= relatorio[f"E{cliente.row}"].value)
        if relatorio[f"F{cliente.row}"].value in ("Encerrado"):
            int_final_ws.cell(row=new_row, column=6, value="Fechado")
        else:
            int_final_ws.cell(row=new_row, column=6, value="Aberto")
        int_final_ws.cell(row=new_row, column=7, value= relatorio[f"D{cliente.row}"].value)
        int_final_ws.cell(row=new_row, column=8, value= relatorio[f"B{cliente.row}"].value)
        int_final_ws.cell(row=new_row, column=9, value= relatorio[f"G{cliente.row}"].value)

        abertura = relatorio[f"O{cliente.row}"].value.split(":")
        atendimento = relatorio[f"N{cliente.row}"].value.split(":")
        print(atendimento, abertura)
        int_final_ws.cell(row=new_row, column=10, value= datetime.timedelta())
    else:
        ext_final_ws.cell(row=new_row, column=1, value=yesterday)
        ext_final_ws.cell(row=new_row, column=2, value= relatorio[f"A{cliente.row}"].value)
        if "Sonepar" in cliente.value:
            split = [x for x in cliente.value.split(" ") if len(x) > 1]
            ext_final_ws.cell(row=new_row, column=3, value= split[0])
            ext_final_ws.cell(row=new_row, column=4, value= split[1])
        ext_final_ws.cell(row=new_row, column=5, value= relatorio[f"E{cliente.row}"].value)
        if relatorio[f"F{cliente.row}"].value in ("Encerrado"):
            ext_final_ws.cell(row=new_row, column=6, value="Fechado")
        else:
            ext_final_ws.cell(row=new_row, column=6, value="Aberto")
        ext_final_ws.cell(row=new_row, column=7, value= relatorio[f"B{cliente.row}"].value)
        ext_final_ws.cell(row=new_row, column=8, value= relatorio[f"D{cliente.row}"].value)
        ext_final_ws.cell(row=new_row, column=9, value= relatorio[f"G{cliente.row}"].value)
        ext_final_ws.cell(row=new_row, column=10, value= datetime.timedelta())

final_wb.save(f"data/suporte_cliente.xlsx")