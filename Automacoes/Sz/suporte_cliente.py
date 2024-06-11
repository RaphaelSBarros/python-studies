from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import pandas as pd
import datetime
import os

# Precisa automatizar isso?

#file_path = filedialog.askopenfilename()
file_path = "C:/Users/P0589/Downloads/Aquivos teste/Suporte_Cliente/20240605095042.csv"
print(file_path)

final_wb = Workbook()
final_ws = final_wb.active
final_ws.title = "suporte_cliente"

titles =["Dia", "Chamado", "Sistema", "Area", "Categoria", "Status", "Atendente", "Descrição", "Data de Fechamento", "SLA Primeiro Atendimento"]

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

for x in range(len(titles)):
    final_ws.cell(row=1, column=x+1, value=titles[x])

df = pd.read_csv(filepath_or_buffer=file_path, sep=";")
xlsx_path=f"{file_path[:-3]}xlsx"
df.to_excel(xlsx_path, sheet_name="testing", index=False)
ws = load_workbook(filename=xlsx_path, data_only=True).active

for sistema in ws["C"][1:]:
    new_row = final_ws.max_row+1
    final_ws.cell(row=new_row, column=1, value=yesterday)
    final_ws.cell(row=new_row, column=2, value=ws[f"A{sistema.row}"].value)
    if ws[f"F{sistema.row}"].value in ("Fechado pelo usuário", "Encerrado"):
        final_ws.cell(row=new_row, column=6, value="Fechado")
    else:
        final_ws.cell(row=new_row, column=6, value="Aberto")

    if sistema.value == "SZ Soluções":
        final_ws.cell(row=new_row, column=7, value=ws[f"D{sistema.row}"].value)
    else:
        final_ws.cell(row=new_row, column=8, value=ws[f"D{sistema.row}"].value)
        final_ws.cell(row=new_row, column=4, value=ws[f"C{sistema.row}"].value)
        final_ws.cell(row=new_row, column=5, value=ws[f"E{sistema.row}"].value)

final_wb.save(f"teste.xlsx")