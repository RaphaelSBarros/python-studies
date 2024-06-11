from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import pandas as pd
import os
import datetime

# Testar na proxima atualização
folder_path = filedialog.askdirectory()
folder = os.listdir(folder_path)
inicio1 = datetime.datetime.now()
final_wb = Workbook()
final_ws = final_wb.active
final_ws.title = "qa_projeto_fechado"

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

relatorio = None
update = None

for file in folder:
    file_path = os.path.join(folder_path, file)
    if "Relatório" in file:
        b = pd.read_excel(io=file_path, sheet_name="QA_PROJETO_FECHADO")
        convfile="data/convfile.xlsx"
        b.to_excel(convfile, index=False)
        relatorio = load_workbook(filename=convfile, data_only=True).active
        os.remove(convfile)
    if "jira" in file:
        update = load_workbook(filename=file_path, data_only=True).active

titles = ["Dia", "ID Atividade", "Sistema", "Tipo", "Status","Descrição", "ATENDENTE", "Data de Fechamento"]
for x in range(len(titles)):
    final_ws.cell(row=1, column=x+1, value=titles[x])

existentes = []
## Atualizar chamados antigos
updt_count=0
for chamado in relatorio["B"][2:]:
    if chamado.value is None:
        break
    existentes.append(chamado.value)
    new_row = final_ws.max_row+1
    final_ws.cell(row=new_row, column=1, value=relatorio[f"A{chamado.row}"].value)
    final_ws.cell(row=new_row, column=2, value=chamado.value)
    final_ws.cell(row=new_row, column=3, value=relatorio[f"C{chamado.row}"].value)
    final_ws.cell(row=new_row, column=4, value=relatorio[f"D{chamado.row}"].value)
    final_ws.cell(row=new_row, column=5, value=relatorio[f"E{chamado.row}"].value)
    final_ws.cell(row=new_row, column=6, value=relatorio[f"F{chamado.row}"].value)
    final_ws.cell(row=new_row, column=7, value=relatorio[f"G{chamado.row}"].value)
    final_ws.cell(row=new_row, column=8, value=relatorio[f"H{chamado.row}"].value)
    for atualizar in update["A"][1:]:
        if chamado.value == atualizar.value:
            if update[f"F{atualizar.row}"].value in ("Concluído", "DEPLOY"):
                final_ws.cell(row=new_row, column=5, value="Fechado")
                final_ws.cell(row=new_row, column=8, value=update[f"O{atualizar.row}"].value)
            else:
                final_ws.cell(row=new_row, column=5, value="Aberto")
            updt_count+=1

add_count = 0
# Adicionar Novos Chamados
inicio2 = datetime.datetime.now()
for novos in update["A"][1:]:
    if novos.value not in existentes:
        new_row = final_ws.max_row+1
        final_ws.cell(row=new_row, column=1, value=yesterday)
        final_ws.cell(row=new_row, column=2, value=novos.value)
        final_ws.cell(row=new_row, column=3, value=update[f"D{novos.row}"].value)
        final_ws.cell(row=new_row, column=4, value=update[f"E{novos.row}"].value)
        if update[f"F{novos.row}"].value in ("Concluído", "DEPLOY"):
            final_ws.cell(row=new_row, column=5, value="Fechado")
        else:
            final_ws.cell(row=new_row, column=5, value="Aberto")
        final_ws.cell(row=new_row, column=6, value=update[f"B{novos.row}"].value)
        final_ws.cell(row=new_row, column=7, value=update[f"M{novos.row}"].value)
        final_ws.cell(row=new_row, column=8, value=update[f"O{novos.row}"].value)
        add_count+=1

final_wb.save(f"{folder_path}/qa_projeto_fechado_att_{yesterday}.xlsx")
fim1 = datetime.datetime.now()
print(f"Primeiro Processo Finalizado em: {fim1 - inicio1}")