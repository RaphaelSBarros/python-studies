from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import pandas as pd
import os
import datetime

folder_path = filedialog.askdirectory()
folder = os.listdir(folder_path)

final_wb = Workbook()
final_ws = final_wb.active

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

relatorio = None
update = None
for file in folder:
    file_path = os.path.join(folder_path, file)
    if "Relatório" in file:
        b = pd.read_excel(io=file_path, sheet_name="SUPORTE_SONEPAR")
        convfile="data/convfile.xlsx"
        b.to_excel(convfile, index=False)
        relatorio = load_workbook(filename=convfile, data_only=True).active
        os.remove(convfile)
    if "att" in file:
        update = load_workbook(filename=file_path, data_only=True).active

titles = ["Dia", "ID Chamado", "Sistema", "Cliente", "Categoria",
          "Status", "Descrição", "Atendente", "Data Fechamento",
          "Aguardando", "Motivo", "SLA de Resolução %", "SLA de Resolução Tempo"] # Título das colunas
for x in range(len(titles)):
    final_ws.cell(row=1, column=x+1, value=titles[x]) # Preenche os títulos nas colunas da planilha final

existentes = [] # Armazenando quais os chamados existentes no relatório
updt_count=0 # Contagem dos chamados com atualização
for chamado in relatorio["B"][2:]:
    if chamado.value is None:
        break
    existentes.append(chamado.value)
    new_row = final_ws.max_row+1
    for row in final_ws[f"A{new_row}:M{new_row}"]:
        for cell in row:
            final_ws.cell(row=new_row, column=cell.column, value=relatorio.cell(row=chamado.row, column=cell.column).value)
    for atualizar in update["B"][1:]:
        if atualizar.value == chamado.value and update[f"E{atualizar.row}"].value != relatorio[f"F{chamado.row}"].value:
            updt_count+=1
            final_ws.cell(row=new_row, column=6, value=update[f"E{atualizar.row}"].value)
            final_ws.cell(row=new_row, column=9, value=update[f"G{atualizar.row}"].value)
            final_ws.cell(row=new_row, column=10, value=update[f"H{atualizar.row}"].value)
            final_ws.cell(row=new_row, column=11, value=update[f"I{atualizar.row}"].value)
            final_ws.cell(row=new_row, column=12, value=update[f"J{atualizar.row}"].value)
            final_ws.cell(row=new_row, column=13, value=update[f"K{atualizar.row}"].value)
            print(f"Chamado Atualizado: {chamado.value}")

add_count = 0
for novos in update["B"][1:]:
    if novos.value not in existentes:
        add_count+=1
        new_row = final_ws.max_row+1
        for row in final_ws[f"A{new_row}:D{new_row}"]:
            for cell in row:
                final_ws.cell(row=new_row, column=cell.column, value=update.cell(row=novos.row, column=cell.column).value)
        for row in final_ws[f"F{new_row}:M{new_row}"]:
            for cell in row:
                final_ws.cell(row=new_row, column=cell.column, value=update.cell(row=novos.row, column=cell.column-1).value)
        print(f"Chamado Adicionado: {novos.value}")

final_wb.save(f"{folder_path}/atualizacao_sonepar.xlsx") # Salva a nova tabela na mesma pasta que estavam os relatórios
print(f"Projeto Finalizado com {updt_count} atualizações e {add_count} adições nos dados") # Apenas para conferência do total de informações