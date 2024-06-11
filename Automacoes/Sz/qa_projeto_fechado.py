from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import os
import datetime

# Testar na proxima atualização

final_wb = Workbook()
final_ws = final_wb.active
final_ws.title = "qa_projeto_fechado"

folder_path = filedialog.askdirectory()
folder = os.listdir(folder_path)

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

relatorio = None
update = None
for file in folder:
    file_path = os.path.join(folder_path, file)
    if "Relatório" in file:
        relatorio = load_workbook(filename=file_path, data_only=True)["QA_PROJETO_FECHADO"]
    if "jira" in file:
        update = load_workbook(filename=file_path, data_only=True).active

titles = ["Dia", "ID Atividade", "Sistema", "Tipo", "Status","Descrição", "ATENDENTE", "Data de Fechamento"]
for x in range(len(titles)):
    final_ws.cell(row=1, column=x+1, value=titles[x])

existentes = []
## Atualizar chamados antigos
for chamados in relatorio["B"][2:]:
    if chamados.value is None:
        break
    existentes.append(chamados.value)


# Adicionar Novos Chamados
for novos in update["A"][1:]:
    if novos.value not in existentes:
        print(novos.value)

    


#final_wb.save(f"{folder_path}/qa_projeto_fechado_att_{yesterday}.xlsx")
print("Processo Finalizado!")