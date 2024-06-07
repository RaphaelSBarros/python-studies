from tkinter import filedialog
from openpyxl import load_workbook, Workbook
final_wb = Workbook()
final_ws = final_wb.active
final_ws.title = "sonepar_support_report"

#file_path = filedialog.askopenfilename()
bulk_sla = load_workbook(filename="C:/Users/P0589/Downloads/task_sla.xlsx", data_only=True).active # Arquivo de SLA contendo os dados de itens requisitados e Incidentes
#file_path = filedialog.askopenfilename()
incident_data = load_workbook(filename="C:/Users/P0589/Downloads/incident.xlsx", data_only=True).active # Dados dos Incidentes
#file_path = filedialog.askopenfilename()
requisiton_data = load_workbook(filename="C:/Users/P0589/Downloads/sc_req_item.xlsx", data_only=True).active # Dados dos Itens Requisitados

# Data de Abrtura / Nº Chamado / Sistema / Cliente / Status (calcular) / Descrição / Atendente / Data Fechamento / Aguardando / Motivo / SLA de Resolução / SLA de Resolução Tempo (calcular)
titles = ["Dia", "ID Chamado", "Sistema", "Cliente",
          "Status", "Descrição", "Atendente", "Data Fechamento",
          "Aguardando", "Motivo", "SLA de Resolução %", "SLA de Resolução Tempo"] # Título das colunas
for x in range(12):
    final_ws.cell(row=1, column=x+1, value=titles[x]) # Preenche os títulos nas colunas da planilha final

for id_inc in incident_data["B"][1:]: # Percorre todos os IDs de chamados na Coluna B da planilha
    new_row = final_ws.max_row+1 # Próxima linha vazia a ser inserido os dados
    x=1
    for row in incident_data[f"A{id_inc.row}:D{id_inc.row}"]: # Preenche os dados das Colunas A, B, C D da planilha de incidentes na planilha final
        for cell in row:
            final_ws.cell(row=new_row, column=x, value=cell.value)
            x+=1
    if incident_data[f"G{id_inc.row}"].value == None: # Faz a verificação do Status
        final_ws.cell(row=new_row, column=x, value="Aberto")
        x+=1
    else:
        final_ws.cell(row=new_row, column=x, value="Fechado")
        x+=1

    for row in incident_data[f"E{id_inc.row}:I{id_inc.row}"]: # Preenche os dados das Colunas E, F, G, H, I da planilha de incidentes na planilha 
        for cell in row:
            final_ws.cell(row=new_row, column=x, value=cell.value)
            x+=1
    for id_chamado in bulk_sla["A"][1:]: # Procura o mesmo Id de chamado dentro da planilha de SLA
        if id_chamado.value == id_inc.value:
            final_ws.cell(row=new_row, column=x, value=bulk_sla[f"I{id_chamado.row}"].value) # Insere o valor da porcentagem na planilha final
            x+=1
            final_ws.cell(row=new_row, column=x, value=(bulk_sla[f"H{id_chamado.row}"].value/86400)) # Faz o cálculo do tempo e insere na planilha final


# Faz a mesma coisa só que para os dados de Itens Requisitados
for id_ritm in requisiton_data["B"][1:]:
    new_row = final_ws.max_row+1 # Próxima linha vazia a ser inserido os dados
    x=1
    for row in requisiton_data[f"A{id_ritm.row}:D{id_ritm.row}"]:
        for cell in row:
            final_ws.cell(row=new_row, column=x, value=cell.value)
            x+=1
    if requisiton_data[f"G{id_ritm.row}"].value == None:
        final_ws.cell(row=new_row, column=x, value="Aberto")
        x+=1
    else:
        final_ws.cell(row=new_row, column=x, value="Fechado")
        x+=1
    for row in requisiton_data[f"E{id_ritm.row}:H{id_ritm.row}"]:
        for cell in row:
            final_ws.cell(row=new_row, column=x, value=cell.value)
            x+=1
    for id_chamado in bulk_sla["A"][1:]:
        if id_chamado.value == id_ritm.value:
            final_ws.cell(row=new_row, column=x, value=bulk_sla[f"I{id_chamado.row}"].value)
            x+=1
            final_ws.cell(row=new_row, column=x, value=(bulk_sla[f"H{id_chamado.row}"].value/86400))

final_wb.save("Sample.xlsx")