'''Automação de Preenchimento das informações recebidas pelo ServiceNow.
    Passo a passo: 
    1. Baixar as informações dos Incidentes da Query: https://soneparprod.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3Dopened_atONYesterday@javascript:gs.beginningOfYesterday()@javascript:gs.endOfYesterday()%5Ebusiness_service%3D30318b1bc35b7954c354254ce00131d3%26sysparm_first_row%3D1%26sysparm_view%3Dess
    2. Baixar as informações das Requisições da Query: https://soneparprod.service-now.com/nav_to.do?uri=%2Fsc_req_item_list.do%3Fsysparm_query%3Drequest.opened_atONYesterday@javascript:gs.beginningOfYesterday()@javascript:gs.endOfYesterday()%5Ebusiness_service%3D30318b1bc35b7954c354254ce00131d3%26sysparm_first_row%3D1%26sysparm_view%3Dess
    ***Caso for baixar as informações em uma segunda-feira, terá que alterar a data de abertura para entre sexta e domingo
    3. Baixas as informações do SLA das tarefas: https://soneparprod.service-now.com/nav_to.do?uri=sys_report_template.do%3Fjvar_report_id%3D6ce3ea2aebe64e543e0af095dad0cdd9
    *** O período precisa ser o mesmo dos incidentes e requisiçoes
    4. Colocar os 3 arquivos numa mesma pasta
    5. Rode o programa
    6. Selecione a pasta em que os arquivos se encontram
    7. Pronto, as informações serão agrupadas e inseridas o SLA em cada chamado.
'''

from module import *

final_wb = Workbook()
final_ws = final_wb.active
final_ws.title = "sonepar_support_report"

yesterday=(datetime.date.today() - datetime.timedelta(days=1))
if yesterday.strftime("%A") == "Sunday":
    yesterday=(datetime.date.today() - datetime.timedelta(days=3))

FOLDER_PATH = filedialog.askdirectory()
folder = os.listdir(FOLDER_PATH)
for file in folder:
    file_path = os.path.join(FOLDER_PATH, file)
    if "task_sla" in file:
        bulk_sla = load_workbook(filename= file_path, data_only=True).active
    if "incident" in file:
        incident_data = load_workbook(filename= file_path, data_only=True).active
    if "sc_req_item" in file:
        requisiton_data = load_workbook(filename= file_path, data_only=True).active

# Título das colunas
titles = ["Aberto",
          "Número",
          "IC Afetado",
          "Empresa Afetada",
          "Aberto por",
          "Descrição resumida",
          "Atribuído a",
          "Encerrado",
          "Estado",
          "SLA de Resolução %",
          "SLA de Resolução Tempo"
        ]

# Preenche os títulos nas colunas da planilha final
for x, title in enumerate(titles):
    final_ws.cell(row=1, column= x+1, value=title)

for id_inc in incident_data["B"][1:]: # Percorre todos os IDs de chamados na Coluna B da planilha
    new_row = final_ws.max_row+1 # Próxima linha vazia a ser inserido os dados
    COLUMN=1
    # Preenche os dados das Colunas A, B, C D da planilha de incidentes na planilha final
    for row in incident_data[f"A{id_inc.row}:I{id_inc.row}"]:
        for cell in row:
            final_ws.cell(row=new_row, column=COLUMN, value=cell.value)
            COLUMN+=1
    for id_chamado in bulk_sla["A"][1:]: # Procura o mesmo Id de chamado dentro da planilha de SLA
        if id_chamado.value == id_inc.value:
            # Insere o valor da porcentagem na planilha final
            final_ws.cell(row=new_row, column=COLUMN,
                          value=bulk_sla[f"I{id_chamado.row}"].value)
            # Faz o cálculo do tempo e insere na planilha final
            final_ws.cell(row=new_row, column=COLUMN+1,
                          value=bulk_sla[f"H{id_chamado.row}"].value/86400)

# Faz a mesma coisa só que para os dados de Itens Requisitados
for id_ritm in requisiton_data["B"][1:]:
    new_row = final_ws.max_row+1 # Próxima linha vazia a ser inserido os dados
    COLUMN=1
    for row in requisiton_data[f"A{id_ritm.row}:I{id_ritm.row}"]:
        for cell in row:
            final_ws.cell(row=new_row, column=COLUMN, value=cell.value)
            COLUMN+=1
    for id_chamado in bulk_sla["A"][1:]:
        if id_chamado.value == id_ritm.value:
            final_ws.cell(row=new_row, column=COLUMN, value=bulk_sla[f"I{id_chamado.row}"].value)
            final_ws.cell(row=new_row, column=COLUMN+1, value=bulk_sla[f"H{id_chamado.row}"].value/86400)

#os.remove(f"{FOLDER_PATH}/task_sla.xlsx")
#os.remove(f"{FOLDER_PATH}/incident.xlsx")
#os.remove(f"{FOLDER_PATH}/sc_req_item.xlsx")

final_wb.save(f"{FOLDER_PATH}/suporte_sonepar_att_{yesterday}.xlsx")
print("Processo Finalizado!")
