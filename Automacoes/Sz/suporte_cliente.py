''' Automação de Preenchimento das informações recebidas pelo SoftDesk.
    Passo a passo:
    1. Acessar o SoftDesk como Atendente
    2. Clicar no menu hamburguer
    3. Informação
    4. Clicar no lápis(Editar relatório)
    5. Avançar
    6. Alterar a data de abertura(sd_chamado.da_chamado) para o período que desejar
    7. Filtrar
    8. Excel
    9. Rode o programa
    10. Selecione o arquivo baixado
    11. Pronto, irá criar um novo arquivo .xlsx com as informações separadas em cliente interno e externo
'''

from module import *

# Precisa automatizar isso?

file_path = filedialog.askopenfilename()
df = pd.read_csv(file_path, delimiter=";")
NEW_FILE = "data/convfile.xlsx"
df.to_excel(NEW_FILE, index=False)
relatorio = load_workbook(NEW_FILE, data_only=True).active
os.remove(NEW_FILE)

final_wb = Workbook()
int_final_ws = final_wb.active
int_final_ws.title = "suporte_cliente_interno"
final_wb.create_sheet("suporte_cliente_externo")
ext_final_ws = final_wb["suporte_cliente_externo"]

int_titles =[
        "Aberto", 
        "Chamado", 
        "Sistema", 
        "Area", 
        "Categoria", 
        "Status", 
        "Atendente", 
        "Descrição", 
        "Data de Fechamento", 
        "SLA Primeiro Atendimento"
    ]
ext_titles =[
        "Aberto", 
        "Chamado",
        "ID Jira",
        "Sistema", 
        "Cliente", 
        "Categoria", 
        "Status", 
        "Descrição", 
        "Atendente", 
        "Data de Fechamento", 
        "SLA Primeiro Atendimento"
    ]

for x, title in enumerate(int_titles):
    int_final_ws.cell(row=1, column=x+1, value=title)

for x, title in enumerate(ext_titles):
    ext_final_ws.cell(row=1, column=x+1, value=title)

for cliente in relatorio["C"][1:]:
    int_new_row = int_final_ws.max_row+1
    ext_new_row = ext_final_ws.max_row+1

    iD,iM,iY,ihh,imm,iss = [int(x) for x in relatorio[f"M{cliente.row}"].value.split("/")]+[int(x) for x in relatorio[f"O{cliente.row}"].value.split(":")]
    idate = datetime.datetime(iY,iM,iD,ihh,imm,iss)

    try:
        fD,fM,fY,fhh,fmm,fss = [int(x) for x in relatorio[f"L{cliente.row}"].value.split("/")]+[int(x) for x in relatorio[f"N{cliente.row}"].value.split(":")]
        fdate = datetime.datetime(fY,fM,fD,fhh,fmm,fss)
    except:
        fdate = None

    fechamento = ""
    if relatorio[f"K{cliente.row}"].value is not None:
        fechamento = [int(x) for x in relatorio[f"K{cliente.row}"].value.split("/")]
        fechamento = datetime.datetime(fechamento[2], fechamento[1], fechamento[0])

    if cliente.value == "SZ Soluções":
        int_final_ws.cell(row=int_new_row,column=1,value= idate)
        int_final_ws.cell(row=int_new_row, column=2, value= relatorio[f"A{cliente.row}"].value)
        int_final_ws.cell(row=int_new_row, column=3, value= cliente.value)
        int_final_ws.cell(row=int_new_row, column=4, value= relatorio[f"E{cliente.row}"].value)
        int_final_ws.cell(row=int_new_row, column=5, value= relatorio[f"H{cliente.row}"].value)
        if fechamento == "":
            int_final_ws.cell(row=int_new_row, column=6, value="Aberto")
        else:   
            int_final_ws.cell(row=int_new_row, column=6, value="Fechado")
        int_final_ws.cell(row=int_new_row, column=7, value= relatorio[f"D{cliente.row}"].value)
        int_final_ws.cell(row=int_new_row, column=8, value= relatorio[f"B{cliente.row}"].value)
        int_final_ws.cell(row=int_new_row, column=9, value= fechamento)
        int_final_ws.cell(row=int_new_row, column=10,value= fdate - idate)
    else:
        ext_final_ws.cell(row=ext_new_row, column=1, value= idate)
        ext_final_ws.cell(row=ext_new_row, column=2, value= relatorio[f"A{cliente.row}"].value)
        ext_final_ws.cell(row=ext_new_row, column=3, value="") # ID JIRA
        if cliente.value is not None and "Sonepar" in cliente.value:
            split = [x for x in cliente.value.split(" ") if len(x) > 1]
            ext_final_ws.cell(row=ext_new_row, column=4, value= split[0])
            ext_final_ws.cell(row=ext_new_row, column=5, value= split[1])
        else:
            ext_final_ws.cell(row=ext_new_row, column=4, value=relatorio[f"I{cliente.row}"].value)
            ext_final_ws.cell(row=ext_new_row, column=5, value= relatorio[f"C{cliente.row}"].value)
        ext_final_ws.cell(row=ext_new_row, column=6, value= relatorio[f"E{cliente.row}"].value)
        if fechamento == "":
            ext_final_ws.cell(row=ext_new_row, column=7, value="Aberto")
        else:
            ext_final_ws.cell(row=ext_new_row, column=7, value="Fechado")
        ext_final_ws.cell(row=ext_new_row, column=8, value= relatorio[f"B{cliente.row}"].value)
        ext_final_ws.cell(row=ext_new_row, column=9, value= relatorio[f"D{cliente.row}"].value)
        ext_final_ws.cell(row=ext_new_row, column=10, value= fechamento)
        if fdate and idate:
            ext_final_ws.cell(row=ext_new_row, column=11, value = fdate - idate)

final_wb.save("data/suporte_cliente.xlsx")
