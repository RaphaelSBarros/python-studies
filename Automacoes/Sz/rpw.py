'''
    Extração dos RPWs e alocação em linhas de planilhas
'''

from module import *

file_path = filedialog.askopenfilename()
rpw_nortel = load_workbook(filename=file_path, data_only=True)['Controle Erros RPW Nortel']
rpw_dime = load_workbook(filename=file_path, data_only=True)['Controle Erros RPW Dimensional']

wb = Workbook()
ws = wb.active
ws.title = "RPW"

titles = ["Data",
          "RPW Afetado",
          "Empresa",
          "Resolução",
          "Identificado",
          "Identificado por"
          ]

for x, title in enumerate(titles):
    ws.cell(row=1, column=x+1, value=title)

for row in rpw_nortel["B"][1:]:
    if row.value is None:
        continue
    if row.value != 0 and "," in row.value:
        rpws = row.value.replace(" ", "").split(",")
        for rpw in rpws:
            if len(rpw) > 1 and "/" not in rpw:
                new_row = ws.max_row+1
                ws.cell(row=new_row, column=1, value=rpw_nortel[f"A{row.row}"].value)
                ws.cell(row=new_row, column=2, value=rpw.upper())
                if "NOR" in rpw:
                    ws.cell(row=new_row, column=3, value="Nortel")
                else:
                    ws.cell(row=new_row, column=3, value="Dimensional")
                ws.cell(row=new_row, column=4, value=rpw_nortel[f"C{row.row}"].value)
                ws.cell(row=new_row, column=5, value=rpw_nortel[f"D{row.row}"].value)
                ws.cell(row=new_row, column=6, value=rpw_nortel[f"E{row.row}"].value)
for row in rpw_dime["B"][1:]:
    if row.value is None:
        break
    if row.value != 0 and "," in row.value:
        rpws = row.value.replace(",", "").split(" ")
        for rpw in rpws:
            if len(rpw) > 1 and "/" not in rpw:
                new_row = ws.max_row+1
                ws.cell(row=new_row, column=1, value=rpw_dime[f"A{row.row}"].value)
                ws.cell(row=new_row, column=2, value=rpw.upper())
                if "NOR" in rpw:
                    ws.cell(row=new_row, column=3, value="Nortel")
                else:
                    ws.cell(row=new_row, column=3, value="Dimensional")
                ws.cell(row=new_row, column=4, value=rpw_dime[f"C{row.row}"].value)
                ws.cell(row=new_row, column=5, value=rpw_dime[f"D{row.row}"].value)
                ws.cell(row=new_row, column=6, value=rpw_dime[f"E{row.row}"].value)


wb.save("data/rpw.xlsx")
