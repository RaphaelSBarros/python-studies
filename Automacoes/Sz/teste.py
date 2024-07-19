from module import *

new_wb = Automatization()
new_wb.create_worksheet("teste2", "teste3", "teste4")
titles =["Período",
         "ID Atividade",
         "Projeto",
         "Tipo",
         "Status",
         "Descrição",
         "Atendente",
         "Data de Fechamento"
        ]
new_wb.entitle_worksheet(titles=titles, sheetname="teste3")
new_wb.fill_worksheet("teste3", 2, datetime.datetime(2024,7,19), 1, "BeeStock", "Bugs", "Aberto", "Algo escrito", "Raphael", datetime.datetime(2024,7,19))
new_wb.save_workbook("nome_teste")