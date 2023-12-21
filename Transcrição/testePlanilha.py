import openpyxl

workbook = openpyxl.load_workbook('EnviarAvisosDeFérias - Layout.xlsm') #Abrir o arquivo de planilha

planilha = workbook['Formulário'] #abrir a aba desejada dentro da planilha

celula = planilha.cell(row=2, column=2) # escolhe a celula

valor = celula.value #armazena o valor em uma variável

celula.value = "Valor teste" #troca o valor da célula para Valor Teste

workbook.save('exemplo.xlsx') #salva as alterações em uma planilha