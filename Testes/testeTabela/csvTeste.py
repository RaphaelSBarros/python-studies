import csv
import pandas as pd

#with open("C:/Users/DESOUR10/Downloads/_Aviso de Férias - Mensalistas - Layout_Python (2).csv", "r") as arquivo:
 #   arquivo_csv = csv.reader(arquivo, delimiter=",")
  #  for i, linha in enumerate(arquivo_csv):
   #     print(linha)
        
tabela = pd.read_csv("C:/Users/DESOUR10/Downloads/_Aviso de Férias - Mensalistas - Layout_Python (2).csv")
print(tabela)
print(tabela["ID"].sum())