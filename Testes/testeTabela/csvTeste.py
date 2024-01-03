import csv

with open("C:/Users/DESOUR10/Downloads/_Aviso de Férias - Mensalistas - Layout_Python (2).csv", 'r') as f:
   reader=csv.reader(f)
   for linha in reader:
      print(linha)