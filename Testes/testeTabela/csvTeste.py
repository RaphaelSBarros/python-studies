import csv

with open("Testes/testeTabela/aviso.csv", 'r') as f:
   reader=csv.reader(f, delimiter=',')
   for linha in reader:
      print(linha)