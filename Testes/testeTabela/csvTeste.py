import csv
import pandas as pd

tabela = pd.read_csv("C:/Users/DESOUR10/Downloads/_Aviso de Férias - Mensalistas - Layout_Python (2).csv")
print(tabela)

with open('C:/Users/DESOUR10/Desktop/Desenvolvimento/Testes/testeTabela/aviso.csv') as csv_file:
   csv_reader = csv.DictReader(csv_file)
   
   with open('new_aviso.csv', "w") as new_file:
      fieldnames = ['RE Colaborador','ID','Nome Destinatario','Email Destinatario', 'Status']
      csv_writer = csv.DictWriter(new_file, fieldnames=fieldnames, delimiter=',')
      csv_writer.writeheader()

      for line in csv_reader:
         csv_writer.writerow(line)