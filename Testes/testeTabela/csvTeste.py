import csv

with open("C:/Users/DESOUR10/Downloads/_Aviso de Férias - Mensalistas - Layout_Python (2).csv", 'r') as f:
   reader=csv.reader(f)
   rows=list(reader)
   
for row in rows:
   row.append('new value')
   
with open('tabela.csv', 'w', newline='') as f:
   writer=csv.writer
   writer.writerows(rows)