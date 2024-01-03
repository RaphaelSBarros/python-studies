import PyPDF2
import os

merger = PyPDF2.PdfMerger()

lista_arquivos = os.listdir("C:/Users/DESOUR10/Desktop/Desenvolvimento/Testes/testePDF/arquivos")
lista_arquivos.sort()

for arquivo in lista_arquivos:
    if ".pdf" in arquivo:
        merger.append(f"C:/Users/DESOUR10/Desktop/Desenvolvimento/Testes/testePDF/arquivos/{arquivo}")

merger.write("PDF final.pdf")