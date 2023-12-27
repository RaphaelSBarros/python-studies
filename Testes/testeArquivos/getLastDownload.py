import os
import glob

lista_arquivos= glob.glob("C:/Users/DESOUR10/Downloads/*") #pegando todos os arquivos dentro da pasta downloads
ultimo_arquivo=max(lista_arquivos, key=os.path.getmtime) #armazenando quem tem a data mais recente dentro da pasta
print(ultimo_arquivo) #exibindo o resultado