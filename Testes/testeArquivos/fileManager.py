import os

folder_path = "C://Users//DESOUR10//Downloads//07-16 - Copiar e Mover Arquivos//"
arquivos = os.listdir(folder_path)

for arquivo in arquivos:
    if 'xlsx' in arquivo:
        if 'Compras' in arquivo:
            source_path = os.path.join(folder_path, arquivo)
            destination_path = os.path.join(folder_path, "Compras", arquivo)
            os.rename(source_path, destination_path)
        elif "Vendas" in arquivo:
            source_path = os.path.join(folder_path, arquivo)
            destination_path = os.path.join(folder_path, "Vendas", arquivo)
            os.rename(source_path, destination_path)