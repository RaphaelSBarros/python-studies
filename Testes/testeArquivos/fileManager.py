import os

folder_path = "C://Users//DESOUR10//Downloads//07-16 - Copiar e Mover Arquivos//" #escolhe a pasta dos arquivos
arquivos = os.listdir(folder_path) #lista todos os arquivos dentro da pasta escolhida

for arquivo in arquivos: #para cada arquivo dentro da lista de arquivos
    if 'xlsx' in arquivo: #caso tenha xlsx no arquivo
        if 'Compras' in arquivo: # e caso tenha Comprar no nome
            source_path = os.path.join(folder_path, arquivo) #pega o arquivo que se encaixa
            destination_path = os.path.join(folder_path, "Compras", arquivo) #cria uma variável com o destino do arquivo C://Users//DESOUR10//Downloads//07-16 - Copiar e Mover Arquivos//Compras
            os.rename(source_path, destination_path) #coloca o arquivo dentro da pasta de destino
        elif "Vendas" in arquivo: #faz o mesmo para os arquivos com Vendas
            source_path = os.path.join(folder_path, arquivo)
            destination_path = os.path.join(folder_path, "Vendas", arquivo)
            os.rename(source_path, destination_path)