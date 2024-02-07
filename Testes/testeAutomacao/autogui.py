import pandas as pd

file_name = 'C:/Users/DESOUR10/Documents/929358.xlsx' # Escolhe o diretório do arquivo
df = pd.read_excel(file_name) # Lê o arquivo

columns = ['Data Acesso', 'Hora Acesso'] # Escolhe as colunas que deseja ler
#print(df)

dataI = 'Z' # Data retirada do RONDA
dataI = f'{dataI[4:]}-{dataI[2:4]}-{dataI[:2]}' # Traduz a data para que o formato seja aceito pelo excel

print(dataI) 

data = df[columns].where(df['Data Acesso'] == dataI) # Traz os dados dos acessos feitos na data escolhida

data = data.dropna() # Remove as linhas que possuirem o valor NaN

if len(data.index) < 1: # Caso o programa não encontre registros na data selecionada:
    print("Não achou") # Ele apenas retorna um erro e se encerra
else: # Caso encontre os registros ele faz o tratamento
    dataInicio = data.index[0]
    dataInicio = data.loc[dataInicio]['Hora Acesso']
    dataInicio = str(dataInicio)
    dataInicio = dataInicio[0:5]
    
    fim = len(data.index) - 1
    dataFim = data.index[fim]
    dataFim = data.loc[dataFim]['Hora Acesso']
    dataFim = str(dataFim)
    dataFim = dataFim[0:5]

    print(f'Hora de Entrada: {dataInicio}') # Devolve a hora do primeiro registro em catraca do dia
    print(f'Hora de Saída: {dataFim}') # Devolve a hora do último registro em catraca do dia
