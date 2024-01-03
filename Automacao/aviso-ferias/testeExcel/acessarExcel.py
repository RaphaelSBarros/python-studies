import pandas as pd

def ler_planilha_excel(caminho_arquivo, nome_aba):
    try:
        # Carrega o arquivo Excel
        df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)

        # Cria listas para cada coluna
        lista_coluna_a = df['A'].tolist() if 'A' in df.columns else []
        lista_coluna_b = df['B'].tolist() if 'B' in df.columns else []
        lista_coluna_c = df['C'].tolist() if 'C' in df.columns else []
        lista_coluna_d = df['D'].tolist() if 'D' in df.columns else []
        lista_coluna_e = df['E'].tolist() if 'E' in df.columns else []
        lista_coluna_f = df['F'].tolist() if 'F' in df.columns else []

                # e os valores são listas contendo as linhas dessa coluna
        dados = {}
        for coluna in df.columns:
            dados[coluna] = df[coluna].tolist()

        return dados
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return None

# Exemplo de uso
caminho_arquivo_excel = 'C:\\Users\\CAMBIA3\\Desktop\\Automation\\Ferias\\EnviarAvisosDeFérias - Layout.xlsm'
nome_da_aba = 'Formulário'

dados = ler_planilha_excel(caminho_arquivo_excel, nome_da_aba)

if dados:
    for coluna, linhas in dados.items():
        print(f'Coluna {coluna}: {linhas}')
else:
    print('Não foi possível ler a planilha.')