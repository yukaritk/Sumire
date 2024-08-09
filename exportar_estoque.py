import pandas as pd
import os
import warnings

# Define os caminhos das pastas e arquivos
FOLDER = '/Users/tatianeyukarikawakami/Desktop/Estoque'
grupo = '/Users/tatianeyukarikawakami/Desktop/Sumire/Grupos.xlsx'

# Lista todos os arquivos na pasta especificada
excel_files = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx') or arquivo.endswith('.xls')]

# Carrega o DataFrame de grupos uma vez fora do loop
df_grupo = pd.read_excel(grupo)

# DataFrame geral para consolidar os dados
df_geral = pd.DataFrame()

for arquivo_excel in excel_files:
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Lê o arquivo Excel para obter o CNPJ
            nota = pd.read_excel(caminho_do_arquivo, engine='openpyxl')
        
        # Extrai o CNPJ da célula específica
        cnpj = nota.iloc[2, 5]

        # Extrai o fabricante
        fabricante = nota.iloc[3, 8]
        
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Lê o arquivo Excel ignorando as primeiras 7 linhas
            df = pd.read_excel(caminho_do_arquivo, engine='openpyxl', skiprows=7)
        
        # Remove colunas indesejadas
        colunms_drop = [
            'Unnamed: 0', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6',
            'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13',
            'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18', 'Unnamed: 19',
            'Unnamed: 20', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23', 'Estoque', 'Unnamed: 25', 'Unnamed: 26'
        ]
        df.drop(columns=colunms_drop, inplace=True, errors='ignore')
        
        # Procurar a loja correspondente ao CNPJ no df_grupo
        loja = df_grupo.loc[df_grupo['CNPJ'] == cnpj, 'Loja'].values

        # Incluir coluna Fabricante
        df['Fabricante'] = fabricante
        
        # Incluir coluna Loja
        df['Loja'] = loja[0].split('-')[0]
        df['Loja'] = df['Loja'].apply(lambda loja: loja.split('-')[0].replace(' ', ''))
        df['Loja'] = df['Loja'].astype(str)

        # Renomear a coluna 'Cód. Produto' para 'Codigo'
        df.rename(columns={'Cód. Produto': 'Codigo'}, inplace=True)

        # Renomear a coluna Estoque
        df.rename(columns={'Unnamed: 24': 'Estoque'}, inplace=True)

        # Padroniza a coluna 'Codigo' para string e remove a parte decimal, se houver
        df['Codigo'] = df['Codigo'].astype(str).apply(lambda x: x.split('.')[0])

        # Adiciona a coluna 'ID'
        df['ID'] = df['Loja'] + '-' + df['Codigo']
        
        # Adiciona ao DataFrame geral
        df_geral = pd.concat([df_geral, df], ignore_index=True)
    except Exception as e:
        print(e)
        continue

# Salva o DataFrame consolidado em um arquivo de texto
df_geral.to_csv('/Users/tatianeyukarikawakami/Desktop/consol_estoque.txt', sep='|', index=False)