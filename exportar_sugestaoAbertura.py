import pandas as pd
import os
import warnings


FOLDER = '/Users/tatianeyukarikawakami/Desktop/Sugestao'

# List all Excel files starting with 'SugAberturaLoja '
excel_files = [arquivo for arquivo in os.listdir(FOLDER)]
consolidada = pd.DataFrame()

for arquivo_excel in excel_files:
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            nota = pd.read_excel(caminho_do_arquivo, engine='openpyxl')

        loja = nota.iloc[1,2]
        fabricante = nota.iloc[0,2]


        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(caminho_do_arquivo, engine='openpyxl', skiprows=6)

        drop_col = ['Unnamed: 0', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 'MAI', 'ABR', 'MAR', 'FEV', 'AGO', 'JUN', 'JUL', 'DE', 'Média VMD','Unnamed: 17', 'Unnamed: 18']
        df.drop(columns=drop_col, inplace=True)
        
        df['Loja'] = loja.split('-')[0]
        df['Loja'] = df['Loja'].apply(lambda loja: loja.split('-')[0].replace(' ', ''))
        df['Fabricante'] = fabricante

        # Rename the column 'Código' to 'Codigo'
        df.rename(columns={'Código': 'Codigo'}, inplace=True)

        # Padroniza a coluna 'Codigo' para string e remove a parte decimal, se houver
        df['Codigo'] = df['Codigo'].astype(str).apply(lambda x: x.split('.')[0])

        # Adiciona uma coluna 'ID' ao df_ativo
        df['ID'] = df['Loja'] + '-' + df['Codigo']
        
        consolidada = pd.concat([consolidada, df])
    except Exception as e:
        print(f'Erro ao abrir arquivo {arquivo_excel} - {e}')
# Configurar pandas para mostrar todas as linhas de índice
pd.set_option('display.multi_sparse', False)
consolidada.to_csv('/Users/tatianeyukarikawakami/Desktop/consol_sugestao_abertura.txt', sep='|', index=False)
