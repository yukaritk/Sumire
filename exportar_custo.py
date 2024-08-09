import pandas as pd
import os
import warnings

# Define os caminhos das pastas e arquivos
FOLDER = '/Users/tatianeyukarikawakami/Desktop/custo'

df_consol = pd.DataFrame()

# Lista todos os arquivos na pasta especificada
excel_files = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx') or arquivo.endswith('.xls')]

for arquivo_excel in excel_files:
    try:
        caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_excel(caminho_do_arquivo, engine='openpyxl', skiprows=4)

        try:
            # Renomeia as colunas
            df.columns = [col.split('-')[0].replace(' ', '') for col in df.columns]
        except:
            continue

        # Rename the column 'Produto' to 'Codigo'
        df.rename(columns={'Produto': 'Codigo'}, inplace=True)

        # Padroniza a coluna 'Codigo' para string e remove a parte decimal, se houver
        df['Codigo'] = df['Codigo'].astype(str).apply(lambda x: x.split('.')[0])

        # Concatena o DataFrame atual com o DataFrame consolidado
        df_consol = pd.concat([df_consol, df], ignore_index=True)
    except Exception as e:
        print(e)
        continue

# Salva o DataFrame consolidado em um arquivo Excel
df_consol.to_csv('/Users/tatianeyukarikawakami/Desktop/consol_custo.txt', sep='|', index=False)