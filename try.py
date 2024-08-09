import pandas as pd
import os
import warnings
import shutil

origem_folder = '/Users/tatianeyukarikawakami/Desktop/Sugestao'
destiny_folder = '/Users/tatianeyukarikawakami/Desktop/Sugestao'

# List all Excel files starting with 'SugAberturaLoja '
excel_files = [arquivo for arquivo in os.listdir(origem_folder)]

def get_unique_filename(directory, filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base}_{counter}{extension}"
        counter += 1
    return new_filename

for arquivo_excel in excel_files:
    caminho_do_arquivo = os.path.join(origem_folder, arquivo_excel)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df_ini = pd.read_excel(caminho_do_arquivo, engine='openpyxl')

            # Extracting the store name
            fabricante = df_ini.iloc[0, 2]
            loja = df_ini.iloc[1, 2]

            # Novo nome do arquivo
            novo_nome_arquivo = f"{loja} - {fabricante}.xlsx"
            novo_nome_arquivo_unico = get_unique_filename(destiny_folder, novo_nome_arquivo)
            novo_caminho_do_arquivo = os.path.join(destiny_folder, novo_nome_arquivo_unico)

            # Transferir arquivo para pasta destino e alterar para novo nome.
            shutil.move(caminho_do_arquivo, novo_caminho_do_arquivo)

    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo_excel}: {e}")
        continue