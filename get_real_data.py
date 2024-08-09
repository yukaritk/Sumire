import pandas as pd
import os
import warnings

FOLDER = '/Users/tatianeyukarikawakami/Desktop/venda_real'

def get_real_data():
    # Iniciando um DataFrame vazio
    lista_completa = pd.DataFrame()

    # Lista todos os arquivos na pasta
    excel_files = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

    # Loop para processar cada arquivo
    for arquivo_excel in excel_files:
        # Constrói o caminho completo para o arquivo
        caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
        
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
                df = pd.read_excel(caminho_do_arquivo, skiprows=3)
                
            df['Total'] = df[[str(i) for i in range(1, 32)]].sum(axis=1)
            df.rename(columns={'Código': 'Codigo'}, inplace=True)
            # Lista das colunas a serem removidas
            cols_to_drop = ['Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Descrição', 'M / D'] + [str(i) for i in range(1, 32)]
            # Removendo as colunas
            df.drop(columns=cols_to_drop, inplace=True, errors='ignore')  # 'errors='ignore'' para ignorar erros se algumas colunas não existirem

            # Concatenando o DataFrame atual com a lista completa
            lista_completa = pd.concat([lista_completa, df], ignore_index=True)
        except Exception as e:
            print(f"Erro ao processar o arquivo {caminho_do_arquivo}: {e}")

    lista_completa.set_index(['Loja', 'Codigo'], inplace=True)
    
    return lista_completa

# # Lista todos os arquivos na pasta
# excel_files = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]
# list = []
# # Loop para processar cada arquivo
# for arquivo_excel in excel_files:
#     # Constrói o caminho completo para o arquivo
#     caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    
#     try:
#         with warnings.catch_warnings():
#             warnings.simplefilter("ignore")
#             # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
#             df = pd.read_excel(caminho_do_arquivo, skiprows=3)
#             list.append(df.iloc[0,1])
#     except:
#         pass

# data = pd.DataFrame(list)
# data.to_excel('/Users/tatianeyukarikawakami/Desktop/vendas_reais.xlsx')