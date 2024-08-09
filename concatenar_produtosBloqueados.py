import pandas as pd
import os

def criar_df():
    pasta_dos_arquivos = '/Users/tatianeyukarikawakami/Desktop/Sumire/Produtos Bloqueados/'

    # Lista todos os arquivos na pasta
    arquivos_excel = [arquivo for arquivo in os.listdir(pasta_dos_arquivos) if arquivo.endswith('.xlsx')]

    # Lista para armazenar DataFrames de cada planilha
    df_prodBlock = []

    # Loop para processar cada arquivo
    for arquivo_excel in arquivos_excel:
        # Constrói o caminho completo para o arquivo
        caminho_do_arquivo = os.path.join(pasta_dos_arquivos, arquivo_excel)

        try:
            # Tenta carregar a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
            dados = pd.read_excel(caminho_do_arquivo, skiprows=3)
            print(dados)

            # Coluna loja
            loja = arquivo_excel.split('-')[1].split('.')[0]
            dados['Loja'] = loja

            # Coluna ID
            dados['ID'] = dados['Código'].astype(str) + "-" + dados['Loja'].astype(str)

            # Adiciona o DataFrame à lista
            df_prodBlock.append(dados)
        
        except Exception as e:
            # Imprime o erro e continua para o próximo arquivo
            print(f"Erro ao processar {arquivo_excel}: {e}")


    # Combina todos os DataFrames em um único DataFrame
    df_prodBlock_combinados = pd.concat(df_prodBlock, ignore_index=True)
    return df_prodBlock_combinados

