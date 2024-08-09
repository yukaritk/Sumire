import pandas as pd
import os

FOLDER = '/Users/tatianeyukarikawakami/Desktop/Sumire/Acomp Vendas/'

# Lista todos os arquivos na pasta
excel = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

# Lista para armazenar DataFrames de cada planilha
df_acompVendas = []

# Loop para processar cada arquivo
for arquivo_excel in excel:
    # Constrói o caminho completo para o arquivo
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    
    # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
    dados = pd.read_excel(caminho_do_arquivo, skiprows=6)

    # Exclui a última linha do DataFrame 'dados'
    dados = dados.iloc[:-1]

    # Obtém o nome da planilha (sem extensão)
    nome_da_planilha = arquivo_excel.split('.')[0].split('-')[0].split("LOJA")[1]

    # Adiciona uma nova coluna com o nome da planilha
    dados['Loja'] = nome_da_planilha

    # Adiciona o DataFrame à lista
    df_acompVendas.append(dados)

# Combina todos os DataFrames em um único DataFrame
df_acompVendas_combinado = pd.concat(df_acompVendas, ignore_index=True)

# Salva o DataFrame combinado em uma única planilha Excel
caminho_da_saida = FOLDER + "AcompVendas_Total.xlsx"
df_acompVendas_combinado.to_excel(caminho_da_saida, index=False)

