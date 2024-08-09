import pandas as pd
import os


MES = 'Dezembro'

# Pasta contendo os arquivos Excel
pasta_dos_arquivos = '/Users/tatianeyukarikawakami/Desktop/Sumire/POR FABRICANTE/12 2023'

# Lista todos os arquivos na pasta
arquivos_excel = [arquivo for arquivo in os.listdir(pasta_dos_arquivos) if arquivo.endswith('.xlsx')]

# Lista para armazenar DataFrames de cada planilha
dataframes = []

# Loop para processar cada arquivo
for arquivo_excel in arquivos_excel:
    # Constrói o caminho completo para o arquivo
    caminho_do_arquivo = os.path.join(pasta_dos_arquivos, arquivo_excel)

    # Carrega a planilha Excel
    dados = pd.read_excel(caminho_do_arquivo)

    # Exclui a primeira linha
    dados = dados.iloc[1:]

    # Excluindo a última linha
    dados = dados.iloc[:-1]

    # Obtém o nome da planilha (sem extensão)
    nome_da_planilha = arquivo_excel.split('.')[0]

    # Adiciona uma nova coluna com o nome da planilha
    dados['Nome da Planilha'] = nome_da_planilha

    # Adiciona uma nova coluna com a informação MES
    dados['Mês'] = MES
    dados['Ano'] = 2023

    # Adiciona o DataFrame à lista
    dataframes.append(dados)

# Combina todos os DataFrames em um único DataFrame
dados_combinados = pd.concat(dataframes, ignore_index=True)

# Salva o DataFrame combinado em uma única planilha Excel
caminho_da_saida = '/Users/tatianeyukarikawakami/Desktop/Sumire/POR FABRICANTE/2023/planilha_combinada_Dezembro.xlsx'
dados_combinados.to_excel(caminho_da_saida, index=False)


