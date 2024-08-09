import pandas as pd
import os


# Pasta contendo os arquivos Excel
pasta_dos_arquivos = '/Users/tatianeyukarikawakami/Desktop/Sumire/POR FABRICANTE/2023'

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

    # Renomeia as colunas conforme especificado
    dados = dados.rename(columns={
        'Unnamed: 0': 'Fabricante',
        'Unnamed: 1': 'Vl. Bruto',
        'Unnamed: 2': 'Vl. Liquido',
        'Unnamed: 3': 'Vl. Custo',
        'Unnamed: 4': 'MC',
        'Unnamed: 5': 'PM',
        'Unnamed: 6': 'Qtd',
        'Unnamed: 7': 'Verba Compra',
        'Unnamed: 8': 'Verba Venda',
        'Nome da Planilha': 'Loja'
    })

    # Adiciona o DataFrame à lista
    dataframes.append(dados)

# Combina todos os DataFrames em um único DataFrame
dados_combinados = pd.concat(dataframes, ignore_index=True)

# Salva o DataFrame combinado em uma única planilha Excel
caminho_da_saida = '/Users/tatianeyukarikawakami/Desktop/Sumire/POR FABRICANTE/2023/planilha_combinada_Geral.xlsx'
dados_combinados.to_excel(caminho_da_saida, index=False)

