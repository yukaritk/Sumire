import pandas as pd
import os
import math
from concatenar_produtosBloqueados import criar_df
import numpy as np

FOLDER = '/Users/tatianeyukarikawakami/Desktop/Sumire/Evol Vendas/'

# Lista todos os arquivos na pasta
excel = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

# Lista para armazenar DataFrames de cada planilha
df_evolVendas = []

# Importar planilha de Produtos Bloqueados
df_prodBlock = criar_df()

# Loop para processar cada arquivo
for arquivo_excel in excel:
    # Constrói o caminho completo para o arquivo
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    
    # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
    dados = pd.read_excel(caminho_do_arquivo, skiprows=5)

    # Obtém o nome da planilha (sem extensão)
    nome_da_planilha = arquivo_excel.split('.')[0]

    # Adiciona uma nova coluna com o nome da planilha
    dados['Fabricante'] = nome_da_planilha

    # Adicionar a nova coluna com o cálculo
    dados['VMD'] = ((dados['JAN'] + dados['DEZ'] + dados['NOV'] + dados['OUT'] + dados['SET']) / 152).apply(lambda x: math.ceil(x * 100) / 100)

    # Adicionar coluna Sugestao
    dados['Sugestao'] = (dados['VMD']*60).round(0)

    # Adicionar coluna dif
    dados['dif'] = dados['Sugestao'] - dados['Estoque Atual']

    # Adicionar coluna valor de estoque
    dados['Valor Estoque'] = (dados['Preço Custo'] * dados['Estoque Atual']).round(2)

    # Adicionar coluna valor sugestao de valor
    dados['Valor Sugestao'] = (dados['Preço Custo'] * dados['Sugestao'])

    # Adicionar a coluna valor dif
    dados['Valor dif'] = dados['dif'] * dados['Preço Custo']

    # Adicionar valor dif e Substituir valores negativos por 0
    dados['Valor Falta'] = (dados['dif'] * dados['Preço Custo']).where(lambda x: x > 0, 0)  

    # Adicionar valor dif e Substituir valores positivo por 0
    dados['Valor Excesso'] = (dados['dif'] * dados['Preço Custo']).where(lambda x: x < 0, 0)

    #Adicionar coluna ID
    dados['ID'] = dados['Código'].astype(str) +'-'+ dados['Loja'].astype(str)

    # Adicionar coluna de bloqueados
    dados['Obs'] = dados['ID'].isin(df_prodBlock['ID']).map({True: 'Bloqueado', False: ''})

    # Adicionar "Itens novos" quando o código for maior que 425070
    dados['Itens Novos'] = np.where(dados['Código'] > 425070, 'Itens novos', '')

    # Adicionar "Tirar" apenas onde a coluna 'Obs' estiver vazia e o sugestao for igual a zero
    dados['Obs'] = np.where((dados['Obs'] == '') & (dados['Sugestao'] == 0), 'Tirar', dados['Obs'])

    # Adicionar "Comprar" apenas onde a coluna 'Obs' estiver vazia e o sugestao for maior a zero
    dados['Obs'] = np.where((dados['Obs'] == '') & (dados['dif'] > 0), 'Comprar', dados['Obs'])

    # Adiciona o DataFrame à lista
    df_evolVendas.append(dados)

# Combina todos os DataFrames em um único DataFrame
df_evolVendas_combinado = pd.concat(df_evolVendas, ignore_index=True)
print(df_evolVendas_combinado)
# Salva o DataFrame combinado em uma única planilha Excel
caminho_da_saida = FOLDER + "EvolVendas_Total.xlsx"
df_evolVendas_combinado.to_excel(caminho_da_saida, index=False)

