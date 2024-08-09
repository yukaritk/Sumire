import os
import pandas as pd

FOLDER = "/Users/tatianeyukarikawakami/Downloads/EvolVendaProduto"

# Lista todos os arquivos na pasta
excel = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

# Lista para armazenar códigos extraídos de cada arquivo junto com o nome do arquivo
codigos_extraidos = []

# Loop para processar cada arquivo
for arquivo_excel in excel:
    # Constrói o caminho completo para o arquivo
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)

    try:
        # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
        dados = pd.read_excel(caminho_do_arquivo, skiprows=3)

        # Adiciona os códigos extraídos e o nome do arquivo à lista de códigos
        codigos_extraidos.append((dados["Código"].iloc[0], caminho_do_arquivo))
    except Exception as e:
        codigos_extraidos.append(('Erro', caminho_do_arquivo))

# Cria o DataFrame com os códigos extraídos e os nomes dos arquivos
df_codigos_extraidos = pd.DataFrame(codigos_extraidos, columns=['Código', 'Nome do Arquivo'])

# Exporta a tabela resultante para um arquivo Excel
df_codigos_extraidos.to_excel('/Users/tatianeyukarikawakami/Desktop/teste/Resultado.xlsx', index=False)