import os
import csv

def export_filenames(directory):
    try:
        # Lista para armazenar os nomes dos arquivos
        filenames = []

        # Obter todos os arquivos na pasta especificada
        for file_name in os.listdir(directory):
            # Verificar se é um arquivo e não um diretório
            if os.path.isfile(os.path.join(directory, file_name)):
                filenames.append(file_name)

        # Caminho para o arquivo CSV de saída
        output_file = os.path.join(directory, 'file_list.csv')

        # Exportar a lista de nomes de arquivos para um arquivo CSV
        with open(output_file, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['File Name'])  # Cabeçalho do CSV
            for name in filenames:
                writer.writerow([name])

        print(f"Lista de nomes de arquivos exportada para: {output_file}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

# Caminho para a pasta
directory_path = '/Users/tatianeyukarikawakami/Desktop/venda_real'

# Chamada da função
export_filenames(directory_path)