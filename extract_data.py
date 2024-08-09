import pandas as pd
import os
import warnings
import datetime


FOLDER = "/Users/tatianeyukarikawakami/Desktop/EvolVendaProduto"
GRUPOS = pd.read_excel('/Users/tatianeyukarikawakami/Desktop/Sumire/Grupos.xlsx')


def datas():
    data_inicial = datetime.date(2023, 6, 1)
    data_final = datetime.date(2024, 5, 31)
    return [data_inicial + datetime.timedelta(days=x) for x in range((data_final - data_inicial).days + 1)]

def get_data():
    # Iniciar um DataFrame vazio.
    df_compilado = pd.DataFrame()

    # Lista todos os arquivos na pasta
    excel = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

    # Loop para processar cada arquivo
    for arquivo_excel in excel:
        # Constrói o caminho completo para o arquivo
        caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
        
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
                dados = pd.read_excel(caminho_do_arquivo, skiprows=3)
                
            # Excluir colunas específicas
            colunas_para_excluir = ['Unnamed: 2','Unnamed: 3','Descrição','Unnamed: 5','Unnamed: 6']
            dados = dados.drop(columns=colunas_para_excluir)

            # Extrair codigo
            codigo_extraido = dados["Código"].iloc[0]

            # Lista para compilar info
            compilado = []

            # Loop para processar cada loja
            for loja in dados['Loja'].unique():
                data = []
                data.append(loja)
                data.append(codigo_extraido)
                perfil = GRUPOS.loc[GRUPOS['Indice']== int(loja), 'Perfil de consumo'].values[0]
                data.append(perfil)
                coluna = datas()

                # Loop para processar cada dia
                for col in coluna:
                    d = col.day
                    m = col.month

                    # Extrair info
                    try:
                        info = dados.loc[(dados['Loja'] == int(loja)) & (dados['M / D'] == int(m)), str(d)].values[0]
                        data.append(info)
                    except:
                        data.append('NaN')
                # Compilar info        
                compilado.append(data)
        except:
            pass

        df_comp = pd.DataFrame(compilado)

        # Concatecar as listas de compilados
        df_compilado = pd.concat([df_compilado, df_comp], ignore_index=True)
    
    # Inserir nome das colunas
    df_compilado.columns = ["Loja","Codigo","Perfil"] + [date.strftime('%d/%m/%Y') for date in coluna]
    # df_compilado.T.to_excel('/Users/tatianeyukarikawakami/Desktop/extract_data.xlsx')

    return df_compilado.T

