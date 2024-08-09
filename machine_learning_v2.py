import pandas as pd
import os
import warnings
import datetime
from pmdarima import auto_arima
from statsmodels.tsa.statespace.sarimax import SARIMAX
import numpy as np

FOLDER = "/Users/tatianeyukarikawakami/Desktop/EvolVendaProduto"
VENDA_REAL_FOLDER = "/Users/tatianeyukarikawakami/Desktop/venda_real"
GRUPOS = pd.read_excel('/Users/tatianeyukarikawakami/Desktop/Sumire/Grupos.xlsx')

def datas():
    data_inicial = datetime.date(2023, 6, 1)
    data_final = datetime.date(2024, 5, 31)
    return [data_inicial + datetime.timedelta(days=x) for x in range((data_final - data_inicial).days + 1)]

# Lista todos os arquivos na pasta
excel = [arquivo for arquivo in os.listdir(FOLDER) if arquivo.endswith('.xlsx')]

# Lista para o resultado Geral
resultado_geral = []

# Loop para processar cada arquivo
for arquivo_excel in excel:
    # Constrói o caminho completo para o arquivo
    caminho_do_arquivo = os.path.join(FOLDER, arquivo_excel)
    
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Carrega a planilha Excel, pulando as quatro primeiras linhas e especificando o engine
            dados = pd.read_excel(caminho_do_arquivo, skiprows=3)
    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo_excel}: {e}")
        continue
           
    # Excluir colunas específicas
    colunas_para_excluir = ['Unnamed: 2','Unnamed: 3','Descrição','Unnamed: 5','Unnamed: 6']
    dados = dados.drop(columns=colunas_para_excluir)

    # Extrair codigo
    code = dados["Código"].iloc[0]

    # Loop para processar cada loja
    for loja in dados['Loja'].unique():
        # Lista para compilar info
        data = []
        compilado = []
        resultado_loja = []

        # Loop para processar cada dia
        for dia in datas():
            d = dia.day
            m = dia.month

            # Extrair info
            try:
                info = dados.loc[(dados['Loja'] == int(loja)) & (dados['M / D'] == int(m)), str(d)].values[0]
                data.append(info)
            except:
                data.append(np.nan)

        # Compilar info        
        compilado.append(data)

        df = pd.DataFrame(compilado)

        # Inserir nome das colunas
        df.columns = [date.strftime('%d/%m/%Y') for date in datas()]
        df = df.T

        # Converter índice para datetime, ignorando erros
        df.index = pd.to_datetime(df.index, format='%d/%m/%Y', errors='coerce')
        df.index.freq = 'D'  # Definir a frequência como diária

        # # Remover outliers usando IQR
        # Q1 = df.quantile(0.25)
        # Q3 = df.quantile(0.75)
        # IQR = Q3 - Q1
        # df = df[~((df < (Q1 - 1.5 * IQR)) | (df > (Q3 + 1.5 * IQR)))]

        # Garantir que todos os dados são numéricos e tratar valores NaN
        df = df.apply(pd.to_numeric, errors='coerce').dropna()

        # Encontrar os melhores parâmetros usando auto_arima
        auto_model = auto_arima(df, seasonal=True, m=12, stepwise=True, error_action='ignore', suppress_warnings=True)
        order = auto_model.order
        seasonal_order = auto_model.seasonal_order
        
        # Imprimir os valores de order e seasonal_order
        print(f"Loja: {loja} | Order: {order} | Seasonal Order: {seasonal_order}")

        # Aplicar o modelo SARIMAX com parâmetros ajustados
        try:
            model = SARIMAX(df, 
                            order=order, 
                            seasonal_order=seasonal_order, 
                            enforce_stationarity=False, 
                            enforce_invertibility=False)
            results = model.fit(disp=False)
        except Exception as e:
            print(f"Erro ao ajustar o modelo SARIMAX para a loja {loja}: {e}")
            continue

        # Fazer previsão para os próximos 30 dias
        try:
            forecast = results.get_forecast(steps=30)
            forecast_media = forecast.predicted_mean
            forecast_conf_int = forecast.conf_int()
            forecast_conf_int[forecast_conf_int < 0] = 0  # Ajuste do intervalo de confiança para não ser negativo

            # Somar as previsões para os próximos 30 dias para estimar o consumo mensal
            consumo_mensal = forecast_media.head(30).sum()
        except Exception as e:
            print(f"Erro ao fazer previsão para a loja {loja}: {e}")
            consumo_mensal = 0
            forecast_conf_int = pd.DataFrame([[0, 0]], columns=['lower', 'upper'])

        # Preparar o DataFrame de Resultado por loja
        max_conf = forecast_conf_int.iloc[:, 1].max()
        if pd.isna(max_conf):
            max_conf = 0

        resultado_loja.append({
            'Loja': loja,
            'Codigo': code,
            'Consumo Mensal': int(consumo_mensal),
            'Confiança Máxima': int(max_conf)
        })

        # Concatenar os resultado_loja
        resultado_geral.extend(resultado_loja)
     
    # Transformar Resultado Geral em um DataFrame
    df_result = pd.DataFrame(resultado_geral)
    df_result['Total'] = df_result['Consumo Mensal'] + df_result['Confiança Máxima']

    # Definir Loja e Código como índice
    df_result.set_index(['Loja', 'Codigo'], inplace=True)

    # Processar arquivos na pasta venda_real
    caminho_venda_real = os.path.join(VENDA_REAL_FOLDER, arquivo_excel)
    
    if os.path.exists(caminho_venda_real):
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                # Carregar a planilha, pular 3 linhas e setar Loja e Código como índice
                venda_real = pd.read_excel(caminho_venda_real, skiprows=3)
                venda_real.rename(columns={'Código': 'Codigo'}, inplace=True)
                venda_real.set_index(['Loja', 'Código'], inplace=True)
                
                # Somar as colunas de 1 a 31
                colunas_dias = [str(dia) for dia in range(1, 32)]
                venda_real['Venda Real'] = venda_real[colunas_dias].sum(axis=1)
                
                # Adicionar a coluna Venda Real ao df_result
                df_result = df_result.join(venda_real['Venda Real'], how='left')
        except Exception as e:
            print(f"Erro ao processar o arquivo de venda real {arquivo_excel}: {e}")
            continue

df_result.to_excel('/Users/tatianeyukarikawakami/Desktop/df_result.xlsx')
print(df_result)

