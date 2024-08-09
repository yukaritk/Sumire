import pandas as pd
from extract_data import get_data
from pmdarima import auto_arima
from statsmodels.tsa.statespace.sarimax import SARIMAX
import time

def get_forecast():

    # Obter os dados
    df = pd.DataFrame(get_data())

    # # Excluindo 6 meses
    # df = df.drop(df.index[2:185])

    # Lista para armazenar os resultados de todas as lojas
    resultado_geral = []

    # # Lista para armazenar Best Model
    # best_model = []

    for loja in df.columns:
        resultado_loja = []  # Lista para armazenar os resultados por loja

        df_loja = pd.DataFrame(df[loja])

        # Converter índice para datetime, ignorando erros
        try:
            df_loja.index = pd.to_datetime(df_loja.index, format='%d/%m/%Y', errors='coerce')
            df_loja.index.freq = pd.infer_freq(df_loja.index)  # Inferir a frequência
        except Exception as e:
            print(f"Erro ao converter índice para datetime: {e}")
        
        loja = df_loja.iloc[0, 0]
        code = df_loja.iloc[1, 0]
        df_loja = df_loja[2:]

        # Garantir que todos os dados são numéricos e tratar valores NaN
        df_loja = df_loja.apply(pd.to_numeric, errors='coerce').dropna()

        # Remover outliers usando IQR
        Q1 = df_loja.quantile(0.25)
        Q3 = df_loja.quantile(0.75)
        IQR = Q3 - Q1
        df_loja = df_loja[~((df_loja < (Q1 - 1.5 * IQR)) | (df_loja > (Q3 + 1.5 * IQR)))]

        if len(df_loja) > 1:
            try:
                # # Encontrar os melhores parâmetros usando auto_arima
                # auto_model = auto_arima(df_loja, seasonal=True, m=12, stepwise=True, error_action='ignore', suppress_warnings=True)
                # order = auto_model.order
                # seasonal_order = auto_model.seasonal_order
                # best_model.append((order, seasonal_order))

                # Aplicar o modelo SARIMAX com os parâmetros definidos manualmente
                model = SARIMAX(df_loja, order=(0, 1, 1), seasonal_order=(0, 0, 0, 12), enforce_stationarity=False, enforce_invertibility=False)
                results = model.fit(disp=False)

                # Fazer previsão para os próximos 31 dias
                forecast = results.get_forecast(steps=30)
                forecast_media = forecast.predicted_mean
                forecast_conf_int = forecast.conf_int()
                forecast_conf_int[forecast_conf_int < 0] = 0  # Ajuste do intervalo de confiança para não ser negativo

                # Somar as previsões para os próximos 31 dias para estimar o consumo mensal
                consumo_mensal = forecast_media.head(30).sum()

                # Preparar o DataFrame de Resultado por loja
                resultado_loja.append(loja)
                resultado_loja.append(code)
                resultado_loja.append(int(consumo_mensal))
                resultado_loja.append(int(forecast_conf_int.iloc[:, 1].max()))  # 100% do valor máximo do intervalo de confiança superior

                # Incluir o resultado_loja na lista resultado_geral
                resultado_geral.append(resultado_loja)

            except Exception as e:
                print(f"Erro ao processar a série para {loja}: {e}")

    # Transformar Resultado Geral em um DataFrame
    colunas = ['Loja', 'Codigo', 'Consumo Mensal', 'Confiança Máxima']
    resultado_df = pd.DataFrame(resultado_geral, columns=colunas)
    resultado_df.set_index(['Loja', 'Codigo'], inplace=True)
    resultado_df['Total'] = resultado_df['Consumo Mensal'] + resultado_df['Confiança Máxima']
    
    return resultado_df

print(get_forecast())