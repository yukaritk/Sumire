import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import numpy as np
from machine_learning import get_forecast  # Certifique-se de que a importação está correta
from get_real_data import get_real_data  # Certifique-se de que a importação está correta
from exportar_sugestaoAbertura import get_suggestion

# Importando os dataframes
forecast = get_forecast()
real = get_real_data()
sugestion = get_suggestion()


# Iniciando listas para comparação
valores_reais = []
valores_sugestao = []
atende = 0
count = 0

# Loop for para incluir valores nas listas
for n in forecast.index:
    try:
        real_valor = real.loc[n, 'Total']
        if isinstance(real_valor, (int, float, np.number)):
            valores_reais.append(real_valor)
        else:
            print(f"Valor inesperado em 'Total': {real_valor} para índice {n}")
            valores_reais.append(0)
    except:
        valores_reais.append(0)

    try:
        sugestao_valor = sugestion.loc[n, 'Qtde Sugerida']
        if isinstance(sugestao_valor, (int, float, np.number)):
            valores_sugestao.append(sugestao_valor)
        else:
            valores_sugestao.append(0)
    except:
        valores_sugestao.append(0)

    # Contagem de acertos
    try:
        if sugestao_valor >= real_valor & real_valor != 0:
            atende += 1
    except:
        pass
    count += 1

# Verificar o conteúdo das listas antes de converter para arrays NumPy
print("Valores Sugeridos:", valores_sugestao)
print("Valores Reais:", valores_reais)

# Convertendo listas em arrays NumPy para cálculos de métricas
valores_previstos = np.array(valores_sugestao).astype(float)
valores_reais = np.array(valores_reais).astype(float)

# Calculando métricas
mae = mean_absolute_error(valores_reais, valores_previstos)
rmse = np.sqrt(mean_squared_error(valores_reais, valores_previstos))
r2 = r2_score(valores_reais, valores_previstos)

# Corrigindo o cálculo do MAPE para evitar divisão por zero
mape = np.mean(np.abs((valores_reais - valores_previstos) / np.where(valores_reais != 0, valores_reais, 1))) * 100

print(f"Erro Médio Absoluto (MAE): {mae}")
print(f"Raiz do Erro Quadrático Médio (RMSE): {rmse}")
print(f"Coeficiente de Determinação (R²): {r2}")
print(f"Mean Absolute Percentage Error (MAPE): {mape:.2f}%")
print(f"Porcentagem de Acertos: {atende/count*100:.2f}%")