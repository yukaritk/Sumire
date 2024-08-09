import pandas as pd

df = pd.read_csv('/Users/tatianeyukarikawakami/Desktop/consol_estoque.txt', sep='|')
unique = df['Codigo'].unique()
df_u = pd.DataFrame(unique)

df_u.to_excel('/Users/tatianeyukarikawakami/Desktop/consol_estoque.xlsx', index=False)