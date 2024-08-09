# 1) Exportar planilha de Sugestao de Abertura
# 2) Exportar planilha de Estoque
# 3) Exportar planilha de Custo

import pandas as pd

def search_value(id_loja, id_code, df_custo):
    try:
        linha = df_custo[df_custo['Codigo'] == id_code]
 
        if id_loja not in df_custo.columns:
            return None  # Loja não encontrada, retornar None

        valor_custo = linha[id_loja].values[0]
        # Verificar se valor_custo é uma string e tentar converter para float
        if isinstance(valor_custo, str):
            valor_custo = valor_custo.replace(',', '.')  # Substituir vírgula por ponto se necessário
            valor_custo = float(valor_custo)
        return valor_custo
    
    except Exception as e:
        return None  # Em caso de erro, retornar None

# Caminhos dos arquivos
SUGESTAO = '/Users/tatianeyukarikawakami/Desktop/consol_sugestao_abertura.txt'
ESTOQUE = '/Users/tatianeyukarikawakami/Desktop/consol_estoque.txt'
CUSTO = '/Users/tatianeyukarikawakami/Desktop/consol_custo.txt'

# Especificar tipos de dados esperados
dtype_dict = {
    0: 'str',     # Coluna 0 como string
    205: 'str', # Coluna 205 como float (ou outro tipo esperado)
    213: 'str',   # Coluna 213 como string
    214: 'str'    # Coluna 214 como string
}

# Lendo os arquivos Excel
df_sugestao = pd.read_csv(SUGESTAO, sep='|')
df_estoque = pd.read_csv(ESTOQUE, sep='|')
df_custo = pd.read_csv(CUSTO, sep='|', dtype=dtype_dict)


# Remover duplicados em df_sugestao
df_sugestao = df_sugestao.drop_duplicates(subset='ID')

# Adicionar coluna 'Status' com a palavra 'Ativo' se o ID estiver em df_sugestao
df_estoque['Status'] = df_estoque['ID'].isin(df_sugestao['ID']).apply(lambda x: 'Ativo' if x else 'Inativo')

# Adicionar coluna 'Sugestao'
df_estoque['Sugestao'] = df_estoque['ID'].map(df_sugestao.set_index('ID')['Qtde Sugerida']).fillna(0).astype(float)

# Converter coluna 'Estoque' para float
df_estoque['Estoque'] = df_estoque['Estoque'].astype(float)

# Aplicar a função ao DataFrame para cada linha
df_estoque['Custo'] = df_estoque.apply(lambda row: search_value(id_loja=row['ID'].split('-')[0], id_code=row['ID'].split('-')[1], df_custo=df_custo), axis=1)

# Incluir coluna Dif
df_estoque['dif'] = df_estoque['Sugestao'] - df_estoque['Estoque']

# Incluir coluna Custo Estoque Total
df_estoque['Estoque Total'] = df_estoque['Custo'] * df_estoque['Estoque']

# Incluir colunas Compra e Excesso
df_estoque['Compra'] = df_estoque.apply(lambda row: row['dif'] * row['Custo'] if row['dif'] > 0 else 0, axis=1)
df_estoque['Excesso'] = df_estoque.apply(lambda row: row['dif'] * row['Custo'] if row['dif'] < 0 else 0, axis=1)

# Exportar para Excel
df_loja14 = df_estoque[(df_estoque['Loja'] == 'LOJA14') & (df_estoque['Excesso'] < 0)]
df_loja14.to_excel('/Users/tatianeyukarikawakami/Desktop/analise_excesso_loja14.xlsx', index=False)

df_compra = df_estoque[df_estoque['Compra'] > 0]
df_compra.to_excel('/Users/tatianeyukarikawakami/Desktop/analise_excesso_compra.xlsx', index=False)
