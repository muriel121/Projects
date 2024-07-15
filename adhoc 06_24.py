# -*- coding: utf-8 -*-
"""
Created on Wed Jun 26 07:33:38 2024

@author: 81256535
"""

#TENTATIVA 0
import pandas as pd

# Carregar o arquivo Excel
adhoc = 'Copy 1 of Ad hoc Leaf all CU 20.06.2024 7.0.xlsx'
df = pd.read_excel(adhoc)

# Selecionar as colunas de interesse
selected_columns = [df.columns[0], df.columns[2], df.columns[3], df.columns[4]] + df.columns[5:].tolist()
df_selected = df[selected_columns].copy()

# Definir a coluna de interesse corretamente
column_of_interest = df_selected.columns[4]  # "Calendar Year/Month"

# Converter os valores da coluna de interesse para strings usando .loc para evitar o aviso
df_selected.loc[:, column_of_interest] = df_selected[column_of_interest].astype(str).str.strip().str.lower()

# Exibir valores únicos na coluna de interesse para inspeção
unique_values = df_selected[column_of_interest].unique()
print("Unique values:", unique_values)

# Verificar as condições com expressões regulares
mask = df_selected[column_of_interest].str.contains(r'\btotal demand\b', case=False, na=False) | \
       df_selected[column_of_interest].str.contains(r'estoque', case=False, na=False)

filtered_df = df_selected[mask]

# Eliminar a linha com o índice 0
filtered_df = filtered_df.drop(index=0)

# Verificar valores negativos
negative_values_mask = filtered_df.select_dtypes(include=['number']) < 0
filtered_df['has_negative_values'] = negative_values_mask.any(axis=1)

# Selecionar as colunas dos próximos 14 meses (ajustar conforme necessário)
month_columns = filtered_df.columns[5:19]

## Separar os dataframes de demanda e estoque
demand_df = filtered_df[filtered_df[column_of_interest] == 'total demand'].copy()
demand_df.info()
stock_df = filtered_df[filtered_df[column_of_interest] == 'estoque'].copy()

# Configurar o índice para manter todas as colunas de contexto
demand_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)
stock_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)

# Encontrar os índices comuns entre ambos os DataFrames
common_index = demand_df.index.intersection(stock_df.index)

# Reindexar ambos os DataFrames para garantir alinhamento
demand_df = demand_df.reindex(common_index)
stock_df = stock_df.reindex(common_index)

# Concatenar os DataFrames verticalmente
combined_df = pd.concat([demand_df, stock_df], keys=['total demand', 'estoque'], names=['Type'])

# Restaurar o índice
combined_df.reset_index(inplace=True)

# Comparar os valores entre demanda e estoque para cada produto
comparison_df = combined_df.pivot_table(index=['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month', 'Type'], values=month_columns).unstack('Type')
comparison_df.columns = [f'{col[1]}_{col[0]}' for col in comparison_df.columns]

# Adicionar uma coluna para informar o valor negativo e o mês em que isso ocorre
negative_values_info = []

for idx, row in comparison_df.iterrows():
    negative_info = "No negative values"
    for month in month_columns:
        demand_value = row.get(f'total demand_{month}', 0)
        stock_value = row.get(f'estoque_{month}', 0)
        if demand_value - stock_value < 0:
            negative_info = f"{demand_value - stock_value} in {month}"
            break
    negative_values_info.append(negative_info)

comparison_df['Negative Values'] = negative_values_info

# Exibir o DataFrame final com a comparação e a coluna adicional
print(comparison_df.head())

# Comparacao TT01_CU

demandTT_df = demand_df[demand_df['Location'] == 'TT01_CU']
estoqueTT_df = estoque_df[estoque_df['Location'] == 'TT01_CU']

# Configurar o índice para manter todas as colunas de contexto
demandTT_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)
estoqueTT_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)


# Comparar os valores entre demanda e estoque
comparisonTT_df = demandTT_df[month_columns].subtract(stockTT_df[month_columns], fill_value=0).reset_index()

# Adicionar uma coluna para informar o valor negativo e o mês em que isso ocorre
negative_values_infoTT = []

for idx, row in comparisonTT_df.iterrows():
    negative_info = "No negative values"
    for month in month_columns:
        if row[month] < 0:
            negative_info = f"{row[month]} in {month}"
            break
    negative_values_infoTT.append(negative_info)

comparisonTT_df['Negative Values'] = negative_values_infoTT

comparisonTT_file = 'comparisonTT_df.xlsx'
comparisonTT_df.to_excel(comparisonTT_file, index=False)


# Comparacao VE02_CU

demandVE_df = demand_df[demand_df['Location'] == 'VE02_CU']
stockVE_df = stock_df[stock_df['Location'] == 'VE02_CU']

# Configurar o índice para manter todas as colunas de contexto
demandVE_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)
stockVE_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)


# Comparar os valores entre demanda e estoque
comparisonVE_df = demandVE_df[month_columns].subtract(stockVE_df[month_columns], fill_value=0).reset_index()

# Adicionar uma coluna para informar o valor negativo e o mês em que isso ocorre
negative_values_infoVE = []

for idx, row in comparisonVE_df.iterrows():
    negative_info = "No negative values"
    for month in month_columns:
        if row[month] < 0:
            negative_info = f"{row[month]} in {month}"
            break
    negative_values_infoVE.append(negative_info)

comparisonVE_df['Negative Values'] = negative_values_infoVE

comparisonVE_file = 'comparisonVE_df.xlsx'
comparisonVE_df.to_excel(comparisonVE_file, index=False)















import pandas as pd

# Carregar o arquivo Excel
adhoc = 'Copy 1 of Ad hoc Leaf all CU 20.06.2024 7.0.xlsx'
df = pd.read_excel(adhoc)

# Selecionar as colunas de interesse
selected_columns = [df.columns[0], df.columns[2], df.columns[3], df.columns[4]] + df.columns[5:].tolist()
df_selected = df[selected_columns].copy()

# Definir a coluna de interesse corretamente
column_of_interest = df_selected.columns[4]  # "Calendar Year/Month"

# Converter os valores da coluna de interesse para strings usando .loc para evitar o aviso
df_selected.loc[:, column_of_interest] = df_selected[column_of_interest].astype(str).str.strip().str.lower()

# Exibir valores únicos na coluna de interesse para inspeção
unique_values = df_selected[column_of_interest].unique()
print("Unique values:", unique_values)

# Verificar as condições com expressões regulares
mask = df_selected[column_of_interest].str.contains(r'\btotal demand\b', case=False, na=False) | \
       df_selected[column_of_interest].str.contains(r'estoque', case=False, na=False)

filtered_df = df_selected[mask]

# Eliminar a linha com o índice 0
filtered_df = filtered_df.drop(index=0)

# Verificar valores negativos
negative_values_mask = filtered_df.select_dtypes(include=['number']) < 0
filtered_df['has_negative_values'] = negative_values_mask.any(axis=1)

# Selecionar as colunas dos próximos 14 meses (ajustar conforme necessário)
month_columns = filtered_df.columns[5:19]

# Separar os dataframes de demanda e estoque
demand_df = filtered_df[filtered_df[column_of_interest] == 'total demand'].copy()
stock_df = filtered_df[filtered_df[column_of_interest] == 'estoque'].copy()

# Configurar o índice para manter todas as colunas de contexto
demandTT_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)
stockTT_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)

# Combinar os dataframes de demanda e estoque
combined_df = pd.concat([demandTT_df, stockTT_df], keys=['total demand', 'estoque'], names=['Type'])

# Restaurar o índice
combined_df.reset_index(inplace=True)

# Comparar os valores entre demanda e estoque para cada produto
comparison_results = []

for idx, product in combined_df.groupby(['Location', 'Cluster Group', 'Trade Type', '% ITT']):
    demand = product[product['Type'] == 'total demand']
    stock = product[product['Type'] == 'estoque']
    
    if not demand.empty and not stock.empty:
        comparison = demand[month_columns].values - stock[month_columns].values
        comparison_df = pd.DataFrame(comparison, columns=month_columns)
        comparison_df.insert(0, 'Location', product['Location'].iloc[0])
        comparison_df.insert(1, 'Cluster Group', product['Cluster Group'].iloc[0])
        comparison_df.insert(2, 'Trade Type', product['Trade Type'].iloc[0])
        comparison_df.insert(3, '% ITT', product['% ITT'].iloc[0])
        
        negative_values_info = []
        
        for idx, row in comparison_df.iterrows():
            negative_info = "No negative values"
            for month in month_columns:
                if row[month] < 0:
                    negative_info = f"{row[month]} in {month}"
                    break
            negative_values_info.append(negative_info)
        
        comparison_df['Negative Values'] = negative_values_info
        comparison_results.append(comparison_df)

# Combinar todos os resultados de comparação em um único DataFrame
final_comparison_df = pd.concat(comparison_results)

# Exportar o DataFrame para um arquivo Excel
output_file = 'comparison_df.xlsx'
final_comparison_df.to_excel(output_file, index=False)

print(f"DataFrame exportado para {output_file}")



#TENTATIVA 2

import pandas as pd

# Carregar o arquivo Excel
adhoc = 'Copy 1 of Ad hoc Leaf all CU 20.06.2024 7.0.xlsx'
df = pd.read_excel(adhoc)

# Selecionar as colunas de interesse
selected_columns = [df.columns[0], df.columns[2], df.columns[3], df.columns[4]] + df.columns[5:].tolist()
df_selected = df[selected_columns].copy()

# Definir a coluna de interesse corretamente
column_of_interest = df_selected.columns[4]  # "Calendar Year/Month"

# Converter os valores da coluna de interesse para strings usando .loc para evitar o aviso
df_selected.loc[:, column_of_interest] = df_selected[column_of_interest].astype(str).str.strip().str.lower()

# Exibir valores únicos na coluna de interesse para inspeção
unique_values = df_selected[column_of_interest].unique()
print("Unique values:", unique_values)

# Verificar as condições com expressões regulares
mask = df_selected[column_of_interest].str.contains(r'\btotal demand\b', case=False, na=False) | \
       df_selected[column_of_interest].str.contains(r'estoque', case=False, na=False)

filtered_df = df_selected[mask]

# Eliminar a linha com o índice 0
filtered_df = filtered_df.drop(index=0)

# Selecionar as colunas dos próximos 14 meses (ajustar conforme necessário)
month_columns = filtered_df.columns[5:19]

# Separar os dataframes de demanda e estoque
demand_df = filtered_df[filtered_df[column_of_interest] == 'total demand'].copy()
stock_df = filtered_df[filtered_df[column_of_interest] == 'estoque'].copy()

# Configurar o índice para manter todas as colunas de contexto
demand_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)
stock_df.set_index(['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], inplace=True)

# Restaurar o índice antes de combinar os DataFrames
demand_df.reset_index(inplace=True)
stock_df.reset_index(inplace=True)

# Combinar os dataframes de demanda e estoque
combined_df = pd.merge(demandTT_df, stockTT_df, on=['Location', 'Cluster Group', 'Trade Type', '% ITT', 'Calendar Year/Month'], suffixes=('_demand', '_stock'))

# Comparar os valores entre demanda e estoque
comparison_df = combined_df.copy()
for month in month_columns:
    comparison_df[month] = combined_df[f'{month}_demand'] - combined_df[f'{month}_stock']

# Adicionar uma coluna para informar o valor negativo e o mês em que isso ocorre
negative_values_info = []

for idx, row in comparison_df.iterrows():
    negative_info = "No negative values"
    for month in month_columns:
        if row[month] < 0:
            negative_info = f"{row[month]} in {month}"
            break
    negative_values_info.append(negative_info)

comparison_df['Negative Values'] = negative_values_info

# Exportar o DataFrame para um arquivo Excel

comparison_df.to_excel



