# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 07:16:09 2024

@author: Muriel Gauer
"""

import pandas as pd

# Carregar os arquivos Excel em DataFrames
df_demanda_estoque = pd.read_excel('demanda_estoque.xlsx')
df_plano_producao = pd.read_excel('plano_producao.xlsx')

# Supondo que ambos os DataFrames têm uma coluna comum 'Produto' para fazer o merge
df_combined = pd.merge(df_demanda_estoque, df_plano_producao, on='Produto')

# Calcular a demanda mensal
df_combined['Demanda_Mensal'] = df_combined['Demanda_Anual'] / 12

# Calcular a produção mensal
df_combined['Producao_Mensal'] = df_combined['Plano_Producao'] / 12

# Calcular o estoque de segurança para seis meses
df_combined['Estoque_Seguranca'] = df_combined['Demanda_Mensal'] * 6

# Calcular o saldo de estoque após 12 meses
df_combined['Saldo_Estoque_12_Meses'] = (
    df_combined['Estoque_Atual'] +
    df_combined['Plano_Producao'] - 
    df_combined['Demanda_Anual'] - 
    df_combined['Estoque_Seguranca']
)

# Calcular a demanda mensal considerando o estoque de segurança como parte da demanda
df_combined['Demanda_Mensal_Com_Estoque'] = (df_combined['Demanda_Anual'] + df_combined['Estoque_Seguranca']) / 12

# Converter o saldo de estoque em duração em meses considerando a demanda mensal com estoque de segurança
df_combined['Duracao_Estoque_12_Meses'] = df_combined['Saldo_Estoque_12_Meses'] / df_combined['Demanda_Mensal_Com_Estoque']

# Verificar o DataFrame combinado com a duração do estoque após 12 meses
print("DataFrame combinado com duração do estoque após 12 meses:")
print(df_combined[['Produto', 'Demanda_Anual', 'Estoque_Seguranca', 'Estoque_Atual', 'Plano_Producao', 'Saldo_Estoque_12_Meses', 'Duracao_Estoque_12_Meses']])
 
