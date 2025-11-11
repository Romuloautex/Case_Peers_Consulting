import pandas as pd
from datetime import datetime

#Case Infomaz - Rômulo A. T.

# Carregar a base de dados
file_path = '../data/Case_Infomaz_Base_de_Dados.xlsx'

# Limpeza e preparação dos dados
df_produtos = pd.read_excel(file_path, sheet_name='Cadastro Produtos', skiprows=1)
df_clientes = pd.read_excel(file_path, sheet_name='Cadastro Clientes', skiprows=1)
df_estoque = pd.read_excel(file_path, sheet_name='Cadastro de Estoque', skiprows=1)
df_fornecedores = pd.read_excel(file_path, sheet_name='Cadastro Fornecedores', skiprows=1)
df_vendas = pd.read_excel(file_path, sheet_name='Transações Vendas', skiprows=1)

# skiprows=1 para ignorar a primeira linha das abas

df_produtos = df_produtos.dropna(how='all')
df_clientes = df_clientes.dropna(how='all')
df_estoque = df_estoque.dropna(how='all')
df_fornecedores = df_fornecedores.dropna(how='all')
df_vendas = df_vendas.dropna(how='all')

df_produtos.columns = ['ID_PRODUTO', 'ID_ESTOQUE', 'NOME_PRODUTO', 'CATEGORIA']
df_clientes.columns = ['ID_CLIENTE', 'NOME_CLIENTE', 'DATA_CADASTRO']
df_estoque.columns = ['ID_ESTOQUE', 'VALOR_ESTOQUE', 'QTD_ESTOQUE', 'DATA_ESTOQUE', 'ID_FORNECEDOR']
df_fornecedores.columns = ['ID_FORNECEDOR', 'NOME_FORNECEDOR', 'DATA_CADASTRO']
df_vendas.columns = ['ID_NOTA', 'DATA_NOTA', 'VALOR_NOTA', 'VALOR_ITEM', 'QTD_ITEM', 'ID_PRODUTO', 'ID_CLIENTE']

df_clientes['DATA_CADASTRO'] = pd.to_datetime(df_clientes['DATA_CADASTRO'])
df_estoque['DATA_ESTOQUE'] = pd.to_datetime(df_estoque['DATA_ESTOQUE'])
df_vendas['DATA_NOTA'] = pd.to_datetime(df_vendas['DATA_NOTA'])

df_vendas['MES_ANO'] = df_vendas['DATA_NOTA'].dt.to_period('M')

# 1) Calcule o valor total de venda dos produtos por categoria. Utilize as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.
vendas_categoria = pd.merge(df_vendas, df_produtos, on='ID_PRODUTO', how='left')

vendas_categoria['VALOR_TOTAL_LINHA'] = vendas_categoria['VALOR_ITEM'] * vendas_categoria['QTD_ITEM']

total_vendas_categoria = vendas_categoria.groupby('CATEGORIA')['VALOR_TOTAL_LINHA'].sum().reset_index()
total_vendas_categoria.columns = ['CATEGORIA', 'VALOR_TOTAL_VENDAS']

print("1. Valor total de venda por categoria:")
print(total_vendas_categoria.sort_values('VALOR_TOTAL_VENDAS', ascending=False))

# 2) Calcule a margem dos produtos subtraindo o valor unitario pelo valor de venda. Utilize as tabelas CADASTRO_ESTOQUE e TRANSACOES_VENDAS.
produtos_estoque = pd.merge(df_produtos, df_estoque, on='ID_ESTOQUE', how='left')

produtos_estoque['VALOR_UNITARIO'] = produtos_estoque['VALOR_ESTOQUE'] / produtos_estoque['QTD_ESTOQUE']

vendas_margem = pd.merge(df_vendas, produtos_estoque, on='ID_PRODUTO', how='left')
vendas_margem['MARGEM'] = vendas_margem['VALOR_ITEM'] - vendas_margem['VALOR_UNITARIO']

margem_produtos = vendas_margem.groupby(['ID_PRODUTO', 'NOME_PRODUTO'])['MARGEM'].mean().reset_index()

print("\n2. Margem média por produto:")
print(margem_produtos.sort_values('MARGEM', ascending=False))

# 3) Calcule um ranking de clientes por quantidade de produtos comprados por mês. Utilize as tabelas CADASTRO_CLIENTES e TRANSACOES_VENDAS.
clientes_qtd = df_vendas.groupby(['ID_CLIENTE', 'MES_ANO'])['QTD_ITEM'].sum().reset_index()

clientes_qtd = pd.merge(clientes_qtd, df_clientes, on='ID_CLIENTE', how='left')

ranking_clientes = clientes_qtd.sort_values(['MES_ANO', 'QTD_ITEM'], ascending=[True, False])

print("\n3. Ranking de clientes por quantidade comprada por mês:")
for mes, dados in ranking_clientes.groupby('MES_ANO'):
    print(f"\n--- Mês: {mes} ---")
    print(dados.drop(columns='MES_ANO'))

# 4) Calcule um ranking de fornecedores por quantidade de estoque disponivel por mês. Utilize as tabelas CADASTRO_FORNECEDORES e CADASTRO_ESTOQUE.
estoque_fornecedores = pd.merge(df_estoque, df_fornecedores, on='ID_FORNECEDOR', how='left')

estoque_fornecedores['MES_ANO'] = estoque_fornecedores['DATA_ESTOQUE'].dt.to_period('M')

ranking_fornecedores = estoque_fornecedores.groupby(['NOME_FORNECEDOR', 'MES_ANO'])['QTD_ESTOQUE'].sum().reset_index()

ranking_fornecedores = ranking_fornecedores.sort_values(['MES_ANO', 'QTD_ESTOQUE'], ascending=[True, False])

print("\n4. Ranking de fornecedores por quantidade em estoque por mês:")
for mes, dados in ranking_fornecedores.groupby('MES_ANO'):
    print(f"\n--- Mês: {mes} ---")
    print(dados.drop(columns='MES_ANO'))

# 5) Calcule um ranking de produtos por quantidade de venda por mês. Utilize a tabela TRANSACOES_VENDAS.
vendas_produto_mes = df_vendas.groupby(['ID_PRODUTO', 'MES_ANO'])['QTD_ITEM'].sum().reset_index()

vendas_produto_mes = pd.merge(vendas_produto_mes, df_produtos, on='ID_PRODUTO', how='left')

ranking_produtos_qtd = vendas_produto_mes.sort_values(['MES_ANO', 'QTD_ITEM'], ascending=[True, False])

print("\n5. Ranking de produtos por quantidade vendida por mês:")
for mes, dados in ranking_produtos_qtd.groupby('MES_ANO'):
    print(f"\n--- Mês: {mes} ---")
    print(dados.drop(columns='MES_ANO'))

# 6) Calcule um ranking de produtos por valor de venda por mês. Utilize a tabela TRANSACOES_VENDAS.
vendas_valor_mes = df_vendas.groupby(['ID_PRODUTO', 'MES_ANO'])['VALOR_ITEM'].sum().reset_index()

vendas_valor_mes = pd.merge(vendas_valor_mes, df_produtos, on='ID_PRODUTO', how='left')

ranking_produtos_valor = vendas_valor_mes.sort_values(['MES_ANO', 'VALOR_ITEM'], ascending=[True, False])

print("\n6. Ranking de produtos por valor de venda por mês:")
for mes, dados in ranking_produtos_valor.groupby('MES_ANO'):
    print(f"\n--- Mês: {mes} ---")
    print(dados.drop(columns='MES_ANO'))

# 7) Calcule a média de valor de venda por categoria de produto por mês. Utiliza as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.
vendas_categoria_mes = pd.merge(df_vendas, df_produtos, on='ID_PRODUTO', how='left')

media_categoria_mes = vendas_categoria_mes.groupby(['CATEGORIA', 'MES_ANO'])['VALOR_ITEM'].mean().reset_index()

print("\n7. Média de valor de venda por categoria por mês:")
#print(media_categoria_mes.sort_values(['MES_ANO', 'VALOR_ITEM'], ascending=[True, False]))
for mes, dados in media_categoria_mes.groupby('MES_ANO'):
    print(f"\n--- Mês: {mes} ---")
    print(dados.drop(columns='MES_ANO'))

# 8) Calcule um ranking de margem de lucro por categoria
margem_categoria = vendas_margem.groupby('CATEGORIA')['MARGEM'].mean().reset_index()

print("\n8. Ranking de margem de lucro por categoria:")
print(margem_categoria.sort_values('MARGEM', ascending=False))

# 9) Liste produtos comprados por clientes
compras_clientes = pd.merge(df_vendas, df_produtos, on='ID_PRODUTO', how='left')
compras_clientes = pd.merge(compras_clientes, df_clientes, on='ID_CLIENTE', how='left')

lista_compras = compras_clientes[['NOME_CLIENTE', 'NOME_PRODUTO', 'CATEGORIA', 'QTD_ITEM', 'VALOR_ITEM', 'DATA_NOTA']]

print("\n9. Lista de produtos comprados por clientes:")
print(lista_compras)

# 10) Ranking de produtos por quantidade de estoque
produtos_estoque_qtd = pd.merge(df_produtos, df_estoque, on='ID_ESTOQUE', how='left')

ranking_estoque = produtos_estoque_qtd.groupby(['ID_PRODUTO', 'NOME_PRODUTO'])['QTD_ESTOQUE'].sum().reset_index()

ranking_estoque = ranking_estoque.sort_values('QTD_ESTOQUE', ascending=False)

print("\n10. Ranking de produtos por quantidade em estoque:")
print(ranking_estoque)

# Criar um arquivo Excel com os resultados
nome_arquivo = f'Resultados_Infomaz_{datetime.now().strftime("%Y%m%d")}.xlsx'

with pd.ExcelWriter(nome_arquivo) as writer:
    total_vendas_categoria.to_excel(writer, sheet_name='Vendas por Categoria', index=False)
    margem_produtos.to_excel(writer, sheet_name='Margem por Produto', index=False)
    ranking_clientes.to_excel(writer, sheet_name='Ranking Clientes', index=False)
    ranking_fornecedores.to_excel(writer, sheet_name='Ranking Fornecedores', index=False)
    ranking_produtos_qtd.to_excel(writer, sheet_name='Ranking Produtos Qtd', index=False)
    ranking_produtos_valor.to_excel(writer, sheet_name='Ranking Produtos Valor', index=False)
    media_categoria_mes.to_excel(writer, sheet_name='Média Categoria', index=False)
    margem_categoria.to_excel(writer, sheet_name='Margem por Categoria', index=False)
    lista_compras.to_excel(writer, sheet_name='Compras por Cliente', index=False)
    ranking_estoque.to_excel(writer, sheet_name='Estoque por Produto', index=False)

print("\nResultados exportados para 'Resultados_Infomaz.xlsx'")

