# Case Peers - Análise de Dados

Projeto desenvolvido para a dinâmica da Peers Consulting, analisando vendas, clientes e estoques com Python.

**Objetivo**: Identificar produtos/categorias mais rentáveis e sugerir estratégias de estoque e precificação.

## 🔍 Métricas Principais
1. Calcule o valor total de venda dos produtos por categoria. Utilize as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.
2. Calcule a margem dos produtos subtraindo o valor unitario pelo valor de venda. Utilize as tabelas CADASTRO_ESTOQUE e TRANSACOES_VENDAS.
3. Calcule um ranking de clientes por quantidade de produtos comprados por mês. Utilize as tabelas CADASTRO_CLIENTES e TRANSACOES_VENDAS.
4. Calcule um ranking de fornecedores por quantidade de estoque disponivel por mês. Utilize as tabelas CADASTRO_FORNECEDORES e CADASTRO_ESTOQUE.
5. Calcule um ranking de produtos por quantidade de venda por mês. Utilize a tabela TRANSACOES_VENDAS.
6. Calcule um ranking de produtos por valor de venda por mês. Utilize a tabela TRANSACOES_VENDAS.
7. Calcule a média de valor de venda por categoria de produto por mês. Utiliza as tabelas CADASTRO_PRODUTOS e TRANSACOES_VENDAS.
8. Calcule um ranking de margem de lucro por categoria
9. Liste produtos comprados por clientes
10. Ranking de produtos por quantidade de estoque

## 🛠️ Tecnologias Usadas
- Python (Pandas)
- Jupyter Notebook

## Observações
- **data/**: Contém `Case_Infomaz_Base_de_Dados.xlsx` (não incluído por confidencialidade).

## 🚀 Melhorias Futuras
- [ ] Calcular métricas complementares
- [ ] Integrar API de preços concorrentes
- [ ] Dashboard Power BI/Streamlit