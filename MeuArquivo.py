import pandas as pd # para análise de dados vindo da planilha
import win32com.client as win32 # para enviar email via outlook

# Importar a Base de Dados (a partir de um arquivo Excel.xlsx)
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados (usando pandas -> pd.read_excel())
# Todas as colunas (.set_option(opção, valor)) -> Configração do pandas
pd.set_option('display.max_columns', None)

# Printando tabela completa e apenas 2 colunas
print(tabela_vendas)
# print(tabela_vendas[['ID Loja', 'Valor Final']])

# Calcular faturamento por loja
tabela_faturamento_lojas = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(tabela_faturamento_lojas)
print('-' * 50)

# Calcular Quantidade de Produtos vendidos por loja
tabela_quantidades_produtosVendidos_lojas = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(tabela_quantidades_produtosVendidos_lojas)
print('-' * 50)

# Calcular o Ticket Médio (Faturamento / Quantidade de Produtos Vendidos pela Loja) em cada Loja
# .to_frame() é necessário para transformar em tabela, pois ao calcular informações de tabelas distintas, vira float
tabela_ticketMedio_lojas = (tabela_faturamento_lojas['Valor Final'] / tabela_quantidades_produtosVendidos_lojas['Quantidade']).to_frame()
# ATribuindo o título da coluna para Ticket Médio
tabela_ticketMedio_lojas = tabela_ticketMedio_lojas.rename(columns={0: 'Ticket Médio'})
print(tabela_ticketMedio_lojas)
print('-' * 50)

# Enviar email com relatório (pré requisito: Outlook e instalar biblioteca pywin32 - Integração Python e Windows)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'devmarcoribeiro@gmail.com'
mail.Subject = 'TESTE - Relatórios de Vendas por Loja (Automatizado)'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{tabela_faturamento_lojas.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de Itens Vendidos:</p>
{tabela_quantidades_produtosVendidos_lojas.to_html(formatters={'Quantidade': '{:,}'.format})}

<p>Ticket Médio dos Produtos de cada loja:</p>
{tabela_ticketMedio_lojas.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>

<p>Att.</p>
<p>Marco Ribeiro</p>
'''

mail.Send()