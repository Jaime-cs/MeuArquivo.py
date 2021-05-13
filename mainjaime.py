import pandas as pd
# importar no Terminal o arquivo pandas e o  arquivo pip install openpyxl

import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#Faturamento da loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de proodutosvendidospor loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('_' * 50)
# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar um e-mail com o relatório
# instalar pip install pywin32


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'cardozodossantos68@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody =  '''
<p>Prezados,</p>

<p>Segue o relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format()})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format()})})}

<p>Qualquer dúvida estamso a disposição.</p>

<p>Att. ,</p>

<p>Jaime</p>
'''

mail.Send()

print('E-mail enviado')