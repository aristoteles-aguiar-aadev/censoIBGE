
from sklearn import metrics
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd
import win32com.client as win32
import numpy as np

# importar a base de dados IBGE
tabela_ibge = pd.read_excel('ibge.xlsx')
df = pd.read_excel('ibge.xlsx')

# visualizar a base de dados IBGE
pd.set_option('display.max.columns', None)
print(tabela_ibge)

print('-' * 50)
# Comparativo do primeiro com o ultimo ano IBGE
comparativo_First = tabela_ibge[[
    'Unid Fed', 'Ano1872']].groupby('Unid Fed').sum()
print(comparativo_First)

print('-' * 50)
# Comparativo em percentual do 1872 - 2010
comparativo_Last = tabela_ibge[[
    'Unid Fed', 'Ano2010']].groupby('Unid Fed').sum()
print(comparativo_Last)

comparativo_First_Last1 = tabela_ibge[[
    'Unid Fed', 'Ano1872', 'Ano2010']].groupby('Unid Fed').sum()
print(comparativo_First_Last1)

print('-' * 50)
# Comparativo em percentual do 1872 - 2010
comparativo_First_Last = (
    comparativo_First['Ano1872'] / comparativo_Last['Ano2010']).to_frame()
print(comparativo_First_Last)
print('-' * 50)
# Soma de TODOS os anos IBGE por Cidade
df['Total'] = df.sum(axis=1)
print(df)

print('-' * 50)
# Soma de TODOS os anos IBGE Geral
df.loc['Total'] = df.sum()
print(df)


plt.figure(figsize=(15, 10))
sns.lineplot(data=tabela_ibge)
plt.show()


# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'seu-email@gmail.com; seu-email@icloud.com'
mail.Subject = 'Relatório IBGE 1872-2010'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Dados do IBGE de 1872 a 2010.</p>

<p>Comparativo entre o ano de 1872 e o ano 2010 IBGE:</p>
{comparativo_First_Last1.to_html()}

<p>Comparativo de Todos os anos IBGE:</p>
{df.to_html()}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Aristoteles Aguiar</p>
<p>https://www.instagram.com/aristoteles_aguiar/</p>

<p>Fonte.: https://ftp.ibge.gov.br/Censos/Censo_Demografico_2010/Sinopse/Brasil/sinopse_brasil_tab_1_4.zip.</p>
'''

mail.Send()

print('Email Enviado.')
