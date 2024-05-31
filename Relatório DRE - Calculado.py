#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().system('pip install pandas-profiling==3.3.0 # Instalar pandas-profiling para perfilagem de dados')
get_ipython().system('pip install xlsxwriter # Instalar xlsxwriter para escrever arquivos Excel')
get_ipython().system('pip install pandas==2.0.3')


# In[2]:


import pandas as pd
import numpy as np
import babel.dates # fornece funcionalidades para lidar com datas e horários em aplicações Python


# # Preparar a base

# In[3]:


df_us_all = pd.read_excel(r'C:\Users\Kkk\3D Objects\Projeto\BASE_US.xlsx', sheet_name="Base")
df_br_all = pd.read_excel(r'C:\Users\Kkk\3D Objects\Projeto\BASE_BR.xlsx', sheet_name="Base")
df_pc_all = pd.read_excel(r'C:\Users\Kkk\3D Objects\Projeto\BASE_US.xlsx', sheet_name="Base")


# In[4]:


#Planilha que preencheremos o plano de contas
df_pc = df_pc_all[['Conta', 'Categoria']]
df_pc = df_pc.rename(columns={'Conta': 'Categoria_Num'})
df_pc['Total'] = 0


# In[5]:


#remover espaços nos nomes das colunas
df_us_all.columns = df_us_all.columns.str.strip()
df_br_all.columns = df_br_all.columns.str.strip()


# In[6]:


# criar variável para identificar as bases 
df_us_all = df_us_all.assign(Origem='US')
df_br_all = df_br_all.assign(Origem='BR')


# In[7]:


#definir somente as variáveis relevantes para o relatório
df_us = df_us_all[['Origem','Data de Competência', 'Categoria','Descrição', 'Cliente / Fornecedor','Valor recebido', 'Valor pago','Conta','Centro de custo', 'Empresa']]
df_br = df_br_all[['Origem','Data de Competência', 'Categoria','Descrição', 'Cliente / Fornecedor','Valor recebido', 'Valor pago','Conta','Centro de custo', 'Empresa']]


# In[8]:


#excluir a primeira linha da base / cabeçalho, para concatenarmos as bases (provenientes de diferentes fontes)
df_br = df_br.iloc[1:]


# In[9]:


df = pd.concat([df_us,df_br])


# In[ ]:


profile = ProfileReport(df)


# In[ ]:


profile


# In[10]:


#Engenharia de recursos,
# criar a coluna 'Total' com o valor recebido positivo e pago negativo
# o tipo entrada /saída
df['Total'] = df['Valor recebido'] - df['Valor pago']
df['Tipo'] = df['Total'].apply(lambda x: 'Entrada' if x > 0 else 'Saída')


# In[11]:


# mes e ano de ref.
# e remover espaços
df['Mês'] = df['Data de Competência'].dt.strftime('%B')
df['Ano'] = df['Data de Competência'].dt.year


# In[12]:


# visualizar as o andamento das alterações no excel, otimiza tempo
#df.to_csv('Base.csv', sep=';', index = False, encoding = 'utf-8-sig')


# In[13]:


# criando chaves extrangeiras onde ausentes
df.loc[df['Categoria'] == 'Transferência de Entrada', 'Categoria'] = '01 - ' + df.loc[df['Categoria'] == 'Transferência de Entrada', 'Categoria']
df.loc[df['Categoria'] == 'Transferência de Saída', 'Categoria'] = '02 - ' + df.loc[df['Categoria'] == 'Transferência de Saída', 'Categoria']
df.loc[df['Categoria'] == 'Saldo Inicial', 'Categoria'] = '00 - ' + df.loc[df['Categoria'] == 'Saldo Inicial', 'Categoria']
df.loc[df['Categoria'] == 'Descontos', 'Categoria'] = '99 - ' + df.loc[df['Categoria'] == 'Descontos', 'Categoria']


# In[14]:


#separar categóricas de númericas e excluir coluna desnecessária 
df[['Categoria_Num', 'Categoria']] = df['Categoria'].str.split('-', n=1, expand=True)
df[['Conta_Num', 'Conta']] = df['Conta'].str.split('-', n=1, expand=True)
df['Categoria_Num'] = df['Categoria_Num'].str.replace('.', '')


# In[15]:


df = df [['Origem','Data de Competência','Mês','Ano', 'Categoria_Num','Categoria','Descrição','Cliente / Fornecedor', 'Valor recebido', 'Valor pago','Total','Centro de custo','Conta', 'Empresa']]


# In[16]:


#removendo os espaços
#df['Categoria_Num'] = df['Categoria_Num'].str.strip()
#df['Conta'] = df['Conta'].str.strip()


# In[17]:


# tirar os pontos da chave extrangeira
#df['Categoria_Num'] = df['Categoria_Num'].astype(str)
#df_pc['Categoria_Num'] = df_pc['Categoria_Num'].str.replace('.', '')


# In[18]:


#dropando os nulos
df.dropna(how='all', inplace=True)


# In[19]:


df.isnull().sum()


# In[20]:


#preencher os valores nulos categoricos
df_cat = ['Centro de custo', 'Descrição', 'Categoria', 'Cliente / Fornecedor', 'Conta', 'Empresa']
for i in df_cat:
    df[i].fillna('NI', inplace=True)


# In[21]:


#preencher os valores nulos numéricos
df_num = ['Categoria_Num', 'Total']
for i in df_num:
    df[i].fillna(0, inplace=True)


# In[22]:


# padronização de formato
for i in df_cat:
    df[i] = df[i].str.title()


# In[23]:


#eita coisa linda de ver
df.info()


# In[24]:


#ajustar o tipo  (não vou alterar os valores monetarios em float devido a inconsistência na base, foi sugestionado a empresa o ajuste  )
df['Categoria_Num'] = df['Categoria_Num'].astype(str)
df['Data de Competência'] = pd.to_datetime(df['Data de Competência']).dt.date


# # Calculando o DRE
# 

# In[25]:


###   (+)
# Calcula a receita bruta agrupando por categoria e número de categoria
# Seleciona linhas onde a categoria começa com '01' ou é igual a '140103'
# Agrupa os resultados por categoria e número de categoria, soma o valor recebido e redefine os nomes das colunas
# Arredonda o valor total para 2 casas decimais
RECEITA_BRUTA = df[(df['Categoria_Num'].str.startswith('01')) | (df['Categoria_Num'] == '140103')].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
RECEITA_BRUTA.columns = ['Categoria_Num', 'Categoria', 'Total']
RECEITA_BRUTA['Total'] = RECEITA_BRUTA['Total'].round(2)


# In[26]:


###   (-)
DEDUCAO_IMPOSTOS = df[df['Categoria_Num'].str.startswith('02')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
DEDUCAO_IMPOSTOS.columns = ['Categoria_Num', 'Categoria', 'Total']
DEDUCAO_IMPOSTOS['Total'] = -DEDUCAO_IMPOSTOS['Total'].round(2)


# In[27]:


###   (-)
CUSTOS_SERV_VENDIDO = df[df['Categoria_Num'].str.startswith('03')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
CUSTOS_SERV_VENDIDO.columns = ['Categoria_Num', 'Categoria', 'Total']
CUSTOS_SERV_VENDIDO['Total'] = -CUSTOS_SERV_VENDIDO['Total'].round(2)


# In[28]:


###   (-)
DESPESAS_OPERACIONAIS = df[df['Categoria_Num'].str.startswith('04')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
DESPESAS_OPERACIONAIS.columns = ['Categoria_Num', 'Categoria', 'Total']
DESPESAS_OPERACIONAIS['Total'] = -DESPESAS_OPERACIONAIS['Total'].round(2)


# In[29]:


###   (-)
DEPRECIACOES_AMORTIZACOES = df[df['Categoria_Num'].str.startswith('05')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
DEPRECIACOES_AMORTIZACOES.columns = ['Categoria_Num', 'Categoria', 'Total']
DEPRECIACOES_AMORTIZACOES['Total'] = -DEPRECIACOES_AMORTIZACOES['Total'].round(2)


# In[30]:


###   (=)
RESULTADO_FINANCEIRO = df[(df['Categoria_Num'].str.startswith('0601')) | (df['Categoria_Num'].str.startswith('0602'))].copy()

RESULTADO_FINANCEIRO_0601 = RESULTADO_FINANCEIRO[RESULTADO_FINANCEIRO['Categoria_Num'].str.startswith('0601')].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
RESULTADO_FINANCEIRO_0601.columns = ['Categoria_Num', 'Categoria', 'Total']
RESULTADO_FINANCEIRO_0601['Total'] = RESULTADO_FINANCEIRO_0601['Total'].round(2)

RESULTADO_FINANCEIRO_0602 = RESULTADO_FINANCEIRO[RESULTADO_FINANCEIRO['Categoria_Num'].str.startswith('0602')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
RESULTADO_FINANCEIRO_0602.columns = ['Categoria_Num', 'Categoria', 'Total']
RESULTADO_FINANCEIRO_0602['Total'] = RESULTADO_FINANCEIRO_0602['Total'].round(2)

RESULTADO_FINANCEIRO = RESULTADO_FINANCEIRO_0601 + RESULTADO_FINANCEIRO_0602


# In[31]:


###   (=)
RESULTADO_N_OPERACIONAL = df[(df['Categoria_Num'].str.startswith('0701')) | (df['Categoria_Num'].str.startswith('0702'))].copy()

RESULTADO_N_OPERACIONAL_0701 = RESULTADO_N_OPERACIONAL[RESULTADO_N_OPERACIONAL['Categoria_Num'].str.startswith('0701')].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
RESULTADO_N_OPERACIONAL_0701.columns = ['Categoria_Num', 'Categoria', 'Total']
RESULTADO_N_OPERACIONAL_0701['Total'] = RESULTADO_N_OPERACIONAL_0701['Total'].round(2)

RESULTADO_N_OPERACIONAL_0702 = RESULTADO_N_OPERACIONAL[RESULTADO_N_OPERACIONAL['Categoria_Num'].str.startswith('0702')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
RESULTADO_N_OPERACIONAL_0702.columns = ['Categoria_Num', 'Categoria', 'Total']
RESULTADO_N_OPERACIONAL_0702['Total'] = RESULTADO_N_OPERACIONAL_0702['Total'].round(2)

RESULTADO_N_OPERACIONAL = pd.concat([RESULTADO_N_OPERACIONAL_0701, RESULTADO_N_OPERACIONAL_0702])


# In[32]:


###   (-)
IR_CSLL = df[df['Categoria_Num'].str.startswith('08')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
IR_CSLL.columns = ['Categoria_Num', 'Categoria', 'Total']
IR_CSLL['Total'] = IR_CSLL['Total'].round(2)


# In[33]:


###   (-)
DISTRIBUICAO_LUCROS = df[df['Categoria_Num'].str.startswith('09')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
DISTRIBUICAO_LUCROS.columns = ['Categoria_Num', 'Categoria', 'Total']
DISTRIBUICAO_LUCROS['Total'] = DISTRIBUICAO_LUCROS['Total'].round(2)


# In[34]:


###   (-)
INVESTIMENTO_IMOBILIZADO_INTANGIVEL = df[df['Categoria_Num'].str.startswith('10')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
INVESTIMENTO_IMOBILIZADO_INTANGIVEL.columns = ['Categoria_Num', 'Categoria', 'Total']
INVESTIMENTO_IMOBILIZADO_INTANGIVEL['Total'] = INVESTIMENTO_IMOBILIZADO_INTANGIVEL['Total'].round(2)


# In[35]:


###   (+)
ENTRADA_EMPRESTIMO = df[(df['Categoria_Num'].str.startswith('11'))].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
ENTRADA_EMPRESTIMO.columns = ['Categoria_Num', 'Categoria', 'Total']
ENTRADA_EMPRESTIMO['Total'] = ENTRADA_EMPRESTIMO['Total'].round(2)


# In[36]:


###   (-)
SAIDA_EMPRESTIMOS = df[df['Categoria_Num'].str.startswith('12')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
SAIDA_EMPRESTIMOS.columns = ['Categoria_Num', 'Categoria', 'Total']
SAIDA_EMPRESTIMOS['Total'] = SAIDA_EMPRESTIMOS['Total'].round(2)


# In[37]:


###   (+)
APORTES_CAPITAL = df[(df['Categoria_Num'].str.startswith('13'))].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
APORTES_CAPITAL.columns = ['Categoria_Num', 'Categoria', 'Total']
APORTES_CAPITAL['Total'] = APORTES_CAPITAL['Total'].round(2)


# In[38]:


OUTRAS_MOV_CAIXA = df[(df['Categoria_Num'].str.startswith('1401')) | (df['Categoria_Num'].str.startswith('1402'))].copy()

OUTRAS_MOV_CAIXA_1401 = OUTRAS_MOV_CAIXA[(OUTRAS_MOV_CAIXA['Categoria_Num'].str.startswith('1401')) & (OUTRAS_MOV_CAIXA['Categoria_Num'] != '140103')].groupby(['Categoria_Num', 'Categoria'])['Valor recebido'].sum().reset_index()
OUTRAS_MOV_CAIXA_1401.columns = ['Categoria_Num', 'Categoria', 'Total']
OUTRAS_MOV_CAIXA_1401['Total'] = OUTRAS_MOV_CAIXA_1401['Total'].round(2)

OUTRAS_MOV_CAIXA_1402 = OUTRAS_MOV_CAIXA[OUTRAS_MOV_CAIXA['Categoria_Num'].str.startswith('1402')].groupby(['Categoria_Num', 'Categoria'])['Valor pago'].sum().reset_index()
OUTRAS_MOV_CAIXA_1402.columns = ['Categoria_Num', 'Categoria', 'Total']
OUTRAS_MOV_CAIXA_1402['Total'] = OUTRAS_MOV_CAIXA_1402['Total'].round(2)

OUTRAS_MOV_CAIXA = pd.concat([OUTRAS_MOV_CAIXA_1401, OUTRAS_MOV_CAIXA_1402]).dropna()


# In[39]:


#RECEITA_BRASIL = df[df['Categoria_Num'].str.startswith('0101')].groupby(['Categoria_Num','Categoria'])['Valor recebido'].sum().reset_index()
#RECEITA_BRASIL.columns = ['Categoria_Num', 'Categoria', 'Total']
#RECEITA_BRASIL['Total'] = RECEITA_BRASIL['Total'].round(2)


# In[40]:


#RECEITA_BRASIL1 = RECEITA_BRASIL.groupby(['Categoria_Num', 'Categoria'])['Total'].sum().reset_index()


# In[41]:


RECEITA_BRUTA


# In[42]:


# Concatena as planilhas de contas para criar o plano de contas consolidado
# Ignora os índices para evitar duplicação de linhas
PLANODECONTAS1 = pd.concat([RECEITA_BRUTA, DEDUCAO_IMPOSTOS, CUSTOS_SERV_VENDIDO, DESPESAS_OPERACIONAIS, DEPRECIACOES_AMORTIZACOES, RESULTADO_FINANCEIRO, RESULTADO_N_OPERACIONAL, IR_CSLL, DISTRIBUICAO_LUCROS, INVESTIMENTO_IMOBILIZADO_INTANGIVEL, ENTRADA_EMPRESTIMO, SAIDA_EMPRESTIMOS, APORTES_CAPITAL, OUTRAS_MOV_CAIXA], ignore_index=True)


# In[43]:


# Agrupa o plano de contas consolidado por categoria número e soma os valores totais
# Reajusta o índice após a agrupação
# Substitui o código '01' por '01 ' para padronizar a formatação
df_pc1 = PLANODECONTAS1.groupby('Categoria_Num')['Total'].sum().reset_index()
df_pc1['Categoria_Num'] = df_pc1['Categoria_Num'].replace('01', '01 ')


# In[44]:


# Mapeia os valores totais do plano de contas consolidado para cada categoria número
# Substitui valores faltantes com 0 para evitar erros de cálculo
df_pc['Total'] = df_pc['Categoria_Num'].map(df_pc1.set_index('Categoria_Num')['Total']).fillna(0)


# In[45]:


# dropando nulos da base PC (plano de contas)
df_pc1.dropna(how='all', inplace=True)


# In[46]:


with pd.ExcelWriter('C:\\Users\\jpk\\OneDrive\\Documentos\\DNC\\Projetos\\Base.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Basededados', index=False)
    df_pc1.to_excel(writer, sheet_name='Plano de contas', index=False)


# In[47]:


#from google.colab import drive
#drive.mount('/content/drive')

