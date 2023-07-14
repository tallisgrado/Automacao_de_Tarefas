#AUTOMAÇÃO DE PROCESSOS:

''' 
    Nesse projeto, automatizei a criação de um relatório de fechamento de mercado 
enviado por email
    Passo 1: Importar módulos e bibliotecas.
    Passo 2: Pegar dados do Ibovespa e do Dólar no Yahoo Finance.
    Passo 3: Manipular os dados para deixa-los nos formatos necessários 
    para fazer contas.
    Passo 4: Calcular o retorno diário, mensal e anual.
    Passo 5: Localizar, dentro da tabela de retornos, os valores de fechamento de 
    irão pro texto anexado no Email.
    Passo 6: Fazer o gráfico dos ativos.
    Passo 7: Enviar Email.
'''

# Passo 1:
import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
from email.message import EmailMessage

# Passo 2:
ativos = ('^BVSP', 'BRL=X')
hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)
dados_mercado = yf.download(ativos, um_ano_atras, hoje)

# Passo 3:
dados_fechamento = dados_mercado['Adj Close']
dados_fechamento.columns = ['Dólar','Ibovespa']
dados_fechamento = dados_fechamento.dropna()

dados_fechamento_mensal = dados_fechamento.resample('M').last()
dados_fechamento_anual = dados_fechamento.resample('Y').last()

# Passo 4:
retorno_no_ano = dados_fechamento_anual.pct_change().dropna()
retorno_no_mes = dados_fechamento_mensal.pct_change().dropna()
retorno_no_dia = dados_fechamento.pct_change().dropna()

# Passo 5:
retorno_dia_dolar = retorno_no_dia.iloc[-1,0]
retorno_dia_ibov = retorno_no_dia.iloc[-1,1]
retorno_mes_dolar = retorno_no_mes.iloc[-1,0]
retorno_mes_ibov = retorno_no_mes.iloc[-1,1]
retorno_ano_dolar = retorno_no_ano.iloc[-1,0]
retorno_ano_ibov = retorno_no_ano.iloc[-1,1]

retorno_dia_dolar = round(retorno_dia_dolar*100, 2)
retorno_dia_ibov = round(retorno_dia_ibov * 100, 2)
retorno_mes_dolar = round(retorno_mes_dolar * 100, 2)
retorno_mes_ibov = round(retorno_mes_ibov * 100, 2)
retorno_ano_dolar = round(retorno_ano_dolar * 100, 2)
retorno_ano_ibov = round(retorno_ano_ibov * 100, 2)

# Passo 6:
plt.style.use('cyberpunk')
dados_fechamento.plot( y= 'Ibovespa', use_index= True, legend= False)
plt.title('Ibovespa')
plt.savefig('ibovespa.png', dpi = 300)

plt.style.use('cyberpunk')
dados_fechamento.plot( y= 'Dólar', use_index= True, legend= False)
plt.title('Dolar')
plt.savefig('dolar.png', dpi = 300)

# Passo 7:
import os
from dotenv import load_dotenv
load_dotenv()
senha = os.environ.get('senha')
email = 'tallescunha2016@gmail.com'
msg = EmailMessage()
msg['subject'] = 'Relatório de Fechamento de Mercado'
msg['from'] = 'tallescunha2016@gmail.com'
msg['to'] = 'tallescunha2016@gmail.com' , 'edificareconstrutora@hotmail.com'
msg.set_content(f'''Prezado diretor, segue o relatório diário.

Bolsa:
                
No ano o Ibovespa está tendo uma rentabilidade de {retorno_ano_ibov}%,
enquanto no mês a rentabilidade é de {retorno_mes_ibov}%.

No ultimo dia útil, o fechamento do Ibovespa foi de {retorno_dia_ibov}%.

Dólar:
                
No ano o Dólar está tendo uma rentabilidade de {retorno_ano_dolar}%,
enquanto no mês a rentabilidade é de {retorno_mes_dolar}%.

No ultimo dia útil, o fechamento do Dólar foi de {retorno_dia_dolar}%.

Grande abraço,

T.M.C


''')

with open('dolar.png', 'rb') as content_file:
    content = content_file.read()
    msg.add_attachment(content, maintype = 'application', subtype = 'png', filename = 'dolar.png')

with open('ibovespa.png', 'rb') as content_file:
    content = content_file.read()
    msg.add_attachment(content, maintype = 'application', subtype = 'png', filename = 'ibovespa.png')

import smtplib
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(email, senha)
    smtp.send_message(msg)
