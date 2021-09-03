# -------------------------------------------------------------------------------- BIBLIOTECAS

from numpy import array, string_, triu_indices, true_divide
from numpy.lib.function_base import copy
import requests
from bs4 import BeautifulSoup
import pandas as pd
import win32com.client as win32


####--------------------------------------------------------------------------------------- DATABASE
  
df = pd.read_excel('livro1.xlsx' ,usecols=[ 'CODIGO' , 'RESULTADO' , 'PRECOS' , 'ANUNCIO' , 'PMA'] )
#pd.set_option('display.max_rows', None)
#pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)
####----------------------------------------------------- LOOP DO LINK 
qtd_linhas = df['ANUNCIO'].count()
numero = qtd_linhas

for linha in range (1 , 15):
 

 links = df['ANUNCIO'] [linha] 

####--------------------------------------------------------------- WEBSCRAPING v
    
 try: 

     url = links
     page = requests.get(url)
    
     bs = BeautifulSoup(page.text, 'html.parser')

     try:
          preco = bs.find_all('li' , class_= 'avista_price_product')
          preco= preco[0].get_text().strip()
          num_preco = preco[5:10]
          pma = df['PMA'] [df.index == linha] = [ num_preco ]  

          print(num_preco)
         
####-------------------------------------------------------------------- ERRO DE VALOR  v    

     except: 
         pma = df['RESULTADO'] [df.index == linha] =  'Erro de Valor'
         print(pma)
                                                                                
####-------------------------------------------------------------------- ERRO DE LINK v

 except : 
     linkin = df['RESULTADO'] [df.index == linha] =  'Link Inválido'
     print(linkin)

####--------------------------------------------------------------- AJUSTE DE PREÇO v 

npma = df['PMA'].str.replace(',', '.' )
nprecos = df['PRECOS'].str.replace(',', '.' )

df['PMA'] = npma 
df['PRECOS'] = nprecos 
df['PRECOS'] = df['PRECOS'].astype(float)
df['PMA'] = df['PMA'].astype(float)

####--------------------------------------------------------------- CULCULO DE PMA v

valores = []
for line in df.itertuples():
 valores.append('Acima do PMA' if line.PMA >= line.PRECOS else 'Abaixo do PMA' )

 
df['RESULTADO'] = valores

####--------------------------------------------------------------- REMOVENDO ESPAÇOS EM "" v

df = df.dropna(subset=['PMA'])

####--------------------------------------------------------------- PUXANDO STRING v

ts = df[df['RESULTADO'] == 'Abaixo do PMA']


print(ts)

####--------------------------------------------------------------- ENVIANDO E-MAIL v


df['RESULTADO'] = df['RESULTADO'].astype(object)

try:
 outlook = win32.Dispatch('outlook.application')

 email = outlook.CreateItem(0)

 email.To = "jose.mellojr19@gmail.com"
 email.Subject = "Opa parece que Avistei um PMA Baixo"
 email.HTMLBody = f''' 
 <p>Olá  sou o Sentinela, Avistei PMA baixo vamos Resolver ? </p>

 <p>Essas São as Informações Do Nosso Produto : </p>

 <p>{ts}</p>

 <p>Abs,   </p>
 <p>Sentinela </p>
 '''

 email.send()

 print("Email enviado")



except :
 print(df)