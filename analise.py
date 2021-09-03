from numpy import array, true_divide
from numpy.lib.function_base import copy
import pandas as pd

####---------------------------------------------------------------   
  
df = pd.read_excel('tabela.xlsx' ,usecols=[ 'ANUNCIO' ,'RESULTADO' , 'PRECOS' , 'PMA'] )
qtd_linhas = df['PRECOS'].count()

numero = qtd_linhas
npma = df['PMA'].str.replace(',', '.' )
nprecos = df['PRECOS'].str.replace(',', '.' )

df['PMA'] = npma 
df['PRECOS'] = nprecos 
#df.dropna(subset=['PMA'])
df['PRECOS'] = df['PRECOS'].astype(float)
df['PMA'] = df['PMA'].astype(float)


#for linha in range (1 , numero):
valores = []
for line in df.itertuples():
 valores.append('Acima do PMA' if line.PMA >= line.PRECOS else 'Abaixo do PMA' )

 
 
    
df['RESULTADO'] = valores


df.to_excel('tabela.xlsx' , sheet_name= 'Teste1')


print(df.dtypes) 
print(df)