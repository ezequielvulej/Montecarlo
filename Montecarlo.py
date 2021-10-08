# -*- coding: utf-8 -*-
"""
Created on Sat Apr  4 10:55:00 2020

@author: evule
"""

import pandas as pd
import numpy as np
import math
import xlwings



Libro=pd.ExcelFile('Montecarlo.xlsx')
Inputs=pd.read_excel(Libro,'Inputs')
#t=10
#n=10
#FechaInicio='2019-12-29'
#FechaCalculo='2019-12-31'
print(Inputs)
t=Inputs.iloc[0,1]
n=Inputs.iloc[1,1]
FechaInicio=Inputs.iloc[2,1]
FechaCalculo=Inputs.iloc[3,1]
NivelConfianza=Inputs.iloc[4,1]


Factores=pd.read_excel('Montecarlo.xlsx', sheet_name='Factores')
Factores.reset_index
Factores['indice']=Factores.index  

iinicial=Factores[Factores.Fecha==FechaInicio].index.values
ifinal=Factores[Factores.Fecha==FechaCalculo].index.values

col=Factores.columns
Retornos=pd.DataFrame(columns=col)
Parametros=pd.DataFrame(index=('Mu','Sigma'),columns=col)

for y in col:
    if y not in ['Fecha','indice']:
        Retornos[y]=Factores.indice.apply(lambda x: math.log(Factores.iloc[int(x)][y]/Factores.iloc[int(x-t)][y]) if (Factores.iloc[int(x)].indice>=int(iinicial) and Factores.iloc[int(x)].indice<=int(ifinal)) else 0)    

Retornos=Retornos.iloc[int(iinicial):int(ifinal)+1]


for y in col:
    if y not in ['Fecha','indice']:
        Parametros[y][0]=Retornos[y].sum()/((int(ifinal)-int(iinicial)+1)*t)
        Parametros[y][1]=(math.sqrt(sum((Retornos[y]-Parametros[y][0]*t)**2)/(int(ifinal)-int(iinicial)+1)))/math.sqrt(t)

Parametros.drop(['Fecha','indice'], axis=1, inplace=True)
#print(Parametros)

Covar=Factores.iloc[int(iinicial):int(ifinal)+1].drop(['Fecha','indice'], axis=1).cov()

Epsilon=0.00001
Covar=Covar+Epsilon*np.identity(len(Covar.columns))
Chol=np.linalg.cholesky(Covar)


Aleat=np.random.normal(loc=0,scale=1,size=(n,len(Covar.columns)))
AleatCorrel=Aleat @ Chol
#print(AleatCorrel)


def exponenciar(x):
    return pd.Series(np.exp(x))

#PreciosSimulados=pd.DataFrame(np.zeros(shape=(n,len(Covar.columns))), columns=Covar.columns)


#esto funciona pero se puede hacer mas rapido con dict y dataframe al final
#Simulaciones=pd.DataFrame(columns=Covar.columns)
#for j in range(0,n):
 #   Simulacion=np.multiply(Precios.iloc[int(ifinal)][1:-1].values,(Parametros.iloc[0]*t+np.multiply(Parametros.iloc[1],AleatCorrel[j])).apply(exponenciar).transpose())
  #  Simulaciones=Simulaciones.append(Simulacion)


SimulacionesFactores=pd.DataFrame(index=range(0,n),columns=Covar.columns)
j=0
while j<n:
    Simulacion=np.multiply(Factores.iloc[int(ifinal)][1:-1].values,(Parametros.iloc[0]*t+np.multiply(Parametros.iloc[1],AleatCorrel[j])).apply(exponenciar).transpose())
    SimulacionesFactores.iloc[[j]]=Simulacion.values
    j+=1
next  
#print(SimulacionesFactores)

# Hasta aca tenemos los factores de riesgo simulados. Ellos son, por ej: factor desc 30 curva ars, factor desc 60 curva ars,..., factor desc 30 curva usd, ..., badlar, cer, tipo de cambio, indice merval, etc.
# Luego, se aproximan los precios de los bonos que componen la cartera de trading. Eso se hace con un analisis de sensibilidad para la VARIACION de precio. Los parametros de dicho analisis deben ser actualizados periodicamente.
# Por ejemplo, el precio simulado del bono x a 10 dias sera igual a: el precio de dicho bono a la fecha de calcula + 0.3 * variacion del factor de descuento a 30 dias de la curva de ARS.

DeltaFactores=SimulacionesFactores-Factores.iloc[int(ifinal)][1:-1].values
#print(DeltaFactores)


Precios=pd.read_excel('Montecarlo.xlsx', sheet_name='Precios')
Precios.reset_index
Precios['indice']=Precios.index  
PreciosBase=Precios.iloc[int(ifinal):int(ifinal)+1].drop(['Fecha','indice'],axis=1)

#print(PreciosBase)

Sensibilidades=pd.read_excel('Montecarlo.xlsx', sheet_name='Sensibilidades')
Sensibilidades.reset_index
Sensibilidades=Sensibilidades.drop(['Bono'], axis=1)
#print(Sensibilidades)

PreciosSimulados=pd.DataFrame(index=range(0,n), columns=PreciosBase.columns)

j=0
for j in range(0,n):
    for i in range(0,len(Sensibilidades)):
        PreciosSimulados.iloc[j,i]=PreciosBase.iloc[0,i]+np.sum(np.multiply(Sensibilidades.iloc[i].values, DeltaFactores.iloc[j]).values)

#print(PreciosSimulados)

Nominales=pd.read_excel('Montecarlo.xlsx', sheet_name='Nominales')
Nominales.reset_index
Nominales['indice']=Nominales.index  
NominalesBase=Nominales.iloc[int(ifinal):int(ifinal)+1].drop(['Fecha','indice'],axis=1)


ValorCarteraBase=np.sum(np.multiply(NominalesBase,PreciosBase), axis=1).values
#print(ValorCarteraBase)


ValorCarteraSimulada=pd.DataFrame(index=range(0,n),columns=['Simulacion'])
j=0
for j in range(0,n):
    Simul=np.sum(np.multiply(NominalesBase,PreciosSimulados.iloc[j]), axis=1).values
    ValorCarteraSimulada.iloc[j]=Simul

    
#print(ValorCarteraSimulada)

Resultados=pd.DataFrame(ValorCarteraSimulada.values-int(ValorCarteraBase[0]))
print(Resultados)

Var=-np.percentile(Resultados,(1-NivelConfianza)*100)
print(Var)

ES=-np.mean(pd.DataFrame(Resultados[Resultados<-Var]).dropna()).values[0]
print(ES)

Output=[FechaCalculo, Var, ES]
#print(Output)


Archivo=xlwings.Book('Montecarlo.xlsx')
Solapa=Archivo.sheets[5]
Archivo.sheets[5].range('A100000:C100000').end('up').offset(1,0).value=Output
Archivo.save('Montecarlo.xlsx')