import pandas as pd

#1 - importando dados
data = pd.read_excel ("data/VendaCarros.xlsx")
print(data)

#2 - lista dos primeiros registros
print (data.head())

#3- lista dos ultimos registros
print (data.tail())

#4 - contagem de valores por fabricante
print (data["Fabricante"].value_counts())


