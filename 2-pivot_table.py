import pandas as pd

#1 - importando dados
data = pd.read_excel ("data/VendaCarros.xlsx")
# print(type(data))

#2- selecionar colunas especificas do dataframe (df)
df = data[["Fabricante", "ValorVenda","Ano"]]
print (df)

#3- criação de tabela pivo (no excel Tab. Dinamica)
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)
print(pivot_table)

#4 - exportando tabela pivo em arquivo excel

pivot_table.to_excel("data/pivot_table.xlsx", "Relatório")


