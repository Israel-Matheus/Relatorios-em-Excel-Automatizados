import pandas as pd

# importando dados
data = pd.read_excel("data/CarSales.xlsx")
print(data)

# lista dos primeiros registros
print(data.head())

# lista os últimos registros
print(data.tail())

# contagem de valores por fabricante
print(data["Fabricante"].value_counts())