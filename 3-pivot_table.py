import pandas as pd

# importando dados
data = pd.read_excel("data/CarSales.xlsx")
# print(type(data))

# selecionando colunas específicas do dataframe
df = data[["Fabricante", "ValorVenda", "Ano"]]
print(df)

# criando tabela pivô
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)

# exportando tabela pivô em arquivo excel
pivot_table.to_excel("data/pivot_table.xlsx", "Relatorio")