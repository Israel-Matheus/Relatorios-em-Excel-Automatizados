from openpyxl import load_workbook

# lê pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"]

# acessando um valor específico
print(sheet["A3"].value)
print(sheet["B3"].value)

# iterando valores por meio de loop
for i in range(2, 6):
    year = sheet["A%s" %i].value # type: ignore
    aston = sheet["B%s" %i].value
    bentley = sheet["C%s" %i].value
    print("{0} o Aston Martin vendeu {1} e o Bentley vendeu {2}".format(year, aston, bentley))
