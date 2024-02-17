from openpyxl import load_workbook

#1 le pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatório"]

#2 Acessando um valor especifico
# print (sheet["B3"].value)

#3 forma automatizada de interar valores por meio de loop
for i in range (2,6): #em python o ultimo digito do range é excluido. por isso colocar 6 ao inves de 5
    ano = sheet["A%s" %i].value
    am = sheet["B%s" %i].value
    bt = sheet["C%s" %i].value
    jg = sheet["D%s" %i].value
    mgb = sheet["E%s" %i].value
    rr = sheet["F%s" %i].value
    tvr = sheet["G%s" %i].value
    tr = sheet["H%s" %i].value
    print ("{0} a Aston Martin vendeu {1} a Bentley vendeu {2} a Jaguar vendeu {3} a MGB vendeu {4} a Rolls Royce vendeu {5} a TVR vendeu {6} e a Triumph vendeu {7}".format(ano, am, bt, jg,mgb,rr,tvr,tr))
    