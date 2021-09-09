import openpyxl 

wb = openpyxl.load_workbook('planilha.xlsx', data_only=True)
dest_filename = 'richard_bryan_eulalio.xlsx'
sheet = wb['Planilha Vendas']

nutril = 0
lira_mg = 0

def faturmaneto_nutril(x,indice):
    global nutril
    if sheet['D' + str(indice)].value == 'Nutril':
        nutril += x        
    else:
        pass

def lira_bottom_mg(x, indice):
    global lira_mg
    if sheet['C' + str(indice)].value == 'Lira Batom':
        if sheet['E' + str(indice)].value == 'Fábrica MG':
            lira_mg += x
        else:
            pass
    else:
        pass
    
def faturamento_10():
    global sheet
    index = 2
    for f_new in sheet['M3':'M12']:
        index += 1
        for m_new in f_new:
            sheet['N' + str(index)] = '=LARGE(A:A;M' + str(index) + ')'

column = 1
for row in sheet['A2':'A169']:
    column += 1
    for row_new in row:
        if int(row_new.value) >= 0 and int(row_new.value) <= 1500:
           sheet['F' + str(column)] = 'Ruim'
           faturmaneto_nutril(int(row_new.value), column)
           lira_bottom_mg(int(row_new.value), column)
        else:
            if int(row_new.value) > 1500 and int(row_new.value) <= 3000:
                sheet['F' + str(column)] = 'Razoável'
                faturmaneto_nutril(int(row_new.value), column)
                lira_bottom_mg(int(row_new.value), column)
            else:
                sheet['F' + str(column)] = 'Boa'
                faturmaneto_nutril(int(row_new.value), column)
                lira_bottom_mg(int(row_new.value), column)

for i in range(11,15):
    sheet['I' + str(i)] = '=COUNTIFS(C:C;H' + str(i) + ')'
op = 3
for i in range(16,20): 
	sheet['N' + str(i)] = '=VLOOKUP(N'+ str(op) + ';A1:B169;2;FALSE)'
	sheet['O' + str(i)] = '=VLOOKUP(N' + str(op) + ';A1:D169;4;FALSE)'
	op += 1

sheet['H17'] = str(nutril)
sheet['H20'] = str(lira_mg)
faturamento_10()
print("Planilha " + dest_filename + " criada com sucesso.")
wb.save(filename = dest_filename)
