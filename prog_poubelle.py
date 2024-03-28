import openpyxl

"""
writebook = openpyxl.load_workbook('fichiers_xls/pour tony_30032024.xlsx')

sheet = writebook['Feuil1']

L = [i for i in range(1,1336)]

for i in range(1,sheet.max_row + 1):
    if sheet.cell(i,1).value in L:
        L.remove(sheet.cell(i,1).value)
print(L)
"""

writebook = openpyxl.load_workbook('fichiers_xls/1472_jeter.xlsx')
writebook2 = openpyxl.load_workbook('fichiers_xls/liste_n=1335.xlsx')

sheet = writebook['Feuil1']
sheet2 = writebook2['Feuil1']

L1 = []
L2 = [] 

for j in range(2, sheet2.max_row + 1):
    L2.append(sheet2.cell(j,4).value)

print(L2)

for i in range(2, sheet.max_row+1):
    if sheet.cell(i,7).value not in L2:
        L1.append(sheet.cell(i,5).value + sheet.cell(i,6).value)
print(len(L1),L1)



