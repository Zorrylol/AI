# Importer xlwings
import xlwings as xw

# Open the source Excel workbooks
wbcopy = xw.Book('jojofremlbstemperaturer.xlsx')

# Open the target Excel workbook
wb = xw.Book('tabel.xlsx')

# Copy the first sheet from the second source workbook and name it
copied_sheet = wbcopy.sheets[0].copy(after=wb.sheets[0])
copied_sheet.name = 'jojofremlbstemperaturer'  # Change '' to the desired name

# Close the second source workbook
wbcopy.close()

# Skifter aktive ark til "Gns. temperatur"
sheet1 = wb.sheets['Gns. temperatur']

# Indsætter 4 kolonner
sheet1.range('L:O').insert()

# Finder antal rækker i dokumentet
num_row1 = sheet1.range('A1').end('down').row

# Beregner vægtet fremløbstemperatur
sheet1['L3'].value = 'Vægtet fremtemp. C°'
sheet1.range('L4:L' + str(num_row1)).formula = '=IF(OR(ISBLANK(P4), ISBLANK(K4), K4=0), "", P4/K4)'

# Beregner vægtet returtemperatur
sheet1['M3'].value = 'Vægtet returtemp. C°'
sheet1.range('M4:M' + str(num_row1)).formula = '=IF(OR(ISBLANK(Q4), ISBLANK(K4), K4=0), "", Q4/K4)'

# Beregner den samlede volumen
sheet1['N1'].value = 'Samlet volumen'
sheet1['N2'].value = '=SUM(K4:K' + str(num_row1)

# Beregner en vægtet fremløbstemperatur ud fra den samlede volumen
sheet1['N3'].value = 'Samlet vægtet frem temp. C°'
sheet1.range('N4:N' + str(num_row1)).formula = '=IF(OR(ISBLANK(K4), K4=0), "", (K4/$N$2)*L4)'
sheet1.range('N' + str(num_row1)).offset(row_offset=1).value = '=SUM(N4:N' + str(num_row1)

# Beregner en vægtet returtemperatur ud fra den samlede volumen
sheet1['O3'].value = 'Samlet vægtet retur temp. C°'
sheet1.range('O4:O' + str(num_row1)).formula = '=IF(OR(ISBLANK(K4), K4=0), "", (K4/$N$2)*M4)'
sheet1.range('O' + str(num_row1)).offset(row_offset=1).value = '=SUM(O4:O' + str(num_row1)

# Skifter aktive ark til "jojofremlbstemperaturer"
sheet2 = wb.sheets['jojofremlbstemperaturer']

# Finder antal rækker i dokumentet
num_row2 = sheet2.range('A1').end('down').row

# Beregner den gennemsnitlige fremløbstemperatur over måneden i Hjerting
sheet2.range('F' + str(num_row2)).offset(row_offset=1).value = '=AVERAGE(F2:F' + str(num_row2)

# Laver nyt ark til resultatet
sheet3 = wb.sheets.add('result')
sheet3['A1'].value = 'Tab i fremløbstemperatur'

sheet3['B1'].formula = sheet2.range('F' + str(num_row2)).offset(row_offset=1).value - \
                           sheet1.range('N' + str(num_row1)).offset(row_offset=1).value

# Gemmer excelfilen
wb.save()

# Lukker excelfilen
wb.close()