# Importer xlwings
import xlwings as xw

# Samler de to excelark i tabel.xlsx
wbcopy = xw.Book('ff-mnedsrapport-data-opl.xlsx')
wb = xw.Book('tabel.xlsx')
wbcopy.sheets[0].copy(after=wb.sheets[0])
wbcopy.close()

# Skifter aktive ark til "Gns. temperatur"
sheet1 = wb.sheets['Gns. temperatur']

# Indsætter 4 kolonner
sheet1.range('L:O').insert()

# Finder antal rækker i dokumentet
num_row1 = sheet1.range('A1').end('down').row

# Beregner vægtet fremløbstemperatur
sheet1['L3'].value = 'Vægtet fremtemp. C°'
sheet1.range('L4:L' + str(num_row1)).formula = '=P4/K4'

# Beregner vægtet returtemperatur
sheet1['M3'].value = 'Vægtet returtemp. C°'
sheet1.range('M4:M' + str(num_row1)).formula = '=Q4/K4'

# Beregner den samlede volumen
sheet1['N1'].value = 'Samlet volumen'
sheet1['N2'].value = '=SUM(K4:K' + str(num_row1)

# Beregner en vægtet fremløbstemperatur ud fra den samlede volumen
sheet1['N3'].value = 'Samlet vægtet frem temp. C°'
sheet1.range('N4:N' + str(num_row1)).formula = '=(K4/$N$2)*L4'
sheet1.range('N' + str(num_row1)).offset(row_offset=1).value = '=SUM(N4:N' + str(num_row1)

# Beregner en vægtet returtemperatur ud fra den samlede volumen
sheet1['O3'].value = 'Samlet vægtet retur temp. C°'
sheet1.range('O4:O' + str(num_row1)).formula = '=(K4/$N$2)*M4'
sheet1.range('O' + str(num_row1)).offset(row_offset=1).value = '=SUM(O4:O' + str(num_row1)

# Skifter aktive ark til "Ark1"
sheet2 = wb.sheets['Ark1']

# Finder antal rækker i dokumentet
num_row2 = sheet2.range('A1').end('down').row

# Beregner effekt fra Hedelund
sheet2['N2'].value = 'Hedelund MW'
sheet2.range('N3:N' + str(num_row2)).formula = '=((((ABS(B3)*980)*4.187*(C3-D3)))/3600)/1000'

# Beregner effekt fra City Vest
sheet2['O2'].value = 'City Vest MW'
sheet2.range('O3:O' + str(num_row2)).formula = '=((((ABS(E3)*980)*4.187*(F3-G3)))/3600)/1000'

# Beregner effekt fra Varde
sheet2['P2'].value = 'Varde MW'
sheet2.range('P3:P' + str(num_row2)).formula = '=((((ABS(H3)*980)*4.187*(I3-J3)))/3600)/1000'

# Beregner effekt fra City Øst
sheet2['Q2'].value = 'City Øst MW'
sheet2.range('Q3:Q' + str(num_row2)).formula = '=((((ABS(K3)*980)*4.187*(L3-M3)))/3600)/1000'

# Beregner samlet effekt
sheet2['R2'].value = 'Samlet MW'
sheet2.range('R3:R' + str(num_row2)).formula = '=SUM(N3:Q3)'

# Beregner fremløbstemperatur i Hjerting
sheet2['S2'].value = 'Fremløb Hjerting C°'

# Beregningsloop
output_range = f'S3:S{num_row2}'
condition_range = f'R3:R{num_row2}'

for row_num in range(3, num_row2 + 1):
    value = wb.sheets[sheet2].range(f'R{row_num}').value
    if value < 80:
        wb.sheets[sheet2].range(f'S{row_num}').value = 75
    elif 80 <= value < 100:
        wb.sheets[sheet2].range(f'S{row_num}').value = 73
    elif 100 <= value < 150:
        wb.sheets[sheet2].range(f'S{row_num}').value = 73
    elif 150 <= value < 200:
        wb.sheets[sheet2].range(f'S{row_num}').value = 73
    elif value >= 200:
        wb.sheets[sheet2].range(f'S{row_num}').value = 75

# finder den gennemsnitlige fremløbstemperatur over måneden
sheet2.range('S' + str(num_row2)).offset(row_offset=1).value = '=AVERAGE(S3:S' + str(num_row2)

# Laver nyt ark til resultatet
sheet3 = wb.sheets.add('result')
sheet3['A1'].value = 'Tab i fremløbstemperatur'
sheet3.range('B1').value = sheet2.range('S' + str(num_row2)).offset(row_offset=1).value - \
                           sheet1.range('N' + str(num_row1)).offset(row_offset=1).value

# Gemmer excelfilen
wb.save()

# Lukker excelfilen
wb.close()
