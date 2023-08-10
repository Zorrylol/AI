import pandas as pd
import xlwings as xw

csv_file = 'jojofremlbstemperaturer.csv'
excel_file = 'jojofremlbstemperaturer.xlsx'

#num_rows = len(pd.read_csv(csv_file))

csv = pd.read_csv(csv_file, sep=';', decimal=',') # ,nrows=num_rows hvis der er problemer med headers
csv.to_excel(excel_file, index=False)

print(f'The Excel file "{excel_file}" has been populated with the data from the CSV file "{csv_file}".')

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_file)

# Check for blank cells (missing values)
blank_cells = df.isnull().sum().sum()

# Print the number of blank cells found
print(f'The Excel file "{excel_file}" has {blank_cells} blank cells.')

# Connect to the Excel application
app = xw.App(visible=True)  # Start Excel application
wb = xw.Book(excel_file)    # Open the Excel file
ws = wb.sheets[0]           # Select the first sheet

# Find the used range of the sheet
used_range = ws.used_range

# Loop through each cell in the used range
for row in used_range.rows:
    for cell in row:
        if cell.value is None or cell.value == "":
            cell.color = (255, 0, 0)  # Highlight with red color

print(f'The Excel file "{excel_file}" has been updated with highlights of the blank cells.')

# Interpolate missing values using the default linear method
df_interpolated = df.interpolate()

# Write data starting from A2 to preserve the header
ws.range('A2').options(index=False, header=False).value = df_interpolated

print(f'The Excel file "{excel_file}" has been updated with interpolated data.')

# Count blank cells (missing values) in the Excel sheet
blank_cells_sheet = 0
for row in ws.range('A1').expand('table').rows:
    for cell in row:
        if cell.value is None or pd.isnull(cell.value):
            blank_cells_sheet += 1

# Print the number of blank cells found in the Excel sheet
print(f'The Excel file "{excel_file}" sheet has {blank_cells_sheet} blank cells.')

# Save the changes and close the workbook
wb.save()
wb.close()

# Close the Excel application
app.quit()
