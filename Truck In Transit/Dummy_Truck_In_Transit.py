import xlwings as xw
from pyxlsb import open_workbook


# Path to your .xlsb files
input_file = r'D:\Neelitech\MIS\python\Truck In Transit\Order ageing.xlsb'
output_file = r'D:\Neelitech\MIS\python\Truck In Transit\Order ageing (1).xlsb'

# Initialize an empty list to store the data
data = []

# Read data from the 'Truck In transit' sheet in Order ageing.xlsb using pyxlsb
with open_workbook(input_file) as wb:
    with wb.get_sheet('Truck In transit') as sheet:
        # Detect the last row with data in the sheet
        for row_idx, row in enumerate(sheet.rows(), start=1):
            if any(item.v is not None for item in row[:22]):  # Check if there's data in columns A to V
                if row_idx > 1:  # Skip the header row (row 1)
                    data.append([item.v for item in row[:22]])  # Get values from columns A to V

# Write data to the 'Truck In transit' sheet in Order ageing (1).xlsb using xlwings
with xw.App(visible=False) as app:  # Open Excel in the background
    wb = app.books.open(output_file)  # Open the destination workbook
    sheet = wb.sheets['Truck In transit']  # Get the target sheet

    # Write the data starting from row 2 (since row 1 is headers)
    for row_idx, row_data in enumerate(data, start=2):  # start=2 to begin writing from row 2
        sheet.range(f'A{row_idx}:V{row_idx}').value = row_data  # Write row data to columns A to V

    # Now extend formulas in column W from row 2 to the last row of data in column V
    last_row = 1 + len(data)  # Find the last row where data was pasted
    formula_source = sheet.range('W2').formula  # Get the formula from cell W2
    sheet.range(f'W2:W{last_row}').formula = formula_source  # Drag the formula down to the last row

    # Save the workbook
    wb.save()

    # Now let's read back the pasted data from the output file and print it
    pasted_data = []
    for row_idx in range(2, 2 + len(data)):  # Read back the rows from 2 onwards
        pasted_row = sheet.range(f'A{row_idx}:V{row_idx}').value  # Read columns A to V
        pasted_data.append(pasted_row)

    # Close the workbook
    wb.close()

# Print the pasted data
print("Pasted Data in 'Transit In transit' sheet (columns A to V):")
for row in pasted_data:
    print(row)
