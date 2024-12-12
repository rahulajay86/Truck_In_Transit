import xlwings as xw

# Load the Excel workbook
file_path = r"D:\Neelitech\MIS\python\Input_Files\merge\2024\Sales Vol_Price Analysis Report.xlsx"
wb = xw.Book(file_path)


# Select the YTD and FTM Jul sheets
ytd_sheet = wb.sheets['YTD']
ftm_jul_sheet = wb.sheets['FTM Jul']
# ftm_jul_sheet = wb.sheets['FTM']

# Get the values in the range (E5:F11) from both sheets
ytd_data = ytd_sheet.range('E5:F11').value
ftm_jul_data = ftm_jul_sheet.range('E5:F11').value


# Print the values and sum them, then write the result to the YTD sheet
print("Summing values from YTD and FTM Jul sheets (E5:F11) and writing back to YTD:")

for idx in range(len(ytd_data)):
    # Get values from YTD sheet
    ytd_e_val = ytd_data[idx][0] or 0  # Column E
    ytd_f_val = ytd_data[idx][1] or 0  # Column F

    # Get values from FTM Jul sheet
    ftm_jul_e_val = ftm_jul_data[idx][0] or 0  # Column E
    ftm_jul_f_val = ftm_jul_data[idx][1] or 0  # Column F

    # Sum values from YTD and FTM Jul
    sum_e = ytd_e_val + ftm_jul_e_val
    sum_f = ytd_f_val + ftm_jul_f_val

    # Write the sum back to YTD sheet, starting at row 5 (idx + 5)
    ytd_sheet.range(f'E{idx + 5}').value = sum_e
    ytd_sheet.range(f'F{idx + 5}').value = sum_f

    print(f"Row {idx + 5}: Sum of Column E: {sum_e}, Sum of Column F: {sum_f} (written to YTD sheet)")

# Save the workbook after writing the values
wb.save(file_path)


# Close the workbook
wb.close()
