import pandas as pd

# Load the two Excel files
file1_path = r"D:\Neelitech\MIS\python\Input_Files\merge\EXPORT1.XLSX"
file2_path = r"D:\Neelitech\MIS\python\Input_Files\merge\EXPORT2.XLSX"
file3_path = r"D:\Neelitech\MIS\python\Input_Files\merge\EXPORT3.XLSX"

# Read the sheets (assuming first sheet in both files)
file1_data = pd.read_excel(file1_path, sheet_name=0)
file2_data = pd.read_excel(file2_path, sheet_name=0)
file3_data = pd.read_excel(file3_path, sheet_name=0)

# Merge the two dataframes by concatenating rows
merged_data = pd.concat([file1_data, file2_data,file3_data], ignore_index=True)

# Save the merged data to a new Excel file
merged_file_path = 'D:\\Neelitech\\MIS\python\\Input_Files\\merge\\2024\\merged_files.xlsx'
merged_data.to_excel(merged_file_path, index=False)

print(f"Files merged and saved to: {merged_file_path}")
