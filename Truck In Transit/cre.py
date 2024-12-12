import os
import xlwings as xw
from datetime import datetime, timedelta
import logging

# Set up logging to write to a notepad (log file)
log_file_path = r'D:\Neelitech\MIS\python\Creditor Overdue\log_file.txt'  # Define the path to your log file
logging.basicConfig(
    filename=log_file_path,  # Log to a file
    level=logging.INFO,  # Set the log level to INFO
    format='%(asctime)s - %(levelname)s - %(message)s'  # Format of the log messages
)

# Add a new line before the first log entry
with open(log_file_path, 'a') as f:
    f.write('\n')


# Function to append the previous day's date to the file name
def append_previous_date_to_filename(file_path):
    try:
        # Get the directory, file name, and extension
        directory, file_name = os.path.split(file_path)
        name, ext = os.path.splitext(file_name)

        # Get the previous day's date
        prev_date = (datetime.now() - timedelta(days=1)).strftime('%d-%m-%Y')
        # prev_date = (datetime.now() - timedelta(days=1)).strftime('%d')
        # Create new file name with previous date appended
        new_file_name = f"{name} {prev_date}{ext}"

        # Construct full path with the new name
        new_file_path = os.path.join(directory, new_file_name)

        return new_file_path
    except Exception as e:
        logging.error(f"Error in generating file name: {e}")
        return None


# Function to notify user if a file does not exist
def notify_file_missing(file_path):
    error_message = f"Error: The file '{file_path}' does not exist. Please check the file name and path."
    logging.error(error_message)
    print(error_message)


# Function to find the first empty cell in column V
def find_first_empty_cell_in_column(sheet, column):
    try:
        # Find the first empty cell in the specified column
        first_empty_cell = sheet.range(f'{column}1').end('down').offset(1, 0)
        return first_empty_cell
    except Exception as e:
        logging.error(f"Error in finding first empty cell in column {column}: {e}")
        return None


# Original input and output file paths
input_file = r"D:\Neelitech\MIS\python\Creditor Overdue\EXPORT_CREDITOR_AGING.XLSX"
output_file =r"D:\Neelitech\MIS\python\Creditor Overdue\Creditors.XLSX"

# Append previous date to the input file name
# input_file_with_date = append_previous_date_to_filename(input_file)

# Log the input and output file paths
logging.info(f"Input file with previous day's date: {input_file}")
logging.info(f"Output file: {output_file}")

# Check if the input file exists
if not os.path.exists(input_file):
    # Notify the user that the file does not exist
    notify_file_missing(input_file)
else:
    try:
        # Proceed with reading the file and processing if it exists
        with xw.App(visible=False) as app:  # Open Excel in the background
            wb_input = app.books.open(input_file)  # Open the source workbook
            sheet_input = wb_input.sheets['sheet1']  # Access the 'FTL' sheet

            # Get the last row with data in column A
            last_row_input = sheet_input.range('A1').end('down').row

            # Read all data from columns A to V in one operation
            data = sheet_input.range(f'A2:T{last_row_input}').value

            print(f"Copied data from {sheet_input.name}: A2:T{last_row_input}")
            # Log the copied data range
            logging.info(f"Copied data from {sheet_input.name}: A2:T{last_row_input}")

            wb_input.close()

        # Write data to the 'Detail Vendor wise' sheet in the output .xlsb file using xlwings
        with xw.App(visible=False) as app:  # Open Excel in the background
            wb_output = app.books.open(output_file)  # Open the destination workbook
            sheet_output = wb_output.sheets['Detail Vendor wise']  # Get the target sheet

            # Find the last row with data in column A of "Detail Vendor wise"
            last_row_output = sheet_output.range('A1').end('down').row

            # Clear existing data from columns A to T starting from row 2 to the last row
            sheet_output.range(f'A2:T{last_row_output}').clear_contents()
            logging.info(f"Cleared existing data from {sheet_output.name}: A2:T{last_row_output}")

            print(f"Cleared existing data from {sheet_output.name}: A2:T{last_row_output}")
            # Paste new data in one operation
            sheet_output.range(f'A2:T{1 + len(data)}').value = data
            logging.info(f"Pasted data to {sheet_output.name}: A2:T{1 + len(data)}")
            print(f"Pasted data to {sheet_output.name}: A2:T{1 + len(data)}")

            # Add borders around the pasted data
            # Define the range where data is pasted
            pasted_range = sheet_output.range(f'A2:T{1 + len(data)}')

            # Use Excel API to set borders
            borders = pasted_range.api.Borders
            borders.LineStyle = 1  # 1 for continuous lines
            borders.Weight = 2  # 2 for thick borders
            borders.Color = 0  # Color black

            # Save the workbook
            wb_output.save()

            # Close the workbook
            wb_output.close()

        # # Print the pasted data
        # print("Pasted Data in 'Truck In transit' sheet (columns A to V):")
        # for row in data:
        #     print(row)

    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")
