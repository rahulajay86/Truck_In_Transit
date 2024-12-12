import os
import xlwings as xw
from datetime import datetime, timedelta
import logging

# Set up logging to write to a notepad (log file)
log_file_path = r'D:\Neelitech\MIS\python\Truck In Transit\log_file.txt'  # Define the path to your log file
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
        prev_date = (datetime.now() - timedelta(days=1)).strftime('%d')

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


# Function to find cells with #N/A in columns W to AE for a given row and print them
def find_and_print_na_cells_in_row(sheet, row):
    na_cells = []
    try:
        for col in range(23, 31):  # Columns W (23) to AE (30)
            cell_value = sheet.range((row, col)).value
            if cell_value == '#N/A':
                na_cells.append(sheet.range((row, col)).address)
                print(f"Column {sheet.range((row, col)).get_address(False, True)} has #N/A")
    except Exception as e:
        logging.error(f"Error in finding #N/A cells in row {row}: {e}")
    return na_cells


# Original input and output file paths
input_file = r'D:\Neelitech\MIS\python\Truck In Transit\New folder\FTL.xlsx'
output_file = r'D:\Neelitech\MIS\python\Truck In Transit\New folder\Order ageing.xlsb'

# Append previous date to the input file name
input_file_with_date = append_previous_date_to_filename(input_file)

# Log the input and output file paths
logging.info(f"Input file with previous day's date: {input_file_with_date}")
logging.info(f"Output file: {output_file}")

# Check if the input file exists
if not os.path.exists(input_file_with_date):
    # Notify the user that the file does not exist
    notify_file_missing(input_file_with_date)
else:
    try:
        # Proceed with reading the file and processing if it exists
        with xw.App(visible=False) as app:  # Open Excel in the background
            wb_input = app.books.open(input_file_with_date)  # Open the source workbook
            sheet_input = wb_input.sheets['FTL']  # Access the 'FTL' sheet

            # Get the last row with data in column A
            last_row_input = sheet_input.range('A1').end('down').row

            # Read all data from columns A to V in one operation
            data = sheet_input.range(f'A2:V{last_row_input}').value

            # Log the copied data range
            logging.info(f"Copied data from {sheet_input.name}: A2:V{last_row_input}")

            wb_input.close()

        # Write data to the 'Truck In transit' sheet in the output .xlsb file using xlwings
        with xw.App(visible=False) as app:  # Open Excel in the background
            wb_output = app.books.open(output_file)  # Open the destination workbook
            sheet_output = wb_output.sheets['TIT F']  # Get the target sheet

            # Find the last row with data in column A of "TIT F"
            last_row_output = sheet_output.range('A1').end('down').row

            # Clear existing data from columns A to V starting from row 2 to the last row
            sheet_output.range(f'A2:V{last_row_output}').clear_contents()
            logging.info(f"Cleared existing data from {sheet_output.name}: A2:V{last_row_output}")

            # Paste new data in one operation
            sheet_output.range(f'A2:V{1 + len(data)}').value = data
            logging.info(f"Pasted data to {sheet_output.name}: A2:V{1 + len(data)}")

            # Extend formulas from W2 to AE2 down to the last row of data in one operation
            formula_source = sheet_output.range('W2:AE2').formula  # Get the formulas from W2 to AE2
            last_row_data = 1 + len(data)  # Last row to apply the formulas
            sheet_output.range(f'W2:AE{last_row_data}').formula = formula_source  # Apply to entire range

            # Log the range where formulas are applied
            logging.info(f"Formulas applied in {sheet_output.name}: W2:AE{last_row_data}")

            # Save the workbook
            wb_output.save()

            # Now let's read back the pasted data (for confirmation)
            pasted_data = sheet_output.range(f'A2:V{last_row_data}').value  # Read back the data

            # Find the first empty cell in column V
            first_empty_cell = find_first_empty_cell_in_column(sheet_output, 'V')
            if first_empty_cell:
                logging.info(f"First empty cell in column V: {first_empty_cell.address}")
                print(f"First empty cell in column V: {first_empty_cell.address}")

                # Find cells with #N/A in columns W to AE for this row and print them along with their columns
                na_cells = find_and_print_na_cells_in_row(sheet_output, first_empty_cell.row)
                if na_cells:
                    logging.info(f"#N/A cells in columns W to AE for row {first_empty_cell.row}: {na_cells}")
                    print(f"#N/A cells found in columns W to AE for row {first_empty_cell.row}")
                else:
                    print(f"No cells with #N/A found in columns W to AE for row {first_empty_cell.row}")

            # Close the workbook
            wb_output.close()

        # Print the pasted data
        print("Pasted Data in 'Truck In transit' sheet (columns A to V):")
        for row in pasted_data:
            print(row)

    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")
