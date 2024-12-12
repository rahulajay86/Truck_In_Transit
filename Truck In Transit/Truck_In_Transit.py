import os
import xlwings as xw
from datetime import datetime, timedelta
import shutil
import logging
from datetime import datetime
# Get the current date
current_date = datetime.now()
year = current_date.year
month = current_date.strftime('%m')  # format as zero-padded month (e.g., 10 for October)
day = current_date.strftime('%d')    # format as zero-padded day (e.g., 23)
today = datetime.today()

# Set up logging to write to a notepad (log file)
log_file_path = r'D:\Neelitech\MIS\python\Truck In Transit\New folder\log_file.txt'  # Define the path to your log file
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

            # Find the first empty cell in column V
            first_empty_cell = find_first_empty_cell_in_column(sheet_output, 'V')
            if first_empty_cell:
                logging.info(f"First empty cell in column V: {first_empty_cell.address}")
                print(f"First empty cell in column V: {first_empty_cell.address}")

                # Dynamically determine the last row with data across columns W to AE
                dynamic_end_row = max(sheet_output.range(f'{col}1').end('down').row for col in 'WXYZABCDEFGH')

                # Clear data only from columns W to AE, starting from the first empty row in column V to dynamic_end_row
                if first_empty_cell.row < dynamic_end_row:
                    sheet_output.range(f'W{first_empty_cell.row}:AE{dynamic_end_row}').clear_contents()
                    logging.info(f"Cleared data from W{first_empty_cell.row} to AE{dynamic_end_row}")
                else:
                    logging.info("No rows to clear as the first empty cell is at the last row of data.")

            # Dynamically add borders around the pasted data and center text in each column
            def apply_borders(sheet, start_row, end_row, start_col='A', end_col='V'):
                try:
                    data_range = sheet.range(f'{start_col}{start_row}:{end_col}{end_row}')
                    for border_id in range(7, 13):  # Border indices for all sides
                        data_range.api.Borders(border_id).LineStyle = 1  # xlContinuous
                        data_range.api.Borders(border_id).Weight = 2     # xlThin
                    logging.info(f"Borders applied around range {start_col}{start_row}:{end_col}{end_row}")
                except Exception as e:
                    logging.error(f"Error applying borders: {e}")

            def center_align_text(sheet, start_row, end_row, columns):
                try:
                    for col in columns:
                        sheet.range(f'{col}{start_row}:{col}{end_row}').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
                    logging.info(f"Text centered in range {columns[0]}{start_row}:{columns[-1]}{end_row}")
                except Exception as e:
                    logging.error(f"Error centering text: {e}")

            def clear_borders_below_data(sheet, last_data_row, start_col='A', end_col='V'):
                try:
                    clear_range = sheet.range(f'{start_col}{last_data_row + 1}:{end_col}1048576')
                    for border_id in range(7, 13):
                        clear_range.api.Borders(border_id).LineStyle = None  # Remove borders
                    logging.info(f"Borders cleared below the range {start_col}{last_data_row + 1}:{end_col}1048576")
                except Exception as e:
                    logging.error(f"Error clearing borders below data: {e}")

            # Apply borders, center text, and clear borders below data
            apply_borders(sheet_output, start_row=2, end_row=last_row_data, start_col='A', end_col='U')
            center_align_text(sheet_output, start_row=2, end_row=last_row_data, columns='ABCDEFGHIJKLMNOPQRSTUV')
            clear_borders_below_data(sheet_output, last_data_row=last_row_data, start_col='A', end_col='U')

            # Save the workbook
            wb_output.save()
            logging.info("Workbook saved with updated borders and alignment.")

            # Close the workbook
            wb_output.close()

    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")
