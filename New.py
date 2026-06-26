"""
=============================================================================
Monitor Log to Excel Writer (WITH FUNCTIONS)
=============================================================================
This script does 3 main things:
1. Extracts a zip file and reads CSV data
2. Calculates average values from the CSV
3. Writes those values to an Excel calculator file

Created: February 2026
=============================================================================
"""

# Import libraries (like adding tools to your toolbox)
import os  # Work with files and folders
import zipfile  # Extract zip files
import glob  # Find files using patterns
import pandas as pd  # Read and analyze CSV files (Excel-like data)
import win32com.client  # Control Excel from Python
import time  # Add delays/pauses
import Test_Data  # Your test data (custom file)


# =============================================================================
# FUNCTION 1: Find and Extract the Newest Zip File
# =============================================================================
def find_and_extract_zip():
    """
    This function finds the newest zip file in Downloads and extracts it.

    Returns:
        extract_path: Where the files were extracted to
    """
    print("📦 PART 1: Finding and extracting zip file\n")

    # Where to look for zip files (your Downloads folder)
    download_path = os.path.join(os.path.expanduser("~"), "Downloads")
    print(f"   Looking in: {download_path}")

    # Find all .zip files in Downloads folder
    zip_files = glob.glob(os.path.join(download_path, "*.zip"))

    # Check if we found any zip files
    if not zip_files:
        print("❌ ERROR: No zip files found in Downloads!")
        exit()  # Stop the program

    # Get the newest zip file (the one created most recently)
    newest_zip = max(zip_files, key=os.path.getctime)
    print(f"   ✅ Found newest zip: {os.path.basename(newest_zip)}\n")

    # Where to extract the zip file
    extract_to = os.path.join(download_path, "extracted")
    os.makedirs(extract_to, exist_ok=True)  # Create folder if it doesn't exist

    # Extract (unzip) the file
    print(f"   Extracting to: {extract_to}")
    with zipfile.ZipFile(newest_zip, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"   ✅ Extraction complete!\n")

    return extract_to  # Return the path where files were extracted


# =============================================================================
# FUNCTION 2: Find and Read the CSV File
# =============================================================================
def find_and_read_csv(extract_path):
    """
    This function finds the newest CSV file and reads it.

    Parameters:
        extract_path: Where to look for CSV files

    Returns:
        df: DataFrame (like an Excel spreadsheet) with the CSV data
        csv_file: Path to the CSV file
    """
    print("📊 PART 2: Finding and reading CSV file\n")

    # Find all .csv files in the extracted folder (including subfolders)
    csv_files = glob.glob(os.path.join(extract_path, "**", "*.csv"), recursive=True)

    # Check if we found any CSV files
    if not csv_files:
        print("❌ ERROR: No CSV files found after extraction!")
        exit()  # Stop the program

    # Get the latest CSV file (the one modified most recently)
    csv_file = max(csv_files, key=os.path.getmtime)
    print(f"   ✅ Found CSV: {os.path.basename(csv_file)}\n")

    # Read the CSV file into a "DataFrame" (think of it like an Excel spreadsheet)
    df = pd.read_csv(csv_file)
    print(f"   CSV has {len(df)} rows and {len(df.columns)} columns\n")

    return df, csv_file  # Return both the data and the file path


# =============================================================================
# FUNCTION 3: Calculate Averages from CSV Data
# =============================================================================
def calculate_averages(df):
    """
    This function calculates the average of each numeric column.

    Parameters:
        df: DataFrame (spreadsheet) with the data

    Returns:
        averages_dict: Dictionary with column names and their averages
    """
    print("📊 PART 3: Calculating averages\n")

    # Create an empty dictionary to store averages
    # Dictionary is like a phone book: {name: phone_number}
    # Here it will be: {column_name: average_value}
    averages_dict = {}

    # Go through each column in the CSV
    for column in df.columns:
        column_data = df[column]  # Get all values in this column

        # Check if the column contains numbers (not text)
        if pd.api.types.is_numeric_dtype(column_data):
            # Calculate the average (mean) and round to 5 decimal places
            average_value = round(float(column_data.mean()), 5)

            # Store in dictionary: column_name → average_value
            averages_dict[column] = average_value

    print(f"   ✅ Calculated averages for {len(averages_dict)} columns\n")

    # Print all averages
    print("Averages Dictionary:", averages_dict)

    return averages_dict  # Return the dictionary of averages


# =============================================================================
# FUNCTION 4: Open Excel File
# =============================================================================
def open_excel_file(excel_file_path):
    """
    This function opens the Excel file and prepares it for writing.

    Parameters:
        excel_file_path: Full path to the Excel file

    Returns:
        excel: The Excel application object
        workbook: The opened workbook
        sheet: The active worksheet
    """
    print("=" * 70)
    print("📝 PART 4: Opening Excel file")
    print("=" * 70 + "\n")

    # Open Excel application
    print("   Opening Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Make Excel window visible
    excel.DisplayAlerts = False  # Don't show Excel pop-up messages

    # Open the workbook (Excel file)
    workbook = excel.Workbooks.Open(excel_file_path)
    sheet = workbook.ActiveSheet  # Get the active worksheet

    print(f"   ✅ Opened: {os.path.basename(excel_file_path)}")
    print(f"   ✅ Active Sheet: {sheet.Name}\n")

    return excel, workbook, sheet  # Return all three objects


# =============================================================================
# FUNCTION 5: Prepare Excel for Writing
# =============================================================================
def prepare_excel_for_writing(excel):
    """
    This function turns off automatic calculation to prevent errors.

    Parameters:
        excel: The Excel application object
    """
    print("🔧 PART 5: Preparing Excel\n")

    # Turn OFF automatic calculation
    # This prevents Excel from calculating while we're writing values
    # (which can cause errors)
    excel.Calculation = -4135  # -4135 = Manual calculation mode
    excel.ScreenUpdating = False  # Don't update screen (faster)
    excel.EnableEvents = False  # Don't trigger Excel events/macros

    print("   ✅ Automatic calculation disabled")
    print("   ✅ Ready to write values\n")


# =============================================================================
# FUNCTION 6: Write Test Data Values to Excel
# =============================================================================
def write_test_data(sheet):
    """
    This function writes Test_Data values to specific Excel cells.

    Parameters:
        sheet: The Excel worksheet to write to

    Returns:
        count: Number of cells written
    """
    print("📝 PART 6: Writing Test_Data values\n")

    # This is a list of Excel cells and the values to write
    # Format: 'Cell': (value, description)
    test_data_cells = {
        'C5': (Test_Data.my_list[0], 'P_act - Pressure'),
        'C6': (Test_Data.my_list[1], 'T_act - Temperature'),
        'C7': (Test_Data.my_list[2], 'Base Temperature'),
        'H5': (Test_Data.my_list[3], 'Parameter H5'),
        'H6': (Test_Data.my_list[4], 'Parameter H6'),
        'H9': (Test_Data.my_list[7], 'Parameter H9'),
        'H10': (Test_Data.my_list[8], 'Parameter H10'),
        'H11': (Test_Data.my_list[9], 'Parameter H11'),
        'C29': (Test_Data.my_list[10], 'Parameter C29'),
        'C32': (Test_Data.my_list[11], 'Parameter C32'),
    }

    count = 0  # Keep track of how many cells we write

    # Go through each cell and write the value
    for cell_address, (value, description) in test_data_cells.items():
        # Check if the cell has a formula
        # (we don't want to overwrite formulas!)
        if not sheet.Range(cell_address).HasFormula:
            # Write the value to the cell
            sheet.Range(cell_address).Value = value
            print(f"   ✅ {cell_address} = {value} ({description})")
            count += 1
        else:
            # Cell has a formula, skip it
            print(f"   🔒 {cell_address} has formula, skipped")

    return count  # Return how many cells were written


# =============================================================================
# FUNCTION 7: Write Monitor Log Averages to Excel
# =============================================================================
def write_monitor_log_averages(sheet, averages_dict):
    """
    This function writes the calculated averages to Excel cells.

    Parameters:
        sheet: The Excel worksheet to write to
        averages_dict: Dictionary with column names and averages

    Returns:
        count: Number of cells written
    """
    print("\n📊 PART 7: Writing monitor log averages\n")

    # This is a list of Excel cells and which CSV column to get data from
    # Format: 'Cell': 'Column Name in CSV'
    monitor_log_cells = {
        'C30': 'Path 1 Transit Time Up (µs)',
        'C31': 'Path 1 Transit Time Down (µs)',
        'C54': 'Path 1 Sound Speed (m/s)',
        'G54': 'Composite 2 Sound Speed (m/s)',
        'G55': 'Composite 2 Molecular Weight (g/mole)',
        'C56': 'Path 1 Raw Velocity (m/s)',
        'C57': 'Path 1 Velocity (m/s)',
        'G57': 'Composite 2 Velocity (m/s)',
        'C58': 'Path 1 Reynolds Number',
        'G58': 'Composite 2 Reynolds Number',
        'G59': 'Composite 2 Actual Volumetric Flow Rate (m³/s)',
        'G60': 'Composite 2 Standard Volumetric Flow Rate (Sm³/s)',
        'G61': 'Composite 2 Mass Flow Rate (kg/s)',
        'D30' : 'Path 2 Transit Time Up (µs)',
        'D31' : 'Path 2 Transit Time Down (µs)',
        'D54' : 'Path 2 Sound Speed (m/s)',
        'D56' : 'Path 2 Raw Velocity (m/s)',
        'D57' : 'Path 2 Velocity (m/s)',
        'D58' : 'Path 2 Reynolds Number',
        'E30' : 'Path 3 Transit Time Up (µs)',
        'E31' : 'Path 3 Transit Time Down (µs)',
        'E54' : 'Path 3 Sound Speed (m/s)',
        'E56' : 'Path 3 Raw Velocity (m/s)',
        'E57' : 'Path 3 Velocity (m/s)',
        'E58' : 'Path 3 Reynolds Number',
        'F30' : 'Path 4 Transit Time Up (µs)',
        'F31' : 'Path 4 Transit Time Down (µs)',
        'F54' : 'Path 4 Sound Speed (m/s)',
        'F56' : 'Path 4 Raw Velocity (m/s)',
        'F57' : 'Path 4 Velocity (m/s)',
        'F58' : 'Path 4 Reynolds Number',
    }

    count = 0  # Keep track of how many cells we write

    # Go through each cell and write the average value
    for cell_address, column_name in monitor_log_cells.items():
        # Check if we have data for this column
        if column_name in averages_dict:
            value = averages_dict[column_name]

            # Check if the value is valid (not None, not NaN)
            if value and not pd.isna(value):
                # Check if cell has a formula
                if not sheet.Range(cell_address).HasFormula:
                    # Write the value
                    sheet.Range(cell_address).Value = value
                    print(f"   ✅ {cell_address} = {value} ({column_name})")
                    count += 1
                else:
                    print(f"   🔒 {cell_address} has formula, skipped")
            else:
                print(f"   ⚠️  {cell_address} - value is invalid")
        else:
            print(f"   ⚠️  {cell_address} - column not found in CSV")

    print(f"\n   ✅ All values written!\n")

    return count  # Return how many cells were written


# =============================================================================
# FUNCTION 8: Calculate Excel Formulas
# =============================================================================
def calculate_excel(excel):
    """
    This function re-enables calculation and calculates the workbook.

    Parameters:
        excel: The Excel application object
    """
    print("=" * 70)
    print("🔄 PART 8: Calculating Excel formulas")
    print("=" * 70 + "\n")

    # Wait a moment for all values to be fully written
    time.sleep(2)

    # Turn automatic calculation back ON
    excel.EnableEvents = True
    excel.Calculation = -4105  # -4105 = Automatic calculation mode

    print("   Re-enabled automatic calculation\n")

    # Calculate the workbook 3 times with delays
    # (This ensures all formulas calculate correctly, including VBA macros)
    for pass_number in [1, 2, 3]:
        print(f"   🔄 Calculation pass {pass_number}...")

        try:
            excel.Calculate()  # Tell Excel to calculate
            print(f"   ✅ Pass {pass_number} complete")

            # Wait between calculations
            if pass_number == 1:
                time.sleep(3)  # Wait 3 seconds after first pass
            else:
                time.sleep(2)  # Wait 2 seconds after other passes

        except Exception as error:
            print(f"   ⚠️  Warning in pass {pass_number}: {error}")

    # Turn screen updating back ON
    excel.ScreenUpdating = True

    print(f"\n   ✅ All calculations complete!\n")


# =============================================================================
# FUNCTION 9: Save and Close Excel
# =============================================================================
def save_and_close_excel(workbook, excel):
    """
    This function saves the workbook and closes Excel.

    Parameters:
        workbook: The Excel workbook to save
        excel: The Excel application to close
    """
    print("=" * 70)
    print("💾 PART 9: Saving and closing Excel")
    print("=" * 70 + "\n")

    # Save the workbook
    try:
        workbook.Save()
        print("   ✅ Workbook saved")
    except Exception as error:
        print(f"   ❌ Save error: {error}")

    # Close the workbook
    try:
        workbook.Close(SaveChanges=False)  # Don't save again (already saved)
        print("   ✅ Workbook closed")
    except Exception as error:
        print(f"   ❌ Close error: {error}")

    # Quit Excel application
    try:
        excel.Quit()
        print("   ✅ Excel closed")
    except Exception as error:
        print(f"   ❌ Quit error: {error}")


# =============================================================================
# MAIN PROGRAM - This is where everything happens!
# =============================================================================
def main():
    """
    This is the main function that runs the entire program.
    It calls all the other functions in the right order.
    """
    # Print a nice header
    print("\n" + "=" * 70)
    print("📦 MONITOR LOG TO EXCEL WRITER")
    print("=" * 70 + "\n")

    # STEP 1: Find and extract the zip file
    extract_path = find_and_extract_zip()

    # STEP 2: Find and read the CSV file
    df, csv_file = find_and_read_csv(extract_path)

    # STEP 3: Calculate averages from the CSV
    averages_dict = calculate_averages(df)

    # Print all averages
    print("Averages Dictionary:", averages_dict)

    # STEP 4: Open the Excel file
    excel_file = r"C:\Users\malproh01\Downloads\XGF1000_Calculator_rev1.4c.xlsm"
    excel, workbook, sheet = open_excel_file(excel_file)

    # STEP 5: Prepare Excel for writing
    prepare_excel_for_writing(excel)

    # STEP 6: Write Test_Data values
    test_data_count = write_test_data(sheet)

    # STEP 7: Write monitor log averages
    monitor_log_count = write_monitor_log_averages(sheet, averages_dict)

    # STEP 8: Calculate Excel formulas
    calculate_excel(excel)

    # STEP 9: Save and close Excel
    save_and_close_excel(workbook, excel)

    # STEP 10: Print summary
    print("\n" + "=" * 70)
    print("✅ COMPLETE - ALL DONE!")
    print("=" * 70)
    print(f"   CSV File: {os.path.basename(csv_file)}")
    print(f"   Rows Processed: {len(df)}")
    print(f"   Test_Data Cells Written: {test_data_count}")
    print(f"   Monitor Log Cells Written: {monitor_log_count}")
    print(f"   Excel File Updated: {os.path.basename(excel_file)}")
    print("=" * 70 + "\n")

    print("🎉 Success! You can now check your Excel file.\n")


# =============================================================================
# START THE PROGRAM
# =============================================================================
# This checks if the script is being run directly (not imported)
if __name__ == "__main__":
    main()  # Run the main function
