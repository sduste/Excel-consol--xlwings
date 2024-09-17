import os
import xlwings as xw

# Define the path to the Downloads folder
downloads_folder = os.path.expanduser("~/Downloads")

# Define the path to the consolidated workbook (already existing)
consol_workbook_path = os.path.join(downloads_folder, "Consolidated_Workbook.xlsx")

# Open Excel Application
excel = xw.App(visible=True)  # Keep Excel visible so that workbooks remain open
excel.display_alerts = False  # Disable any alerts to avoid interruption

# Open the consolidated workbook
consol_workbook = excel.books.open(consol_workbook_path)

# Sheet name counter to create unique names
sheet_counter = 1

# Loop through all Excel files in the Downloads folder
for filename in os.listdir(downloads_folder):
    if filename.endswith(".xlsx") and filename != "Consolidated_Workbook.xlsx" and not filename.startswith('~$'):
        file_path = os.path.join(downloads_folder, filename)
        
        try:
            # Open each source workbook
            source_workbook = excel.books.open(file_path)

            # Check if the 'Power BI Upload' sheet exists
            if 'Power BI Upload' in [sheet.name for sheet in source_workbook.sheets]:
                print(f"'Power BI Upload' sheet found in {filename}")

                # Get the 'Power BI Upload' sheet
                source_sheet = source_workbook.sheets['Power BI Upload']

                # Create a new sheet in the consolidated workbook
                new_sheet_name = f"Power BI Upload ({sheet_counter})"
                new_sheet = consol_workbook.sheets.add(name=new_sheet_name, before=consol_workbook.sheets[0])

                # Copy the data range from 'Power BI Upload' (you can adjust this range)
                # Assume a max range of A1:Z1000 if expand() doesn't capture all data
                source_data = source_sheet.range('A1:AH1000').value  # Adjust range to cover your data
                new_sheet.range('A1').value = source_data  # Paste as values

                print(f"Copied data from '{filename}' to Consolidated Workbook as '{new_sheet_name}'")

                sheet_counter += 1

            else:
                print(f"'Power BI Upload' sheet not found in {filename}")
                
        except Exception as e:
            print(f"Error processing {filename}: {e}")

# Workbooks remain open until manually closed
print("Workbooks remain open. You can manually close them.")
