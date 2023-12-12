import pyperclip
import win32com.client


# Assume you have data in a string, for example:
with open(r'C:\\Users\\Joshua\\OneDrive\Documents\\Coding Projects\\Rapier\\Pick Place for Millbrite-T_Panel(Production).txt', 'r') as f1:
    data_to_copy = f1.read()

with open(r'C:\\Users\\Joshua\\OneDrive\Documents\\Coding Projects\\Rapier\\micropicsmt4v0 Bill of Material 27-7-2022 Manuf .txt', 'r') as f2:
    data_to_copy2 = f2.read()

# Copy the data to the clipboard
pyperclip.copy(data_to_copy2)

# Create a new Excel application
excel = win32com.client.Dispatch("Excel.Application")

# Make Excel visible (optional)
excel.Visible = True

# Create a new workbook
workbook = excel.Workbooks.Add()

# Select the active worksheet
sheet = workbook.ActiveSheet

# Paste the clipboard content into the selected cell
sheet.Paste()


# Select all cells in the worksheet and copy the data
sheet.UsedRange.Copy()

# Specify the output text file
output_text_file = 'output.txt'

# Open the text file for writing
# Open the text file for writing
with open(output_text_file, 'w') as text_file:
    # Loop through rows in the active sheet
    for row in range(1, sheet.UsedRange.Rows.Count + 1):
        # Access the values in the tuple and join them with tabs
        row_data = '\t'.join([str(cell) for cell in sheet.Range(f"A{row}", f"Z{row}").Value[0]])
        text_file.write(row_data + '\n')

# Close the Excel workbook
workbook.close()

# Close Excel (optional)
excel.Quit()