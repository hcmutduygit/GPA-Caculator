from openpyxl import load_workbook

# Existing Excel file
existing_file = 'hello.xlsx'

# New data to append
new_data = [['Bob', 28, 55000]]

# Load existing workbook
wb = load_workbook(existing_file)

# Select the active sheet
ws = wb.active

# Append new data
for row in new_data:
	ws.append(row)

# Save the workbook
wb.save(existing_file)

