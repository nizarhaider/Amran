from openpyxl import load_workbook

# Load Excel file
workbook = load_workbook(filename="sheet.xlsx")
worksheet = workbook.active

# Get the range of cells to read from
column_i = worksheet['I']
start_row = 9  # Row index starts at 1, so start at 9 instead of 8
end_row = len(list(filter(lambda x: x.value is not None, column_i)))

for row_index in range(start_row, end_row+1):
    # Get the cell value and remove commas
    cell_value = str(column_i[row_index-1].value).replace(",", "")
    print(cell_value)
    # input_field = driver.find_element_by_name("input_field_name")
    # input_field.clear()
    # input_field.send_keys(cell_value)
    # input_field.send_keys(Keys.RETURN)