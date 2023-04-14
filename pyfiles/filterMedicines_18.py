# Filter out list that has only diabetes related medicine
import pandas as pd
import openpyxl

wb = openpyxl.load_workbook(filename="filter_medicine_list.xlsx", read_only=True)
ws = wb.active

# Load the rows
rows = ws.rows
headers = [cell.value for cell in next(rows)]

# column for name of medicine prescribed to patient
YEAR_COL = headers.index('RXBEGYRX')

data = []
for row in rows:
    record = {}
    if row[YEAR_COL].value == 2018:
        for key, cell in zip(headers, row):
          record[key] = cell.value
        data.append(record)

# Convert to a df
df = pd.DataFrame(data)

# Save to file
df.to_excel("filter_medicine_2018.xlsx")

wb.close()
