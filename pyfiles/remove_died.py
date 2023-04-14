import pandas as pd
import openpyxl
from openpyxl.worksheet.filters import DateGroupItem

wb = openpyxl.load_workbook(filename="../intermediateResults/output.xlsx", read_only=True)
ws = wb.active
# Load the rows
rows = ws.rows
headers = [cell.value for cell in next(rows)]

# column for age at which patient diagnosed with diabetes - DIABAGY1
DIED_COL = headers.index('DIED')

data = []
for row in rows:
    record = {}
    has_diabetes = False
    diagnosis_legal_age = False
    # DIED 1 YES

    if row[DIED_COL].value == 0:
        for key, cell in zip(headers, row):
          record[key] = cell.value
        data.append(record)


# Convert to a df
df = pd.DataFrame(data)

print(df)
# Save to file
df.to_excel("filtered_died.xlsx")

wb.close()