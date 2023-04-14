# Removing patients who are institutionalised
import pandas as pd
import openpyxl
from openpyxl.worksheet.filters import DateGroupItem

wb = openpyxl.load_workbook(filename="../intermediateResults/filtered_died.xlsx", read_only=True)
ws = wb.active
# Load the rows
rows = ws.rows
headers = [cell.value for cell in next(rows)]

# column for age at which patient diagnosed with diabetes - DIABAGY1
INST_COL = headers.index('INST')

data = []
for row in rows:
    record = {}
    has_diabetes = False
    diagnosis_legal_age = False
    # INST 1 YES

    if row[INST_COL].value == 0:
        for key, cell in zip(headers, row):
          record[key] = cell.value
        data.append(record)


# Convert to a df
df = pd.DataFrame(data)

print(df)
# Save to file
df.to_excel("../intermediateResults/filtered_students.xlsx")

wb.close()