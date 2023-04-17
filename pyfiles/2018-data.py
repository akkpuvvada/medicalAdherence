import pandas as pd
import openpyxl
from openpyxl.worksheet.filters import DateGroupItem

wb = openpyxl.load_workbook(filename="../originalDataFiles/h217.xlsx", read_only=True)
ws = wb.active
# Load the rows
rows = ws.rows
headers = [cell.value for cell in next(rows)]

# column for age at which patient diagnosed with diabetes - DIABAGY1
DIABAGY1_col = headers.index('DIABAGY1')

# column at which this diagnosis is mentioned - DIABDXY1_M18
DIABDXY1_M18_col = headers.index('DIABDXY1_M18')
DIABDXY2_M18_col = headers.index('DIABDXY2_M18')

data = []
idx = 0
for row in rows:
    record = {}
    has_diabetes = False
    diagnosis_legal_age = False
    print(idx)
    # DIABAGY1 -> 0 - 85 AGE AT DIAGNOSIS
    # DIABDXY1_M18 -> DIABETES DIAGNOSIS 18 filled value = 1

    if row[DIABAGY1_col].value > 0 and row[DIABDXY1_M18_col].value == 1:
        for key, cell in zip(headers, row):
          record[key] = cell.value
        data.append(record)
    idx += 1

# Convert to a df
df = pd.DataFrame(data)

print(df)
# Save to file
df.to_excel("../results2/filter_output1.xlsx")

wb.close()
