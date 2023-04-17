# Filter out list that has only diabetes related medicine
import pandas as pd
import openpyxl
from openpyxl.worksheet.filters import DateGroupItem

wb = openpyxl.load_workbook(filename="../originalDataFiles/h206a.xlsx", read_only=True)
ws = wb.active

# Load the rows
rows = ws.rows
headers = [cell.value for cell in next(rows)]

# column for name of medicine prescribed to patient
NAME_COL = headers.index('RXNAME')

medicinesList = ['GLIPIZIDE', 'GLYBURIDE', 'GLICLAZIDE', 'GLIMEPIRIDE', 'TOLBUTAMIDE', 'REPAGLINIDE', 'NATEGLINIDE', 'METFORMIN', 'ROSIGLITAZONE', 'PIOGLITAZONE', 'ACARBOSE', 'MIGLITOL', 'VOGLIBOSE', 'SITAGLIPTIN', 'SAXAGLIPTIN', 'VILDAGLIPTIN', 'LINAGLIPTIN', 'ALOGLIPTIN', 'DAPAGLIFLOZIN', 'CANAGLIFLOZIN', 'EMPAGLIFLOZIN']

data = []
for row in rows:
    record = {}
    if row[NAME_COL].value in medicinesList:
        for key, cell in zip(headers, row):
          record[key] = cell.value
        data.append(record)


# Convert to a df
df = pd.DataFrame(data)

print(df)
# Save to file
df.to_excel("../results2/filter_medicines1.xlsx")

wb.close()
