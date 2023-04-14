# Filter out list that has only diabetes related medicine
import pandas as pd
import openpyxl

medicine_pre_final = openpyxl.load_workbook(filename="../intermediateResults/medicine_final_list.xlsx", read_only=True)
medicine_pre_final_worksheet = medicine_pre_final.active

patient_file = openpyxl.load_workbook(filename="../intermediateResults/patients_pre_final.xlsx", read_only=True)
patient_file_worksheet = patient_file.active
MAX_ROW = patient_file_worksheet.max_row
# MAX_ROW = 5

patients_array = []

# for i in range(1, MAX_ROW + 1):
#     # print(i)
#     cell_obj = patient_file_worksheet.cell(row = i, column = 1)
#     patients_array.append(cell_obj.value)

# print(patients_array)
medicineRows = medicine_pre_final_worksheet.rows
headers = [cell.value for cell in next(medicineRows)]
# print('headers', medicineRows)

patient_idx = headers.index('DUID')
quant_idx = headers.index('RXQUANTY')
med_name = headers.index('RXNAME')
days_up = headers.index('RXDAYSUP')
monthStartIdx = headers.index('RXBEGMM')

medicinesList = ['GLIPIZIDE', 'GLYBURIDE', 'GLICLAZIDE', 'GLIMEPIRIDE', 'TOLBUTAMIDE', 'REPAGLINIDE', 'NATEGLINIDE', 'METFORMIN', 'ROSIGLITAZONE', 'PIOGLITAZONE', 'ACARBOSE', 'MIGLITOL', 'VOGLIBOSE', 'SITAGLIPTIN', 'SAXAGLIPTIN', 'VILDAGLIPTIN', 'LINAGLIPTIN', 'ALOGLIPTIN', 'DAPAGLIFLOZIN', 'CANAGLIFLOZIN', 'EMPAGLIFLOZIN']
daysUP = [0 for i in range(len(medicinesList))] 
TotalMedsTobeTaken = [0 for i in range(len(medicinesList))] 
medPrescribed = [0 for i in range(len(medicinesList))]
medDosage = [0 for i in range(len(medicinesList))]
adherence = [0 for i in range(len(medicinesList))]

data = []
idx = 0

# for patient in patients_array:
  #  print(patient)

# for patient in patients_array:
#   print(patient)
patient = 2320543
record = {}
for row in medicineRows:
  idx += 1
  # print(type(row[patient_idx].value), type(patient))
  print(row[patient_idx].value, patient)
  # print(row[patient_idx].value == patient)
  if row[patient_idx].value == patient:
      medicineIndex = medicinesList.index(row[med_name].value)
      print('medicineIndex', row[days_up].value)
      daysUP[medicineIndex] += row[days_up].value
      # TotalMeds[medicineIndex] += row[quant_idx].value
      if medPrescribed[medicineIndex] != 0 or medPrescribed[medicineIndex] < row[quant_idx].value:
          TotalMedsTobeTaken[medicineIndex] = row[quant_idx].value * (13-row[monthStartIdx].value)
          medDosage[medicineIndex] = row[quant_idx].value / row[days_up].value
id = 0
for med_row in medicinesList:
  no_of_medicines = 0
  if daysUP[id] > 0:
      daysCovered = daysUP[id] * medDosage[id]
      medsTobeTaken = TotalMedsTobeTaken[id]
      adherenceCalculated = medsTobeTaken / daysCovered
      adherence[id] = medsTobeTaken / daysCovered
  id += 1

count = 0
overallAdherence = 0
for value in adherence:
  if value > 0:
    count += 1
    overallAdherence += value
    print(count, overallAdherence)
if count > 0:
  overallAdherence = overallAdherence / count
  record['PATIENT'] = patient
  record['ADHERENCE'] = overallAdherence
print('in herrhr')
data.append(record)

medicine_pre_final.close()
patient_file.close()

# Convert to a df
df = pd.DataFrame(data)
print(df)
df.to_excel("../intermediateResults/finalResult.xlsx")
