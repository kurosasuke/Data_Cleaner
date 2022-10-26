import openpyxl
import numpy
import matplotlib.pyplot as plt

# Get Excel data
worksheet = openpyxl.load_workbook(filename=r"C:\Users\darii\OneDrive\Desktop\FC_data.xlsx")

workbook = openpyxl.Workbook()
newworksheet = workbook.active

sheet = worksheet.active

# Start column (Voltage)
col = "D"

# initialization
i = 1
xxx = numpy.zeros(len(sheet[col]) + 1)
yyy = numpy.zeros(len(sheet[col]) + 1)

# Delete short circuit values
for x in range(len(sheet[col])):
    s = x + 1
    if sheet.cell(row=s, column=4).value == 0 and sheet.cell(row=s - 1, column=4).value != 0 and \
            sheet.cell(row=s + 1, column=4).value != 0:
        print("")
    else:
        newworksheet.cell(row=i, column=1).value = sheet.cell(row=s, column=4).value
        newworksheet.cell(row=i, column=2).value = sheet.cell(row=s, column=5).value
        if type(newworksheet.cell(row=i, column=1).value) != str:
            xxx[s] = sheet.cell(row=s, column=4).value
            yyy[s] = sheet.cell(row=s, column=5).value
        i += 1

# Save right values into new file
workbook.save("FC_Data_Cleaned.xlsx")

# Scatter plot results
plt.scatter(xxx, yyy)
plt.show()
