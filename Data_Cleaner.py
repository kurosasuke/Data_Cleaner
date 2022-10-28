import numpy as np
import openpyxl
import matplotlib.pyplot as plt

# Get Excel data
worksheet = openpyxl.load_workbook(filename=r"C:\Users\darii\OneDrive\Desktop\FC_data.xlsx")

sheet = worksheet.active
col_T = "C"
col_V = "D"
col_C = "E"

Time = np.zeros(len(sheet[col_T]))
Volt = np.zeros(len(sheet[col_V]))
Curr = np.zeros(len(sheet[col_C]))

# Import Seconds from Excel
for i in range(len(sheet[col_T])):
    if isinstance(sheet.cell(row=i + 1, column=3).value, int):
        val_t = sheet.cell(row=i + 1, column=3).value
        Time[i] = val_t

# Import Voltage from Excel
for i in range(len(sheet[col_V])):
    if isinstance(sheet.cell(row=i + 1, column=4).value, float) or \
            isinstance(sheet.cell(row=i + 1, column=4).value, int):
        val_v = sheet.cell(row=i + 1, column=4).value
        Volt[i] = val_v

# Import Current from Excel
for i in range(len(sheet[col_C])):
    if isinstance(sheet.cell(row=i + 1, column=5).value, float) or \
            isinstance(sheet.cell(row=i + 1, column=5).value, int):
        val_c = sheet.cell(row=i + 1, column=5).value
        Curr[i] = val_c

# Ramping
for i in range(len(Time)):
    if Time[i] == 0 and Time[i + 1] == 0:
        new_Time = np.array_split(Time, 3)

# Plot results
plt.scatter(Volt, Curr)

# Trend-line
z = np.polyfit(Volt, Curr, 7)
p = np.poly1d(z)

# Add trend-line to scatter plot
plt.plot(Volt, p(Volt), color="purple", linewidth=1, linestyle="--")

# Show plot on Pycharm (maybe you doesn't need it)
plt.show()
