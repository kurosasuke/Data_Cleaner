import numpy as np
import openpyxl
import matplotlib.pyplot as plt

# Get Excel data
worksheet = openpyxl.load_workbook(filename=r"C:\Users\darii\OneDrive\Desktop\FC_data.xlsx")

sheet = worksheet.active
col_T = "C"
col_V = "D"
col_C = "E"

Time = []
Volt = []
Curr = []

# Import Seconds from Excel
for i in range(len(sheet[col_T])):
    if isinstance(sheet.cell(row=i + 1, column=3).value, int):
        val_t = sheet.cell(row=i + 1, column=3).value
        Time.append(val_t)

# Import Voltage from Excel
for i in range(len(sheet[col_V])):
    if isinstance(sheet.cell(row=i + 1, column=4).value, float) or \
            isinstance(sheet.cell(row=i + 1, column=4).value, int):
        val_v = sheet.cell(row=i + 1, column=4).value
        Volt.append(val_v)

# Import Current from Excel
for i in range(len(sheet[col_C])):
    if isinstance(sheet.cell(row=i + 1, column=5).value, float) or \
            isinstance(sheet.cell(row=i + 1, column=5).value, int):
        val_c = sheet.cell(row=i + 1, column=5).value
        Curr.append(val_c)

# Time start from zero
Time = Time[250:]
Volt = Volt[250:]
Curr = Curr[250:]

# Clean data from zeros
Volt_n = []
Curr_n = []

for i in range(len(Volt)):
    if Volt[i] != 0:
        Volt_n.append(Volt[i])
        Curr_n.append(Curr[i])

# Ramping variables
new_Time = []
fake_time = []

# Distance between two ramping sets
for i in range(len(Time)):
    if Time[i - 1] != 0 and Time[i] == 0 and Time[i + 1] == 0:
        new_Time.append(fake_time)
        fake_time = []
    else:
        fake_time.append(Time[i])

# Delete I<40A
new_Volt = []
new_Curr = []

for i in range(len(Volt_n)):
    if Volt_n[i] > 40:
        new_Volt.append(Volt_n[i])
        new_Curr.append(Curr_n[i])

# Plot results
plt.scatter(new_Volt, new_Curr)

# Trend-line
z = np.polyfit(new_Volt, new_Curr, 7)
p = np.poly1d(z)

# Add trend-line to scatter plot
plt.plot(new_Volt, p(new_Volt), color="purple", linewidth=1, linestyle="--")

# Show plot on Pycharm (maybe you doesn't need it)
plt.show()
