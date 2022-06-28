from csv import excel
from pathlib import Path
from numpy import true_divide
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pyparsing import col
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

excel_path = Path.cwd() / "cleaning_effectiveness_results.xlsx"
wb = load_workbook(excel_path, data_only=True)
ws = wb.active
max_row = ws.max_row
max_col = ws.max_column

#parameters = ["100Hz 80V 8.5A", "100Hz 60V 6.3A", "80Hz 80V 6.9A", "80Hz 60V 5.9A", ]
cleaning_data = {}
derivative_dict = {}
for j in range(1, max_col, 3):
    table_values = {}
    derivative_vals = []
    parameter = ws.cell(row = 1, column = j).value
    table_values[ws.cell(row = 3, column = j).value] = [0, 0, 0]
    i = 4
    while(True):
        if ws.cell(row = i, column = j).value == None:
            break
        else:
            derivative = (ws.cell(row = i, column = j+1).value - ws.cell(row = i-1, column = j+1).value) / ((ws.cell(row = i, column = j).value - ws.cell(row = i-1, column = j).value))
            derivative_vals.append(derivative)
            table_values[ws.cell(row = i, column = j).value] = [ws.cell(row = i, column = j+1).value, ws.cell(row = i, column = j+2).value, derivative]
            i += 1
    cleaning_data[parameter] = table_values
    derivative_dict[parameter] = derivative_vals



cases = list(cleaning_data.keys())

to_80 = {}
energy_dict = {}
time_dict = {}

for case in cases:
    time_keys = list(cleaning_data[case].keys())
    derivatives = list(derivative_dict[case])
    time_and_energy = {}
    for j in range(len(time_keys)):
        time_key = time_keys[j]
        clean_percent = cleaning_data[case][time_key][0]
        energy = cleaning_data[case][time_key][1]
        derivative = cleaning_data[case][time_key][2]
        time_to_80 = None
        energy_to_80 = None
        if derivative == max(derivatives):
            reference_time = time_key
        if clean_percent == 80:
            time_to_80 = time_key - reference_time
            energy_to_80 = (energy / time_key) * time_to_80 / 1000
            break
        elif clean_percent > 80:
            time_to_80 = (time_key - time_keys[j-1]) / (clean_percent - cleaning_data[case][time_keys[j-1]][0]) * (80 - cleaning_data[case][time_keys[j-1]][0]) + time_keys[j-1] - reference_time
            energy_to_80 = energy / time_key * time_to_80 / 1000
            break
    if time_to_80 != None:
        time_and_energy["Time"] = time_to_80 
        time_and_energy["Energy"] = energy_to_80
        to_80[case] = time_and_energy
        energy_dict[case] = energy_to_80
        time_dict[case] = time_to_80


inputs = list(energy_dict.keys())
energies = list(energy_dict.values())
times = list(time_dict.values())

#fig = plt.figure(figsize= (10, 5))
#plt.bar(inputs, energies)
#plt.xlabel("Input Parameters")
#plt.ylabel("Energy Consumed (kJ)")
#plt.title("Total Energy Consumed to Reach 80% Cleaning Effectiveness")
#plt.show()
#
#fig = plt.figure(figsize= (10, 5))
#inputs = list(energy_dict.keys())
#times = list(time_dict.values())
#plt.bar(inputs, energies)
#plt.xlabel("Input Parameters")
#plt.ylabel("Energy Consumed (kJ)")
#plt.title("Total Energy Consumed to Reach 80% Cleaning Effectiveness")
#plt.show()


fig, ax = plt.subplots(figsize = (16,6))
labels = inputs
x = np.arange(len(inputs))
ax2 = ax.twinx()

ax.set_xlabel("Input Parameters")
ax.set_ylabel("Energy Consumed (kJ)")
ax2.set_ylabel("Time to 80% Cleaning (s)")

color = ['red', 'royalblue']
width = 0.25

p1 = ax.bar(x-width, energies, width = width, color = color[0], align = 'edge', label = 'Energy')
p2 = ax2.bar(x, times, width = width, color = color[1], align = 'edge', label = 'Time')

lns = [p1, p2]
ax.legend(handles = lns, loc = 'best')
ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.set_title("Time and Energy Demands to Reach 80% CLeaning Effectiveness",fontsize=18, weight='bold')
plt.show()
