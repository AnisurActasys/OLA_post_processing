from pathlib import Path
from openpyxl import Workbook, load_workbook
from pyparsing import col
import matplotlib.pyplot as plt
import numpy as np

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
clean_speed_dict = {}
cutoff = 80

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
        if clean_percent == cutoff:
            time_to_80 = time_key - reference_time
            energy_to_80 = (energy / time_key) * time_to_80 / 1000
            break
        elif clean_percent > cutoff:
            time_to_80 = (time_key - time_keys[j-1]) / (clean_percent - cleaning_data[case][time_keys[j-1]][0]) * (cutoff - cleaning_data[case][time_keys[j-1]][0]) + time_keys[j-1] - reference_time
            energy_to_80 = energy / time_key * time_to_80 / 1000
            break
    if time_to_80 != None:
        time_and_energy["Time"] = time_to_80 
        time_and_energy["Energy"] = energy_to_80
        to_80[case] = time_and_energy
        energy_dict[case] = energy_to_80
        time_dict[case] = time_to_80
        clean_speed_dict[case] = cutoff / time_to_80


inputs = list(energy_dict.keys())
energies = list(energy_dict.values())
times = list(time_dict.values())
cleaning_speeds = list(clean_speed_dict.values())


def OLA_plots(inputs, var1, var2, axis_labels, legend_labels, title, figure_number):
    
    fig, ax = plt.subplots(figsize = (16,6))
    labels = inputs
    x = np.arange(len(inputs))
    ax2 = ax.twinx()

    ax.set_xlabel("Input Parameters")
    ax.set_ylabel(axis_labels[0])
    ax2.set_ylabel(axis_labels[1])

    color = ['red', 'royalblue']
    width = 0.25

    p1 = ax.bar(x-width, var1, width = width, color = color[0], align = 'edge', label = legend_labels[0])
    p2 = ax2.bar(x, var2, width = width, color = color[1], align = 'edge', label = legend_labels[1])

    lns = [p1, p2]
    ax.legend(handles = lns, loc = 'best')
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_title(title,fontsize=18, weight='bold')
    return plt.figure(figure_number)

p1 = OLA_plots(inputs, cleaning_speeds, times, ["Clean Speed (%/s)", "Cleaning Time (s)"], ["Clean Speed", "Time"], "Cleaning Speeds and Cleaning Times", 1)
p2 = OLA_plots(inputs, energies, times, ["Energy Consumed (kJ)", "Cleaning Times (s)"], ["Energy", "Time"], "Energy Consumption and Cleaning Times", 2)
p3 = OLA_plots(inputs, energies, cleaning_speeds, ["Energy Consumed (kJ)", "Clean Speeds (%/s)"], ["Energy", "Clean Speed"], "Energy Consumption and Cleaning Speeds", 3)


plt.show()



