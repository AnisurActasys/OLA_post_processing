from csv import excel
from pathlib import Path
from numpy import true_divide
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pyparsing import col

excel_path = Path.cwd() / "cleaning_effectiveness_results.xlsx"
wb = load_workbook(excel_path, data_only=True)
ws = wb.active
max_row = ws.max_row
max_col = ws.max_column

#parameters = ["100Hz 80V 8.5A", "100Hz 60V 6.3A", "80Hz 80V 6.9A", "80Hz 60V 5.9A", ]
cleaning_data = {}

for j in range(1, max_col, 3):
    table_values = {}
    parameter = ws.cell(row = 1, column = j).value
    table_values[ws.cell(row = 3, column = j).value] = [0, 0, 0]
    i = 4
    while(True):
        if ws.cell(row = i, column = j).value == None:
            break
        else:
            derivative = (ws.cell(row = i, column = j+1).value - ws.cell(row = i-1, column = j+1).value) / ((ws.cell(row = i, column = j).value - ws.cell(row = i-1, column = j).value))
            table_values[ws.cell(row = i, column = j).value] = [ws.cell(row = i, column = j+1).value, ws.cell(row = i, column = j+2).value, derivative]
            i += 1
    cleaning_data[parameter] = table_values



cases = list(cleaning_data.keys())
#print(cleaning_data[cases[0]][time_keys[1]][0])
to_80 = {}
#derivative_dict = {}
#derivative_vals = []
#for case in cases:
#    time_keys = list(cleaning_data[case].keys())
#    for time in time_keys:
#        derivative_vals.append(cleaning_data[case][time][2])
#    derivative_dict[case] = derivat

for case in cases:
    time_keys = list(cleaning_data[case].keys())
    #derivtives = list(cleaning_data[case].values()[2])
    time_and_energy = {}
    for j in range(len(time_keys)):
        time_key = time_keys[j]
        clean_percent = cleaning_data[case][time_key][0]
        energy = cleaning_data[case][time_key][1]
        #derivative = cleaning_data[case][time_key][2]
        time_to_80 = None
        energy_to_80 = None
        #if derivative == max(derivtives):
            #reference_time = time_key
        if clean_percent == 80:
            time_to_80 = time_key
            energy_to_80 = energy
            break
        elif clean_percent > 80:
            time_to_80 = (time_key - time_keys[j-1]) / (clean_percent - cleaning_data[case][time_keys[j-1]][0]) * (80 - cleaning_data[case][time_keys[j-1]][0]) + time_keys[j-1]
            energy_to_80 = (energy - cleaning_data[case][time_keys[j-1]][1]) / (time_key - time_keys[j-1]) * (time_to_80 - time_keys[j-1]) + cleaning_data[case][time_keys[j-1]][1]
            break
    time_and_energy["Time"] = time_to_80 
    time_and_energy["Energy"] = energy_to_80
    to_80[case] = time_and_energy

print(to_80)




