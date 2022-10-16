import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl import Workbook


def get_compiled_data(data_specifier):
    battery_info = {
        'filenames': [],
        'data': []
    }

    if data_specifier == 'capacity':
        individual_data_wb = Workbook()
        individual_data_wb.remove(individual_data_wb.active)
        for i in range(1, 8):
            individual_data_wb.create_sheet('Box ' + str(i))
            individual_data_wb['Box ' + str(i)]['A1'].value = 'Cell'
            individual_data_wb['Box ' + str(i)]['B1'].value = 'OCV'
            individual_data_wb['Box ' + str(i)]['C1'].value = 'Capacity'
            individual_data_wb['Box ' + str(i)]['D1'].value = 'DCIR 1 (Idle-P1)'
            individual_data_wb['Box ' + str(i)]['E1'].value = 'DCIR 2 (P1-P2)'

        cell_tracker_file = 'Battery Cell Tracker_Capacity_OCV.xlsx'

        cap_wb = pd.read_excel(io='CompiledDischargeFile.xlsx', sheet_name='8 Discharge')

        cap_data = cap_wb.to_numpy().transpose()

        battery_info['filenames'] = cap_data[0]
        battery_info['data'] = cap_data[1]
    else:
        individual_data_wb = load_workbook('Individual_Data.xlsx')
        cell_tracker_file = 'Battery Cell Tracker_DCIR.xlsx'

        battery_info['filenames'], battery_info['data'] = get_dcir_data()

    # Initialize cell tracker workbook
    tracker_wb = load_workbook(cell_tracker_file)

    # Iterate through the compiled DCIR/capacity data and
    for i in range(len(battery_info['filenames'])):
        filename_parts = battery_info['filenames'][i].split('_')

        # If the box number in the filename is formatted as "box_#" instead of "box#", add the "#" element of
        # filename_parts (element 3) to the "box" element (making it "box#") and remove the "#" element
        if (len(filename_parts[2]) < 4):
            filename_parts[2] += filename_parts[3]
            del filename_parts[3]

        box = filename_parts[2][3:]
        group = int(filename_parts[3][5:])
        row = int(filename_parts[5])

        t_ws = tracker_wb['Box ' + box]
        id_ws = individual_data_wb['Box ' + box]

        bat_id = t_ws.cell(row=(row + 2), column=(2 * group)).value

        if data_specifier == 'capacity':
            next_row = len(id_ws['A'])

            id_ws.cell(row=(next_row + 1), column=1).value = bat_id
            id_ws.cell(row=(next_row + 1), column=2).value = t_ws.cell(row=(row + 2), column=(2 * group + 1)).value
            id_ws.cell(row=(next_row + 1), column=3).value = battery_info['data'][i]
        else:
            for j in range(1, len(id_ws['A']) + 1):
                if id_ws['A' + str(j)].value == bat_id:
                    id_ws.cell(row=j, column=4).value = battery_info['data'][0][i]
                    id_ws.cell(row=j, column=5).value = battery_info['data'][1][i]
                    break

    individual_data_wb.save('Individual_Data.xlsx')


def get_dcir_data():
    wb_idle = pd.read_excel(io='CompiledDCIR.xlsx', sheet_name='1 Idle') # Idle before pulse 1
    wb_ccd1 = pd.read_excel(io='CompiledDCIR.xlsx', sheet_name='2 CCD') # Pulse 1
    wb_ccd2 = pd.read_excel(io='CompiledDCIR.xlsx', sheet_name='3 CCD') # Pulse 2

    headers = np.array(list(wb_idle.columns.values))
    headers = headers[~np.char.startswith(headers, 'Unnamed')]

    raw_data_idle = wb_idle.to_numpy().transpose()
    raw_data_pulse1 = wb_ccd1.to_numpy().transpose()
    raw_data_pulse2 = wb_ccd2.to_numpy().transpose()

    data_idle = []
    data_pulse1 = []
    data_pulse2 = []

    # Get data from Idle, CCD (pulse 1), and CCD (pulse 2). Only add data for all 3 stages if CCD (pulse 1) has data
    for i in range(len(raw_data_idle)):
        tmp2 = raw_data_pulse1[i]
        tmp2 = tmp2[~pd.isna(tmp2)][1:]

        if len(tmp2) != 0 and (raw_data_pulse1[i][0] == 'V' or raw_data_pulse1[i][0] == 'mA'):
            tmp = raw_data_idle[i]
            tmp = tmp[~pd.isna(tmp)][-1]
            tmp3 = raw_data_pulse2[i]
            tmp3 = tmp3[~pd.isna(tmp3)][1]

            data_idle.append(tmp)
            data_pulse1.append(tmp2)
            data_pulse2.append(tmp3)

    dcir1 = []
    dcir2 = []

    for i in range(int(len(data_idle) / 2)):
        # Multiply both quotients by -1000: 1000 to convert from kOhms to Ohms, -1 to account for negative signs on
        #                                   current being dropped when parsing the txt files
        dcir1.append(-1000 * (data_pulse1[i * 2][0] - data_idle[i * 2]) / (data_pulse1[i * 2 + 1][0] - data_idle[i * 2 + 1]))
        dcir2.append(-1000 * (data_pulse2[i * 2] - data_pulse1[i * 2][-1]) / (data_pulse2[i * 2 + 1] - data_pulse1[i * 2 + 1][-1]))

    return headers, [dcir1, dcir2]


# Main function
if __name__ == '__main__':
    get_compiled_data('capacity')
    get_compiled_data('dcir')
