from cgi import test
import openpyxl
import datetime
import pandas as pd
import numpy as np
import math

# THINGS TO DO:

# Functional:
# Make the whole writing to the new spreadsheet segment of code. -------------
# Integrate all of test.py into Ollie.py

# Testing:
# Validate zeros testing. Check if null, zeros, etc. works always ------------

# Efficiency:
# Change all arrays to numpy arrays or pandas or something

def test_func():
# Path to the Excel files, hardcoded for now
    file_path = "/Users/lukediehm/Downloads/Dad Project/Test/Q-Vendor Choice Report copy.xlsx"

    file_path2 = "/Users/lukediehm/Downloads/Dad Project/Test/Jonesboro Forecast.xlsx"

    data_CR = pd.read_excel(file_path)
    data_JF = pd.read_excel(file_path2)

# Here we are constructing the data columns into CR and JF. Part number not included in this big array of arrays.
   
    # CR_columns_to_transfer = [1,3,4,5,7,9]
    column_headers = data_CR.keys().tolist()
    CR_columns_to_transfer = [column_headers.index("ICPROD"), column_headers.index("Total On Hand"), column_headers.index("On Order"), column_headers.index("Consigned"), column_headers.index("Required"), column_headers.index("SumOfWIP")]
    Headers_CR = data_CR.columns[CR_columns_to_transfer] 
    
    column_headers = data_JF.keys().tolist()
    JF_columns_to_transfer = [column_headers.index("Part Number"), column_headers.index("CURRENT Inventory"), column_headers.index("CURR. Orders"), column_headers.index("KBM Cons."), column_headers.index("CURRENT Demand"), column_headers.index("WIP")]
    Headers_JF = data_JF.columns[JF_columns_to_transfer] # part number, current inventory, current orders, KBM const, current demand, WIP
    
    # Here we are making a list of lists of all the part data. Part numbers NOT included for this one. Part numbers are stored in part_numbers_existing_in_JF and part_number_ChoiceReport
    CR = np.array([data_CR[Headers_CR[1]].tolist(), data_CR[Headers_CR[2]].tolist(), data_CR[Headers_CR[3]].tolist(), data_CR[Headers_CR[4]].tolist(), data_CR[Headers_CR[5]].tolist()])
    JF = np.array([data_JF[Headers_JF[1]].tolist(), data_JF[Headers_JF[2]].tolist(), data_JF[Headers_JF[3]].tolist(), data_JF[Headers_JF[4]].tolist(), data_JF[Headers_JF[5]].tolist()])

    ## This is making the part numbers array from Jonesboro Forecast, ensuring no whitespace.
    part_numbers_existing_in_JF = []
    for cell in data_JF[Headers_JF[0]].tolist():
        if not pd.isnull(cell):
            cell = cell.strip()
        part_numbers_existing_in_JF.append(cell)


    #Now we begin the iterating process through the whole excel sheet. Each part will be read from the Choice Report and searched for in the Jonesboro Forecast.
    parts_changed = []
    parts_not_found = []
    is_changed = False
    test = 0
    #We dont do the same part number array building in the Choice Report, so I strip each part number in the for loop as they're read to ensure no white space.
    part_number_ChoiceReport = data_CR[Headers_CR[0]].tolist()

    for ind, part_number in enumerate(part_number_ChoiceReport):
        part_number = part_number.strip()
        test += 1
        if part_number in part_numbers_existing_in_JF:
            is_changed = False
            temp = []
            part_number_ind = part_numbers_existing_in_JF.index(part_number)
            temp.append((part_number, part_number_ind+1))
            # if test < 5:
            #     print(part_number)
            for i in range(5):
                cell_CR_null = CR[i][ind]
                cell_JF_null = JF[i][ind]

                try:
                    cell_CR = int(cell_CR_null)
                except:
                    cell_CR = 0

                try:
                    cell_JF = int(float(cell_JF_null))
                except:
                    print("JF at " + str(part_number_ind) + "was null\n")
                    cell_JF = 0

                if cell_CR != cell_JF:
                    is_changed = True
                    temp.append((i, cell_JF, cell_CR))
                    


            if is_changed:
                parts_changed.append(temp)

        else: # Part not found! Important to acknowledge.
            parts_not_found.append((part_number, ind))


    print("\n\nParts not found:\n")
    print(parts_not_found)
    print("\n\n\n\n\n\n\n\n")
    print(parts_changed)
    return parts_changed, parts_not_found, JF_columns_to_transfer


def write_to_Jonesboro_Forecast(parts_changed, JF_columns_to_transfer):
    JF_columns_to_transfer.pop(0)
    print("Hi")

    # print(JF_column_headers.pop(0))
    workbook = openpyxl.load_workbook("/Users/lukediehm/Downloads/Dad Project/Test/Jonesboro Forecast.xlsx")
    sheet = workbook.active
    print("\n\n\n")
    print(parts_changed)
    test = 0
    for part in parts_changed:
        test += 1
        part_number = part.pop(0)

        # print("Part number is " + str(part_number))
        # print(part)

        for change_in_quantity in part:
            Row = part_number[1] + 1
            Col = JF_columns_to_transfer[change_in_quantity[0]] + 1
            cell = sheet.cell(row = Row, column = Col)

            # print(cell.value)
            cell.value = change_in_quantity[2]
            # print(cell.value)

    workbook.save("done.xlsx")
    workbook.close()

if __name__ == "__main__":

    parts_changed, parts_not_found, JF_columns_to_transfer = test_func()

    write_to_Jonesboro_Forecast(parts_changed, JF_columns_to_transfer)



    # CR_columns_to_transfer = [1,3,4,5,7,9]
    # JF_columns_to_transfer = [0,4,130,134,3,131] 