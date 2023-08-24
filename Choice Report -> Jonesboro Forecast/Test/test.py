from cgi import test
import openpyxl
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import font


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
   
    CR_column_names = ["ICPROD","Total On Hand","On Order","Consigned","Required","SumOfWIP"]
    column_headers = data_CR.keys().tolist()
    CR_columns_to_transfer = [column_headers.index(CR_column_names[0]), column_headers.index(CR_column_names[1]), column_headers.index(CR_column_names[2]), column_headers.index(CR_column_names[3]), column_headers.index(CR_column_names[4]), column_headers.index(CR_column_names[5])]
    Headers_CR = data_CR.columns[CR_columns_to_transfer] 
    
    JF_column_names = ["Part Number","CURRENT Inventory","CURR. Orders","KBM Cons.","CURRENT Demand", "WIP"]
    column_headers = data_JF.keys().tolist()
    JF_columns_to_transfer = [column_headers.index(JF_column_names[0]), column_headers.index(JF_column_names[1]), column_headers.index(JF_column_names[2]), column_headers.index(JF_column_names[3]), column_headers.index(JF_column_names[4]), column_headers.index(JF_column_names[5])]
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
                cell_JF_null = JF[i][part_number_ind]

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
    # print(parts_changed)
    return parts_changed, parts_not_found, JF_columns_to_transfer, JF_column_names


def write_to_Jonesboro_Forecast(parts_changed, JF_columns_to_transfer):
    JF_columns_to_transfer.pop(0)
    print("Hi")

    # print(JF_column_headers.pop(0))
    workbook = openpyxl.load_workbook("/Users/lukediehm/Downloads/Dad Project/Test/Jonesboro Forecast.xlsx")
    sheet = workbook.active
    print("\n\n\n")

    test = 0
    for part in parts_changed:
        test += 1
        part_number = part[0]

        print("Part number is " + str(part_number))
        print(part)

        for change_in_quantity in part[1:]:
            print(change_in_quantity)
            Row = part_number[1] + 1
            Col = JF_columns_to_transfer[change_in_quantity[0]] + 1
            cell = sheet.cell(row = Row, column = Col)

            # print(cell.value)
            cell.value = change_in_quantity[2]
            # print(cell.value)

    workbook.save("done.xlsx")
    workbook.close()

def print_output(parts_changed, parts_not_found, JF_column_names):


    built_string = "Parts changed:\n"
    # Adding text based on changed values
    for changed_part in parts_changed:
        part_num = changed_part.pop(0)
        
        built_string += "\n" + str(part_num[0]) + ": "
        number_for_spacing_output = 0
        for change in changed_part:
            built_string += str(JF_column_names[change[0] + 1]) + " changed from " + str(change[1]) + " to " + str(change[2]) + ". "
            number_for_spacing_output += 1
            if number_for_spacing_output == 2:
                built_string += "\n"
        if number_for_spacing_output != 2:
            built_string += "\n"
    
    built_string += "\n\nParts not found:\n"

    for part in parts_not_found:
        built_string += f"\nDid not found part {part[0]} in the Jonesboro Forecast. \nThis part was seen at position {str(part[1])} in the Choice Report.\n"
    # Create the main window
    root = tk.Tk()

    # Set the window title
    root.title("Results from data transfer")
    custom_font = font.Font(family="Arial", size=18)
    
    root.configure(bg="#ECECEC")

    # Create a Text widget for displaying text
    text_widget = tk.Text(root, font = custom_font)
    text_widget.pack()
    text_widget.insert(tk.END, built_string)

    # text_widget.insert(tk.END, "Neener!\n")

    # Create a Close button to close the window
    close_button = tk.Button(root, text="Close", command=root.destroy)
    close_button.pack()

    # Run the main event loop for the results 
    root.mainloop()

if __name__ == "__main__":

    parts_changed, parts_not_found, JF_columns_to_transfer, JF_column_names = test_func()

    write_to_Jonesboro_Forecast(parts_changed, JF_columns_to_transfer)

    print_output(parts_changed, parts_not_found, JF_column_names)


    # CR_columns_to_transfer = [1,3,4,5,7,9]
    # JF_columns_to_transfer = [0,4,130,134,3,131] 