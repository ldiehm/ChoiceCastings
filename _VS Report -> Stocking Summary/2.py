from math import nan
import PySimpleGUI as sg
import pandas as pd
import openpyxl
import numpy as np
import tkinter as tk
from tkinter import font


class Application:

    def __init__(self, appName, sourceFile, sourceFileColumns, targetFile, targetFileColumns, savedFileName):
        # Choice Report -> VS Report  ------ sourceFile
        # Jonesboro Forecast -> Stocking Summary --- targetFile
        self.savedFileName = savedFileName

        self.sourceFileName = sourceFile
        self.targetFileName = targetFile
        self.sourceFile_path = None
        self.targetFile_path = None
        self.save_file_name = None
        
        self.parts_changed = None
        self.parts_not_found = None

        self.sourceFileColumns = sourceFileColumns
        self.targetFileColumns = targetFileColumns

        
        self.layout = [
            # [sg.Button(f'Open {sourceFile}'), sg.Button(f'Open {targetFile}'), sg.Button('Do Data transfer')],
            [sg.Checkbox("Check this box", default=False, key="checkbox_key")],
            [sg.Text(size=(40, 1), key='-OUTPUT-')],
        ]
        self.window = sg.Window(appName, self.layout)

    def read_excel_file(self, file_name):
        try:
            data = pd.read_excel(file_name)
            
        except:
            self.window['-OUTPUT-'].update(f'Invalid file for {file_name}.')
            return None
        return data

    def get_requested_data(self, data, columns_to_transfer):
        
        column_indices = []
        column_headers = data.keys().tolist()
        print(column_headers)
        for column_name in columns_to_transfer:
            column_indices.append(column_headers.index(column_name))
        dataColumns = data.columns[column_indices]
        all_requested_data = data[dataColumns]

        return all_requested_data
    
    def get_column_indices(self, data, headers):
        columnNames = data.keys().tolist()
        columnIndices = []
        for headerNum in headers:
            columnIndices.append(columnNames.index(headers[headerNum]))
        return columnIndices
    
    def get_data_except_part_numbers(self, data, columnNames):
        
        notNumpyArray = []
        print(data)
        indexArray = range(1, len(columnNames))
        for ind in indexArray:
            notNumpyArray.append(data[columnNames[ind]])
        data_without_partNumbers = np.array(notNumpyArray)
        # data_without_partNumbers = np.array([data[headers[1]].tolist(), data[headers[2]].tolist(), data[headers[3]].tolist(), data[headers[4]].tolist(), data[headers[5]].tolist()])
        return data_without_partNumbers
        
    def get_part_numbers(self, data, columnNames):
        part_numbers = []
        
        for cell in data[columnNames[0]]:
            if not pd.isnull(cell):
                if type(cell) == str:
                    cell = cell.strip()
            part_numbers.append(cell)
        return part_numbers
    
    def get_save_file_name(self):
        folders = self.targetFile_path.split('/')
        return folders[-1]
        
    def read_data_files(self):

        allSourceData = self.read_excel_file(self.sourceFile_path)

        allTargetData = self.read_excel_file(self.targetFile_path)     

        ## Here we are constructing the data columns into CR and JF. Part number not included in this big array of arrays.
        #Need to make CR -> VS R
        # self.sourceFileColumns = ["ICPROD","Total On Hand","On Order","Consigned","Required","SumOfWIP"] #REMOVE EVENTUALLY

        requestedSourceData = self.get_requested_data(allSourceData, self.sourceFileColumns)
#
        # column_headers = data_VSR.keys().tolist()
        # CR_columns_to_transfer = [column_headers.index(VSR_column_names[0]), column_headers.index(VSR_column_names[1]), column_headers.index(VSR_column_names[2]), column_headers.index(VSR_column_names[3]), column_headers.index(VSR_column_names[4]), column_headers.index(VSR_column_names[5])]
        # Headers_CR = data_VSR.columns[CR_columns_to_transfer] 
#

        # self.targetFileColumns = ["Part Number","CURRENT Inventory","CURR. Orders","KBM Cons.","CURRENT REQMNT.", "WIP"]

        requestedTargetData = self.get_requested_data(allTargetData, self.targetFileColumns)

        # column_headers = data_SS.keys().tolist()
        # self.JF_columns_to_transfer = [column_headers.index(self.SS_column_names[0]), column_headers.index(self.SS_column_names[1]), column_headers.index(self.SS_column_names[2]), column_headers.index(self.SS_column_names[3]), column_headers.index(self.SS_column_names[4]), column_headers.index(self.SS_column_names[5])]
        # Headers_JF = data_SS.columns[self.JF_columns_to_transfer] # part number, current inventory, current orders, KBM const, current demand, WIP

        ## Here we are making a list of lists of all the part data. Part numbers NOT included for this one. Part numbers are stored in part_numbers_existing_in_JF and part_number_ChoiceReport
        
        
        #CR = np.array([all_data_VSR[Headers_CR[1]].tolist(), data_CR[Headers_CR[2]].tolist(), data_CR[Headers_CR[3]].tolist(), data_CR[Headers_CR[4]].tolist(), data_CR[Headers_CR[5]].tolist()])
        #JF = np.array([data_JF[Headers_JF[1]].tolist(), data_JF[Headers_JF[2]].tolist(), data_JF[Headers_JF[3]].tolist(), data_JF[Headers_JF[4]].tolist(), data_JF[Headers_JF[5]].tolist()])
        sourceDataNoPartNumbers = self.get_data_except_part_numbers(requestedSourceData, self.sourceFileColumns)

        targetDataNoPartNumbers = self.get_data_except_part_numbers(requestedTargetData, self.targetFileColumns)

        ## This is making the part numbers array from Jonesboro Forecast, ensuring no whitespace.

        sourcePartNumbers = self.get_part_numbers(requestedSourceData, self.sourceFileColumns)
        targetPartNumbers = self.get_part_numbers(requestedTargetData, self.targetFileColumns)

        # part_numbers_existing_in_JF = []
        # for cell in data_JF[Headers_JF[0]].tolist():
        #     if not pd.isnull(cell):
        #         cell = cell.strip()
        #     part_numbers_existing_in_JF.append(cell)

        ## Now we begin the iterating process through the whole excel sheet. Each part will be read from the Choice Report and searched for in the Jonesboro Forecast.
        self.parts_changed = []
        self.parts_not_found = []
        is_changed = False
        #We dont do the same part number array building in the Choice Report, so I strip each part number in the for loop as they're read to ensure no white space.
        # data_CR[Headers_CR[0]].tolist()

        for ind, part_number in enumerate(sourcePartNumbers):
            
            # part_number = part_number.strip()/

            if part_number in targetPartNumbers:
                is_changed = False
                part_changes = []
                part_number_ind = targetPartNumbers.index(part_number)
                part_changes.append((part_number, part_number_ind+1))
                
                for i in range(len(self.targetFileColumns) - 1):
                    cell_source_null = sourceDataNoPartNumbers[i][ind]
                    target_cell_null = targetDataNoPartNumbers[i][part_number_ind]

                    try:
                        sourceCell = int(cell_source_null)
                    except:
                        sourceCell = 0

                    try:
                        targetCell = int(float(target_cell_null))
                    except:
                        targetCell = 0

                    if sourceCell != targetCell:
                        is_changed = True
                        part_changes.append((i, targetCell, sourceCell))
                        
                if is_changed:
                    self.parts_changed.append(part_changes)

            else: # Part not found! Important to acknowledge.
                self.parts_not_found.append((part_number, ind))

        return

    #output is Stocking Summary in this case
    def write_to_output(self):
        self.targetFileColumns.pop(0)

        workbook = openpyxl.load_workbook(self.targetFile_path)
        sheet = workbook.active

        for part in self.parts_changed:
            part_number = part[0]

            for change_in_quantity in part[1:]:
                Row = part_number[1] + 1
                Col = change_in_quantity[0] + 1
                cell = sheet.cell(row = Row, column = Col)

                cell.value = change_in_quantity[2]

        workbook.save(self.save_file_name + "_result")
        workbook.close()

    def print_output(self):

        built_string = "Parts changed:\n"
        # Adding text based on changed values
        for changed_part in self.parts_changed:
            part_num = changed_part.pop(0)
            print(changed_part)
            built_string += "\n" + str(part_num[0]) + ": "
            number_for_spacing_output = 0
            for change in changed_part:
                print(change)
                built_string += str(self.targetFileColumns[change[0] + 1]) + " changed from " + str(change[1]) + " to " + str(change[2]) + ". "
                number_for_spacing_output += 1
                if number_for_spacing_output == 2:
                    built_string += "\n"
            if number_for_spacing_output != 2:
                built_string += "\n"
        
        built_string += "\n\nParts not found:\n"

        for part in self.parts_not_found:
            built_string += f"\nDid not found part {part[0]} in {self.targetFileName}. This part was not updated. \nThis part was seen at position {str(part[1] + 2)} in {self.targetFileName}\n"
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
        def _quit():
            root.quit()
            root.destroy() 
        # Create a Close button to close the window
        close_button = tk.Button(root, text="Close", command=_quit)
        close_button.pack()

        # Run the main event loop for the results 
        root.mainloop()
        return    
    
    def run(self):

        while True:
            event, _ = self.window.read()
            
            if event == sg.WIN_CLOSED or event == 'Exit':
                break
            
            elif event == f'Open {self.sourceFileName}':
                self.sourceFile_path = sg.popup_get_file('Select a file', no_window=True)
                if self.sourceFile_path:
                    self.window['-OUTPUT-'].update(str(self.sourceFile_path))

            elif event == f'Open {self.targetFileName}':
                self.targetFile_path = sg.popup_get_file('Select a file', no_window=True)
                if self.targetFile_path:
                    self.window['-OUTPUT-'].update(str(self.targetFile_path))
            
            elif event == 'Do Data transfer':
                if not self.sourceFile_path and not self.targetFile_path:
                    self.window['-OUTPUT-'].update('Need valid files.')
                    continue
                if not self.sourceFile_path:
                    self.window['-OUTPUT-'].update(f'Need valid file for {self.sourceFileName}.')
                    continue
                if not self.targetFile_path:
                    self.window['-OUTPUT-'].update(f'Need valid file for {self.targetFileName}.')
                    continue

                self.save_file_name = self.get_save_file_name()
                self.read_data_files()
                self.write_to_output()
                self.window['-OUTPUT-'].update('Data transfer completed.')
                self.print_output()
                
if __name__ == "__main__":

    app = Application(appName= "Jonesboro Forecast Data Transfer", sourceFile= 'VS Report', sourceFileColumns= ["ICPROD", "Total On Hand","Consigned", "Required", "SumOfWIP"], targetFile= 'Stocking Summary', targetFileColumns=["T&B     Part #", "Total Inv:  JBS & T&B", "VS Cons  On Hand", "Current Demand", "WIP"], savedFileName= 'VS Stocking Summary.xlsx')
    app.run()