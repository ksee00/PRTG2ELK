from pathlib import Path  # core python module

import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
import re
import csv
import os

def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False

def is_valid_device_id(_deviceId):
    if _deviceId:
        match = re.search(r'.*ID\d{4}', _deviceId)
        if match:
            return True

    sg.popup_error("Device ID is invalid.\n" + "Ensure ID is uppercase followed by 4 digits")
    return False

def display_excel_file(excel_file_path, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).name
    sg.popup_scrolled(df.dtypes, "=" * 50, df, title=filename)


def convert_to_csv(excel_file_path, output_folder, sheet_name, separator, decimal):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, sep=separator, decimal=decimal, index=False)
    sg.popup_no_titlebar("Done! :)")


'''
Func:   add_str_to_lines(f_name, output_folder, str_to_add) 
Desc:   1. To add string to the end of each string.
        2. To add Fieldname "DeviceId" to header
'''
def add_str_to_lines(f_name, output_folder, str_to_add):
    header_string_to_add = f',"DeviceId"'    # hardcoded
    column_to_add = f',"{str_to_add}"'
    with open(f_name, "r") as f:
        lines = f.readlines()
        for index, line in enumerate(lines):
            if index == 0:
                lines[index] =  line.strip() + header_string_to_add  +"\n"
                continue
            
            lines[index] = line.strip() + column_to_add + "\n"
    filename = Path(f_name).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    
    with open(outputfile, "w") as f:
        for line in lines:
            f.write(line)


    cleansing_data(outputfile, output_folder)
    # sg.popup_no_titlebar("Done! :)")


def cleansing_data(f_name, output_folder):
    processing_file  = f_name
    final_processed_file = f_name

    filename = Path(f_name).stem
    interim_processed_file  = Path(output_folder) / f"{filename}_temp.csv"

    # Obtains the total of lines in the processing file.
    total_lines_in_processing_file = totalLinesProcessingFile(f_name)
    print(f'Total Lines in Processing File {total_lines_in_processing_file} lines.')

        # Obtains the total of lines in the processing file.
    total_columns_in_processing_row = totalColumnsProcessingRow(f_name)
    print(f'Total Columns in Processing Row {total_columns_in_processing_row} columns.')

    with open(interim_processed_file, mode='w') as processed_file_handler: 
        csv_writer = csv.writer(processed_file_handler, delimiter=',', lineterminator='\r', quoting=csv.QUOTE_MINIMAL, doublequote=False )
        with open(processing_file) as processing_file_handler:
            csv_reader = csv.reader(processing_file_handler, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if line_count == 0:
                    #print(f'Column names are {", ".join(row)}')
                    line_count += 1
                else:                                   
                    # 1. Ensure number of columns is as per heading.
                    number_of_readin_cols = len(row)

                    if (number_of_readin_cols != total_columns_in_processing_row):
                        continue

                    # 2. If value in first column consists of "values", skip 
                    if(in_(row[0], "values")):
                        continue

                    if(in_(row[0], "Date")):
                        print("Found Date")
                        continue


                    csv_writer.writerow(row)
                    line_count += 1
            print(f'Processed {line_count} / {total_lines_in_processing_file} lines.')
        processing_file_handler.close()
    processed_file_handler.close()


    with open(interim_processed_file, 'r') as f, open(final_processed_file, 'w') as fo:
        for line in f:
            fo.write(line.replace('"', '').replace("'", ""))

    f.close()
    fo.close()

    if os.path.isfile(interim_processed_file):
        os.remove(interim_processed_file)

# 
# Desc: To obtain the total of lines in the file.
# 
def totalLinesProcessingFile(_processing_file):
    _input_file = open(_processing_file,"r+")
    _reader_file = csv.reader(_input_file)
    _value = len(list(_reader_file))
    return _value
# 
# Desc: To obtain the total of columns in the row.
# 
def totalColumnsProcessingRow(_processing_file):
    with open(_processing_file, 'r') as csv:
        _first_line = csv.readline()
        _your_data = csv.readlines()

    _number_of_col = _first_line.count(',') + 1 
    return _number_of_col

def in_(s, other):
    return other in s


def settings_window(settings):
    # ------ GUI Definition ------ #
    layout = [[sg.T("SETTINGS")],
              [sg.T("Separator"), sg.I(settings["CSV"]["separator"], s=1, key="-SEPARATOR-"),
               sg.T("Decimal"), sg.Combo(settings["CSV"]["decimal"].split("|"),
                                   default_value=settings["CSV"]["decimal_default"],
                                   s=1, key="-DECIMAL-"),
               sg.T("Sheet Name:"), sg.I(settings["EXCEL"]["sheet_name"], s=20, key="-SHEET_NAME-")],
              [sg.B("Save Current Settings", s=20)]]

    window = sg.Window("Settings Window", layout, modal=True, use_custom_titlebar=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "Save Current Settings":
            # Write to ini file
            settings["CSV"]["separator"] = values["-SEPARATOR-"]
            settings["CSV"]["decimal_default"] = values["-DECIMAL-"]
            settings["EXCEL"]["sheet_name"] = values["-SHEET_NAME-"]

            # Display success message & close window
            sg.popup_no_titlebar("Settings saved!")
            break
    window.close()


def main_window():
    # ------ Menu Definition ------ #
    menu_def = [["Help", ["About", "Exit"]]]


    # ------ GUI Definition ------ #
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)],
              [sg.T("Input File:", s=15, justification="r"), sg.I(key="-IN-", s=80, font=('Arial', 10), do_not_clear=False), sg.FileBrowse(file_types=(("CSV", "*.csv*"),))],
              [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-", s=80, font=('Arial', 10), do_not_clear=False), sg.FolderBrowse()],
              [sg.T("Device Id:", s=15, justification="r"), 
                    sg.I(key="-DEVICE_ID-", s=10, do_not_clear=False), 
                    sg.T("* Must be prefix with ID. E.g. ID1234", font=('Lucida', 8), text_color='Gray')],
              [sg.Exit(s=16, button_color="tomato"),sg.B("Data cleansing", s=16)],]

    window_title = "Data Cleansing"
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "About":
            window.disappear()
            sg.popup(window_title, "Version 1.0", "Data Cleansing", grab_anywhere=True)
            window.reappear()
        if event == "Data cleansing":
            if(is_valid_device_id(values["-DEVICE_ID-"])):
                pass

            if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):

                window.disappear()
                
                add_str_to_lines(
                    f_name = values["-IN-"], 
                    output_folder = values["-OUT-"], 
                    str_to_add = values["-DEVICE_ID-"])
                
                sg.popup(window_title, "Completed", grab_anywhere=True)
                window.reappear()


    window.close()


if __name__ == "__main__":
    theme = "Reddit"
    font_family = "Arial"
    font_size = 12
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()