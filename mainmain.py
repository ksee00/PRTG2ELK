from pathlib import Path  # core python module

import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
import re


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
    menu_def = [["Help", ["Settings", "About", "Exit"]]]


    # ------ GUI Definition ------ #
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)],
              [sg.T("Input File:", s=15, justification="r"), sg.I(key="-IN-", s=80, font=('Arial', 10), do_not_clear=False), sg.FileBrowse(file_types=(("CSV", "*.csv*"),))],
              [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-", s=80, font=('Arial', 10), do_not_clear=False), sg.FolderBrowse()],
              [sg.T("Device Id:", s=15, justification="r"), 
                    sg.I(key="-DEVICE_ID-", s=10, do_not_clear=False), 
                    sg.T("* Must be prefix with ID. E.g. ID1234", font=('Lucida', 8), text_color='Gray')],
              [sg.Exit(s=16, button_color="tomato"),sg.B("Data cleansing", s=16)],]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "About":
            window.disappear()
            sg.popup(window_title, "Version 1.0", "Data Cleansing", grab_anywhere=True)
            window.reappear()
        if event in ("Command 1", "Command 2", "Command 3", "Command 4"):
            sg.popup_error("Not yet implemented")
        if event == "Settings":
            settings_window(settings)
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
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="conf\\config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()