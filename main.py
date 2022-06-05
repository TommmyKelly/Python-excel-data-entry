from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

sg.theme('DarkAmber')  # Add a touch of color
# All the stuff inside your window.
layout = [[sg.Text('First Name'), sg.Push(), sg.Input(key="FIRST_NAME")],
          [sg.Text('Last Name'), sg.Push(), sg.Input(key="LAST_NAME")],
          [sg.Text('TEL:'), sg.Push(), sg.Input(key="NUMBER")],
          [sg.Text('Male / Female'), sg.Radio('Male', 'RADIO_GENDER', key='Male'),
           sg.Radio('Female', 'RADIO_GENDER', key='Female')],
          [sg.Text('Due Date'), sg.Push(), sg.Input(key='-DUE_DATE-', size=(34, 1)),
           sg.CalendarButton("Pick Date", close_when_date_chosen=True, target="-DUE_DATE-",
                             format='%d/%m/%y', no_titlebar=False, title="Pick due date")],
          [sg.Input() ,sg.FileBrowse(file_types=(("Text Files", "*.txt"),))],
          [sg.Push(), sg.Button('Submit'), sg.Button('Close'), sg.Push()]]

# Create the Window
window = sg.Window('Data Entry', layout)

while True:
    event, values = window.read()
    print(event, values)
    gender = None
    if event == sg.WIN_CLOSED or event == 'Close':  # if user closes window or clicks cancel
        break
    if event == 'Submit':
        try:
            wb = load_workbook('Book1.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(time_stamp)

            if values['Male']:
                gender = 'Male'
            else:
                gender = 'Female'

            data = [ID, values['FIRST_NAME'], values['LAST_NAME'],
                    values['NUMBER'], gender, datetime.strptime(values['-DUE_DATE-'],
                                                                '%d/%m/%y').strftime('%d/%m/%Y')
                    , time_stamp]

            sheet.append(data)

            wb.save(filename='Book1.xlsx')

            print('You entered ', values)

            window["FIRST_NAME"].update(value='')
            window["LAST_NAME"].update(value='')
            window["NUMBER"].update(value='')
            window["FIRST_NAME"].set_focus()

            sg.popup('Success', 'File Updated')
        except PermissionError:
            sg.popup('File in use', 'File in use by another user.\nPlease try again.')

window.close()
