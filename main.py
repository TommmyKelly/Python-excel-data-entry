from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime


sg.theme('DarkAmber')

layout = [[sg.Text('First Name'),sg.Push(), sg.Input(key='FIRST_NAME')],
          [sg.Text('Last Name'),sg.Push(), sg.Input(key='LAST_NAME')],
          [sg.Text('TEL:'),sg.Push(), sg.Input(key='NUMBER')],
          [sg.Button('Submit'), sg.Button('Close')]]

window = sg.Window('Data Entry', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    if event == 'Submit':
        try:
            wb = load_workbook('Book1.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [ID, values['FIRST_NAME'], values['LAST_NAME'], values['NUMBER'], time_stamp]

            sheet.append(data)

            wb.save('Book1.xlsx')

            window['FIRST_NAME'].update(value='')
            window['LAST_NAME'].update(value='')
            window['NUMBER'].update(value='')
            window['FIRST_NAME'].set_focus()

            sg.popup('Success', 'Data Saved')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')
        


window.close()