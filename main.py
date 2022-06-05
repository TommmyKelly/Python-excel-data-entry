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
          [sg.Text('File Attachment'), sg.Push(), sg.Input(), sg.FileBrowse(file_types=(("Text Files", "*.txt"),))],
          [sg.Checkbox('Married', key='-MARRIED-')],
          [sg.Checkbox('Has Children', key='-HAS_CHILDREN-')],
          [sg.Push(), sg.Button('Submit'), sg.Button('Close'), sg.Push()]]

# Create the Window
window = sg.Window('Data Entry', layout)


def check_input(values):
    if values['-DUE_DATE-'] == "":
        sg.popup('No date entered')
        return None
    for key, value in values.items():
        print(key)
        print(value)

    try:
        wb = load_workbook('Book1.xlsx')
        sheet = wb['Sheet1']
        id_num = len(sheet['ID']) + 1
        time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        # print(time_stamp)

        if values['Male']:
            gender = 'Male'
        else:
            gender = 'Female'

        data = [id_num, values['FIRST_NAME'],
                values['LAST_NAME'],
                values['NUMBER'],
                gender,
                values['-MARRIED-'],
                values['-HAS_CHILDREN-'],
                datetime.strptime(values['-DUE_DATE-'], '%d/%m/%y').strftime('%d/%m/%Y'),
                time_stamp]

        sheet.append(data)

        wb.save(filename='Book1.xlsx')

        # print('You entered ', values)

        window["FIRST_NAME"].update(value='')
        window["LAST_NAME"].update(value='')
        window["NUMBER"].update(value='')
        window["FIRST_NAME"].set_focus()
        window["Male"].reset_group()
        window["Female"].reset_group()
        window["-MARRIED-"].update(False)
        window["-HAS_CHILDREN-"].update(False)

        sg.popup('Success', 'File Updated')
    except PermissionError:
        sg.popup('File in use', 'File in use by another user.\nPlease try again.')


while True:
    event, values = window.read()
    print(event, values)

    if event == sg.WIN_CLOSED or event == 'Close':  # if user closes window or clicks cancel
        break
    if event == 'Submit':
        check_input(values)


window.close()
