from datetime import datetime, date


def convert_to_excel_serial(day, month, year):
    current = date(year, month, day)

    n = current.toordinal()
    offset = 693594
    return n - offset
