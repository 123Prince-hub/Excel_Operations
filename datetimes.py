from datetime import datetime, date
def dtfrmt(dt):
    try:
        if datetime.strptime(dt, '%d-%b-%Y'):
            print("This is the Correct date string format.")
        else:
            print("Format Not Proper")
    except Exception as e:
        print(e)
a = dtfrmt("31-DEC-2021")

exit()

import xlwings as xw
from datetime import datetime, date

wb = xw.Book("exl.xlsx").sheets['dates']
ws = wb.range("A2").expand().options(numbers=int).value
# wb["B1"].color = (255,255,204)

rowlen = len(ws)
col = input("Enter Your Coloum Name : ").upper()


Choose_Format = input("""
        =======================================================================
        ######################  Choose Your Date Format  ######################
        =======================================================================
                                Press 1 => DD/MM/YY
                                Press 2 => DD/MM/YYYY
                                Press 3 => MM/DD/YY
                                Press 4 => MM/DD/YYYY
                                Press 5 => DD/Mon/YYYY
                                Press 6 => Mon/DD/YYYY
                                Press 7 => DD/Month/YYYY
                                Press 8 => Month/DD/YYYY
        Type Here :->  """)

def format1():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%d-%m-%y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format2():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%d-%m-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format3():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%m-%d-%y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format4():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%m-%d-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format5():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%d-%b-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format6():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%b-%d-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


def format7():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%d-%B-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = f"'{dtes}"
    else:
        pass


def format8():
    global dte
    if dte.startswith("'"):
        dte = row[2].replace("'", "")
        dtes = datetime.strptime(dte, '%d-%b-%Y')
        dtes = datetime.strftime(dtes, '%B-%d-%Y')
        wb.range(f"{col}2:{col}{rowlen+1}").value = dtes
    else:
        pass


for row in ws:
    dte = row[1]
    if Choose_Format == '1':
        format1()
    elif Choose_Format == '2':
        format2()
    elif Choose_Format == '3':
        format3()
    elif Choose_Format == '4':
        format4()
    elif Choose_Format == '5':
        format5()
    elif Choose_Format == '6':
        format6()
    elif Choose_Format == '7':
        format7()
    elif Choose_Format == '8':
        format8()
    else:
        print("You have choose invalid Date format key.")

def inpt_Date():
    pass

    
