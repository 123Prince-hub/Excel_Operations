import xlwings as xw

wb = xw.Book("exl.xlsx").sheets['digits']
ws = wb.range("A2").expand().options(transpose=True, numbers=int).value
digits = ws[1]

col = input("Enter Your Coloum Name : ").upper()
num = 2 
for row in digits:
    wb.range(f"{col}{num}").value = f"'{row}"
    num += 1