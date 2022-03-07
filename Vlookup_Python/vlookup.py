import pandas as pd
import numpy as np
from pathlib import Path
import xlwings as xw

# print(Path('data.xlsx').absolute())     #************ get complete file path
file_path = Path.cwd() / 'data.xlsx'

excel_file = pd.ExcelFile(file_path)
sheets_name = excel_file.sheet_names

orders = pd.read_excel(file_path, sheet_name='Orders')
returns = pd.read_excel(file_path, sheet_name='Returns')    
shipping = pd.read_excel(file_path, sheet_name='Shipping')

# dataframe = {} 
# for sheet_name in sheets_name:
#     dataframe[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)


df1 = orders.merge(returns, left_on="Order ID", right_on="ID", how="left")
df2 = df1.merge(shipping, left_on="Ship Mode", right_on="Ship Mode", how="left")


# file_output = Path.cwd() / 'a.xlsx'
#******************Export to new excel workbook
# df2.to_excel(file_output, sheet_name="Output", index=False)

# ***************** Export to same excel workbook
try:
    wb = xw.Book(file_path)
    new_sht = wb.sheets.add('Output', after='Shipping')
    new_sht.range('A1').options(index=False).value = df2
except Exception as e:
    print(e)


