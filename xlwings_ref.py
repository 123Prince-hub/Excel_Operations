from time import sleep
import xlwings as xw
import datetime as dt
import pandas as pd
import numpy as np

wb = xw.Book('test.xlsx')

# wb1 = xw.Book('test.xlsx').sheets[0]
# wb2 = xw.Book('test.xlsx').sheets[1]
# wb3 = xw.Book('test.xlsx').sheets[2]
# wb.save('test.xlsx')
# print(wb1.name)
# print(wb2.name)
# print(wb3.name)


# ws = wb.sheets[0].range('A1').value = "'123647784547"       # text to coloumn in excel
# wb1.range('A1').value = "Hello World"
# wb1.cells(1, 2).value = 100      # asign cell value(row number, coloumn number)        
# wb1.cells(2, "C").value = 200    # asign cell value(row number, Coloumn name)
# wb1.range('A1').clear()      # clear value form specific cell
# wb1.range('A1:E20').value = 100
# wb1.range('A1').value.clear_contents()     # clear all content form sheet

# wb1.range('A1').value = [10, 20, 50]       # insert value using list
# wb1.range('A1').value = [['Col A', 'Col B'], [10, 20], [30, 40]]      # create table 

# wb1.range('A1').options(transpose=True).value = [10, 20, 50, 60, 100]     # insert value verticle format using list
# wb1.range('B1').options(transpose=True).value = [100, 200, 500, 600, 700]     # insert value verticle format using list

# print(wb1.range('A1').expand().value)     # get all data using expand in list


# ***************** ndim  **************
# wb = xw.Book('test.xlsx')
# ws = wb.sheets[0].range('A1:A1').options(ndim=2).value     # ndim convert value in list like ndim=1 is one list and ndim=2 is double list  
# # ws = wb.sheets[0].range('A1:B1').options(ndim=2).value     # ndim convert value in list like ndim=1 is one list and ndim=2 is double list  
# print(ws)




# ***************** numbers  **************
# wb = xw.Book('test.xlsx')
# ws = wb.sheets[0].range('A1:A2').options(numbers=int).value     # numbers convert value in int like 
# print(ws)




# ***************** numbers  **************
# wb = xw.Book('test.xlsx')
# ws = wb.sheets[0].range('A1').options(dates=dt.date).value     # numbers convert value in datetime formate 
# # ws = wb.sheets[0].range('A1').value      
# print(type(ws))
# print(ws)

# *******************************    date to string    ******************************
# from datetime import datetime
# now = datetime.now()
# date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
# print("date and time:",date_time)
# print(date_time[0])





# # ***************** empty  **************
# wb = xw.Book('test.xlsx')
# # ws = wb.sheets[0].range('A1:C1').value     # empty/None cell value 
# ws = wb.sheets[0].range('A1:C1').options(empty="NA").value     # empty/None cell value convert default   
# print(ws)




# *************************    data analytic operation using pandas and numpy   *************************



# *************************    dictionary format one dimantional   *************************
# wsDict = wb.sheets['Dict']
# print(wsDict.range('A1:B2').options(dict, numbers=int).value)     # multiple option attribute at a time
# print(wsDict.range('A1:B2').options(dict, transpose=True).value)



# *************************    numpy format one/multi dimantional   *************************
# wsNumpy = wb.sheets['Numpy']
# arry1 = np.array([10, 20, 30])
# wsNumpy.range('A1').options(transpose=True).value = arry1

# print(wsNumpy.range('A1').options(np.array, expand='table', ndim=1).value)
# print(wsNumpy.range('A1').options(np.array, expand='table', ndim=2).value)




# # *************************    pandas series format one/multi dimantional   *************************
# wsPandas = wb.sheets['Pandas']
# # wsPd = wsPandas.range('A1').expand().value
# # wsPd = wsPandas.range('A1').options(pd.Series, index=False, header=True, expand='table').value

# series1 = wsPandas.range('A1').expand().value
# adding = wsPandas.range('D1').expand().value = series1
# print(adding)



# # **************    pandas datafreame format one/multi dimantional   *************************
# wsdf = wb.sheets['pd_dataframe']
# df1 = wsdf.range('A1').options(pd.DataFrame, header=1, expand='table').value
# df2 = wsdf.range('A1').options(pd.DataFrame, index=False, header=2, expand='table').value
# print(df1)
# print(df2)



# **************    Xlwings Api   *************************
ws = wb.sheets['Xlwings_api']
LastRow = ws.cells(ws.api.rows.count, "A")
print(LastRow)
