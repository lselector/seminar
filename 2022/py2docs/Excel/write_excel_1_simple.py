
""" 
# test_excel_1_simple.py
# writes into files/simple.xlsx
# using pandas.ExcelWriter()
# puts data and a chart
"""

import os, sys
os.environ["PYTHONDONTWRITEBYTECODE"] = "1" 
os.environ["PYTHONUNBUFFERED"] = "1" 
import pandas as pd
import xlsxwriter

print("Create a Pandas dataframe from the data.")
df = pd.DataFrame([10, 20, 30, 20, 15, 30, 45])

print("Create a Pandas Excel writer using xlsxwriter as the engine.")
fname = "files/simple.xlsx"
writer = pd.ExcelWriter(fname, engine='xlsxwriter')

mysheet = 'test1'
print("Write dataframe to sheet", mysheet)
df.to_excel(writer, sheet_name=mysheet)

print("get xlsxwriter objects - workbook and worksheet")
workbook  = writer.book
worksheet = writer.sheets[mysheet]

print("Create a chart object.")
chart = workbook.add_chart({'type': 'column'})

print("Configure the series of the chart from the dataframe data.")
chart.add_series({'values': '='+mysheet+'!$B$2:$B$8'})

print("Insert the chart into the worksheet.")
worksheet.insert_chart('D2', chart)

print("Close the Pandas Excel writer and output the Excel file", fname)
writer.save()

