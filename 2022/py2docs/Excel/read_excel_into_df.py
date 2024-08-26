
"""
# test_excel_into_df.py
# read Excel file into pandas DataFrame
"""

import os, sys
import pandas as pd

fname = 'files/cdi_pref_mapping.xlsx'
sheet_name = 'ESP Pref Type & Value'

print(f"opening file {fname} using pd.ExcelFile()")
xf = pd.ExcelFile(fname)

sheets = [sname for sname in xf.sheet_names]
print("sheet names = ", sheets)

print(f"parse the sheet into the pandas DataFrame")
df =  xf.parse(sheet_name)
print(df)

print(f"read the sheet using pd.read_excel()")
df2 = pd.read_excel(fname, sheet_name=sheet_name)
print(df2)
