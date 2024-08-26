
"""
# write_excel_3_simple.py
# 
# very simple demo of using XlsxWriter to write DataFrame
# with different formatting and alignment for different cells
# 
# http://xlsxwriter.readthedocs.io/chart_examples.html
"""

import sys, os
import pandas as pd
import numpy as np
import xlsxwriter

from mybag import * # library which creates a custom dictionary class

# --------------------------------------------------------------
def add_data(bag):
    """
    # bag.data
    """
    print("add_data()")
    bag.cols = ['my_id', 'my_name', 'usd']
    my_arr = [
      [ 25,    b'\xe5\xa4\xa7'.decode(),        54321.0 ],
      [ 36,    "Abraham Lincoln", 123456789.0 ],
      [ 19546, "Mary Stuart",            27.0 ],
      [ 25,    b'\xe5\xb7\xa8\xe5\xa4\xa7'.decode(),        54321.0 ],
      [ 36,    "Abraham Lincoln 2", 123456789.0 ],
      [ 19546, "Mary Stuart 2",            27.0 ],
    ]
    bag.df = pd.DataFrame(data=my_arr, columns = bag.cols)

# --------------------------------------------------------------
def add_styles_and_formats(bag):
    """
    # Add styles and formats
    """
    print("add_styles_and_formats()")
    bag.fmt = MyBunch()
    bag.fmt.bold                       = bag.workbook.add_format({'bold': 1})
    bag.fmt.dol_int                    = bag.workbook.add_format({'num_format': '$#,##0'})
#    bag.fmt.dol_float6                 = bag.workbook.add_format({'num_format': '$0.000000'})
#    bag.fmt.dol_acc_int                = bag.workbook.add_format({'num_format': '_($* #,##0_);[red]_($* (#,##0);_($* "-"??_);_(@_)'})
#    bag.fmt.dol_acc_float6             = bag.workbook.add_format({'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
     
##    bag.fmt.fg_color_orange            = bag.workbook.add_format()
##    bag.fmt.fg_color_orange.set_fg_color('#FE9901')
##    bag.fmt.fg_color_black             = bag.workbook.add_format()
##    bag.fmt.fg_color_black.set_fg_color('#000000')
##    #bag.fmt.col_title                  = bag.workbook.add_format({'bold': True, 'border': True, 'fg_color':'#FE9901'})  #orange

    bag.fmt.col_title                  = bag.workbook.add_format({'bold':1, 'border':1, 'fg_color':'#fbd190'})
    bag.fmt.val_row_all_borders        =   bag.workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'right': 1, 'border_color':'#000000'})
    bag.fmt.val_row_left_right_borders =   bag.workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE' , 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
    bag.fmt.val_row_left_right_borders_shade =   bag.workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)', 'fg_color':'#DCE6F1'})
    bag.fmt.val_row_all_borders        =   bag.workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
    bag.fmt.row_top_border             =   bag.workbook.add_format({'top':1, 'border_color':'#000000'})


# --------------------------------------------------------------
def populate_table(bag):
    """
    # populate data
    # Generate shading to alternate rows
    """
    print("populate_table()")
    row0 = bag.row0
    col0 = bag.col0
    ws   = bag.worksheet
    wb   = bag.workbook
    headers = list(bag.df.columns)
    for ii in range(len(headers)):
        col = col0+ii
        txt = headers[ii]
        ws.write(row0, col  , txt, bag.fmt.col_title)

    d_str0 = { 'font_size':12, 'left':1, 'right':1, 'bottom':1,
                  'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE'}
    d_str1 = d_str0.copy()
    d_str1['fg_color'] = '#f6f6f6'
    d_usd0 = d_str0.copy()
    d_usd0['num_format'] = '$#,##0'
    d_usd1 = d_usd0.copy()
    d_usd1['fg_color'] = '#f6f6f6'

    bag.fmt.str0 = wb.add_format(d_str0)
    bag.fmt.str1 = wb.add_format(d_str1)
    bag.fmt.usd0 = wb.add_format(d_usd0)
    bag.fmt.usd1 = wb.add_format(d_usd1)

    ii = -1
    for row in bag.df.itertuples():
        ii += 1
        (my_id, my_name, usd) = row[1:]
        if ii%2 == 0:  
            ws.write(row0+1+ii, col0  , my_id,   bag.fmt.str0)
            ws.write(row0+1+ii, col0+1, my_name, bag.fmt.str0)
            ws.write(row0+1+ii, col0+2, usd,     bag.fmt.usd0)
        else:
            ws.write(row0+1+ii, col0  , my_id,   bag.fmt.str1)
            ws.write(row0+1+ii, col0+1, my_name, bag.fmt.str1)
            ws.write(row0+1+ii, col0+2, usd,     bag.fmt.usd1)


    row_bottom = row0+1+len(bag.df)
    for ii in range(len(headers)):
        ws.write(row_bottom, col0+ii, '', bag.fmt.row_top_border)

# --------------------------------------------------------------
def main(bag):
    """
    # main execution
    """
    bag.row0 = 5
    bag.col0 = 2
    bag.files_path = 'files'
    bag.fname = bag.files_path + '/test1_xlsx_demo.xlsx'
    bag.workbook = xlsxwriter.Workbook(bag.fname) 
    bag.worksheet = bag.workbook.add_worksheet(name="Demo1")
    print(f"creating file {bag.fname}")

    # -------------------------------------
    # Adjust the column width
    # bag.worksheet.set_column(0,0, 20)
    # bag.worksheet.set_column(1, 2, 15)
    # -------------------------------------
    add_data(bag)
    add_styles_and_formats(bag)
    populate_table(bag)
    bag.workbook.close()

# --------------------------------------------------------------
if __name__ == "__main__":
    # print("running Python " + str(sys.version_info))
    bag = MyBunch()
    main(bag)


