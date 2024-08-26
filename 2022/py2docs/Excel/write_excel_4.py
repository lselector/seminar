
"""
# write_excel_4.py
# 
# demo of using XlsxWriter 
# to create Excel tables and different charts
# 
# http://xlsxwriter.readthedocs.io/chart_examples.html
"""

import os, sys, xlsxwriter
import pandas as pd
import numpy as np
from mybag import * # simple dictionary class

# --------------------------------------------------------------
def add_charts_data(bag):
    """
    # bag.chart_data that the charts will refer to.
    """
    print("add_charts_data()")
    bag.chart_headings = ['Member ID', '2013-06', '2013-07']
    bag.chart_data = [
        ['Happy Kids', 'Happy Parents', 'Google AdExchange', 
         'Hurraa Media', 'Crocodile Inc.', 'RedHat', 'Bell Labs'],
        [0.11291, 0.22604, 0.09389, 0.02800, 0.04635, 0.17445, 0.007512],
        [0.09444, 0.19879, 0.08919, 0.01877, 0.04173, 0.15331, 0.003999],
    ]
    bag.nn = len(bag.chart_data[0])

# --------------------------------------------------------------
def add_styles_and_formats(bag):
    """
    # Add styles and formats
    """
    print("add_styles_and_formats()")
    bag.fmt = MyBunch()
    bag.fmt.bold                       = bag.workbook.add_format({'bold': 1})
    bag.fmt.dol_int                    = bag.workbook.add_format({'num_format': '$#,##0'})
    bag.fmt.dol_float6                 = bag.workbook.add_format({'num_format': '$0.000000'})
    bag.fmt.dol_acc_int                = bag.workbook.add_format({'num_format': '_($* #,##0_);[red]_($* (#,##0);_($* "-"??_);_(@_)'})
    bag.fmt.dol_acc_float6             = bag.workbook.add_format({'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
    bag.fmt.fg_color_orange            = bag.workbook.add_format()
    bag.fmt.fg_color_orange.set_fg_color('#FE9901')
    bag.fmt.fg_color_black             = bag.workbook.add_format()
    bag.fmt.fg_color_black.set_fg_color('#000000')
    #bag.fmt.col_title                  = bag.workbook.add_format({'bold': True, 'border': True, 'fg_color':'#FE9901'})  #orange
    bag.fmt.col_title                  = bag.workbook.add_format({'bold':1, 'border':1, 'fg_color':'#99CCFF'})  #orange
    bag.fmt.val_row_all_borders        =   bag.workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'right': 1, 'border_color':'#000000'})
    bag.fmt.val_row_left_right_borders =   bag.workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE' , 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
    bag.fmt.val_row_left_right_borders_shade =   bag.workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)', 'fg_color':'#DCE6F1'})
    bag.fmt.val_row_all_borders        =   bag.workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
    bag.fmt.row_top_border             =   bag.workbook.add_format({'top':1, 'border_color':'#000000'})

# --------------------------------------------------------------
def draw_orange_black_lines(bag):
    """
    # draw orange and black horizontal lines below the logo 
    """
    print("draw_orange_black_lines()")
    for col in range(100):
        bag.worksheet.write(6, col, '', bag.fmt.fg_color_orange)
        bag.worksheet.write(7, col, '', bag.fmt.fg_color_black)

# --------------------------------------------------------------
def populate_table(bag):
    """
    # populate data
    # Generate shading to alternate rows
    """
    print("populate_table()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    # header
    bag.worksheet.write(row0, col0  , ''                   , bag.fmt.fg_color_orange)
    bag.worksheet.write(row0, col0  , bag.chart_headings[0], bag.fmt.col_title)
    bag.worksheet.write(row0, col0+1, bag.chart_headings[1], bag.fmt.col_title)
    bag.worksheet.write(row0, col0+2, bag.chart_headings[2], bag.fmt.col_title)

    shade_odd  = bag.fmt.val_row_left_right_borders_shade
    shade_even = bag.fmt.val_row_left_right_borders

    for ii in range(len(bag.chart_data[0])):
        if ii%2 == 1:
            bag.worksheet.write(row0+1+ii, col0  , data[0][ii],  shade_odd)
            bag.worksheet.write(row0+1+ii, col0+1, data[1][ii],  shade_odd)
            bag.worksheet.write(row0+1+ii, col0+2, data[2][ii],  shade_odd)
        else:
            bag.worksheet.write(row0+1+ii, col0  , data[0][ii], shade_even)
            bag.worksheet.write(row0+1+ii, col0+1, data[1][ii], shade_even)
            bag.worksheet.write(row0+1+ii, col0+2, data[2][ii], shade_even)

    # Draw bottom table border
    for col in range(col0+3):
        bag.worksheet.write(row0+1+bag.nn, col, '', bag.fmt.row_top_border)

# --------------------------------------------------------------
def add_column_bar_chart(bag):
    """
    #
    """
    print("add_column_bar_chart()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    chart1 = bag.workbook.add_chart({'type': 'column'})
    
    # Configure the first series.
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+1],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+1, row0+8, col0+1],
    })
    #    'name':       '=Demo2!$B$1',
    #    'categories': '=Demo2!$A$2:$A$8',
    #    'values':     '=Demo2!$B$2:$B$8',
    
    # -------------------------------------
    # Configure a second series. 
    # Note the use of alternative syntax to define ranges.
    # List is [ sheet_name, first_row, first_col, last_row, last_col ].
    
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+2],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+2, row0+8, col0+2],
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Matrix Type: Overall'})
    chart1.set_x_axis({'name': 'GMV Groupes by Sellers'})
    chart1.set_y_axis({'name': 'GMV ($)'})
    
    # Set an Excel chart style.
    chart1.set_style(10)
    
    # Insert the chart into the worksheet (with an offset).
    bag.worksheet.insert_chart('D2', chart1, {'x_offset': 70, 'y_offset': 180})


# --------------------------------------------------------------
def add_area_chart(bag):
    """
    #
    """
    print("add_area_chart()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    chart1 = bag.workbook.add_chart({'type': 'area'})
    
    # Configure the first series.
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+1],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+1, row0+8, col0+1],
    })
    
    # Configure a second series. 
    # Note the use of alternative syntax to define ranges.
    # List is [ sheet_name, first_row, first_col, last_row, last_col ].
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+2],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+2, row0+8, col0+2],
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Matrix Type: Overall'})
    chart1.set_x_axis({'name': 'GMV Groupes by Sellers'})
    chart1.set_y_axis({'name': 'GMV ($)'})
    
    # Set an Excel chart style.
    chart1.set_style(18)
    
    # Insert the chart into the worksheet (with an offset).
    bag.worksheet.insert_chart('D2', chart1, {'x_offset': 600, 'y_offset': 180}) # x=600, y=180
        

# --------------------------------------------------------------
def populate_table2(bag):
    """
    # second table
    """
    print("populate_table2()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    bag.worksheet.write(row0, col0  , '', bag.fmt.fg_color_orange)
    bag.worksheet.write(row0, col0  , bag.chart_headings[0], bag.fmt.col_title)
    bag.worksheet.write(row0, col0+1, bag.chart_headings[1], bag.fmt.col_title)
    bag.worksheet.write(row0, col0+2, bag.chart_headings[2], bag.fmt.col_title)
    
    shade_odd  = bag.fmt.val_row_left_right_borders_shade
    shade_even = bag.fmt.val_row_left_right_borders
    
    for ii in range(len(bag.chart_data[0])):
        if ii%2 == 1:
            bag.worksheet.write(row0+1+ii, col0  , data[0][ii], shade_odd) # dol_acc_float6
            bag.worksheet.write(row0+1+ii, col0+1, data[1][ii], shade_odd)
            bag.worksheet.write(row0+1+ii, col0+2, data[2][ii], shade_odd)
        else:
            bag.worksheet.write(row0+1+ii, col0  , data[0][ii], shade_even)
            bag.worksheet.write(row0+1+ii, col0+1, data[1][ii], shade_even)
            bag.worksheet.write(row0+1+ii, col0+2, data[2][ii], shade_even)
    
    # -------------------------------------
    # Draw bottom table border
    for col in range(col0+3):
        bag.worksheet.write(row0+1+bag.nn, col, '', bag.fmt.row_top_border)

# --------------------------------------------------------------
def add_pie_chart(bag):
    """
    # pie chart
    """
    print("add_pie_chart()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    chart1 = bag.workbook.add_chart({'type': 'pie'})
    
    # Configure the first series.
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+1],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+1, row0+8, col0+1],
    })
    
    # Configure a second series. 
    # Note the use of alternative syntax to define ranges.
    # List is [ sheet_name, first_row, first_col, last_row, last_col ].
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+2],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+2, row0+8, col0+2],
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Matrix Type: Overall'})
    chart1.set_x_axis({'name': 'GMV Groupes by Sellers'})
    chart1.set_y_axis({'name': 'GMV ($)'})
    
    # Set an Excel chart style.
    chart1.set_style(10)
    
    # Insert the chart into the worksheet (with an offset).
    bag.worksheet.insert_chart('D2', chart1, {'x_offset': 70, 'y_offset': 580})

def add_percent_stacked_chart(bag):
    """
    # add_percent_stacked_chart
    """
    print("add_percent_stacked_chart()")
    data = bag.chart_data
    row0 = bag.row0
    col0 = bag.col0

    chart1 = bag.workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})

    # Configure the first series.
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+1],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+1, row0+8, col0+1],
    })
    
    # Configure a second series. 
    # Note use of alternative syntax to define ranges.
    # List is [ sheet_name, first_row, first_col, last_row, last_col ].
    chart1.add_series({
        'name'       : '%s' % bag.chart_headings[col0+2],
        'categories' : ['Demo2', row0+1 , col0  , row0+8, col0],
        'values'     : ['Demo2', row0+1 , col0+2, row0+8, col0+2],
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Matrix Type: Overall'})
    chart1.set_x_axis({'name': 'GMV Groupes by Sellers'})
    chart1.set_y_axis({'name': 'GMV ($)'})
    
    # Set an Excel chart style.
    chart1.set_style(18)
    
    # Insert the chart into the worksheet (with an offset).
    bag.worksheet.insert_chart('D2', chart1, {'x_offset': 600, 'y_offset': 580}) # x=600, y=180

# --------------------------------------------------------------
def main(bag):
    """
    # main execution
    """
    bag.row0 = 10
    bag.col0 = 0
    bag.files_path = 'files'
    bag.fname = bag.files_path + '/test2_xlsx_demo.xlsx'
    print(f"creating file {bag.fname}")
    bag.workbook = xlsxwriter.Workbook(bag.fname) 
    bag.worksheet = bag.workbook.add_worksheet(name="Demo2")
    # -------------------------------------
    # Adjust the column width
    bag.worksheet.set_column(0,0, 20)
    bag.worksheet.set_column(1, 2, 15)
    # -------------------------------------
    # Insert a logo image
    bag.worksheet.insert_image(0, 0, '%s/mylogo.png' % bag.files_path, {'x_scale': 0.5, 'y_scale': 0.5}) 
    bag.worksheet.hide_gridlines(2)
    # -------------------------------------
    add_charts_data(bag)
    add_styles_and_formats(bag)
    draw_orange_black_lines(bag)
    populate_table(bag)
    add_column_bar_chart(bag)
    add_area_chart(bag)
    bag.row0 += 20
    populate_table2(bag)
    add_pie_chart(bag)
    add_percent_stacked_chart(bag)

    bag.workbook.close()

# --------------------------------------------------------------
if __name__ == "__main__":
    # print("running Python " + str(sys.version_info))
    bag = MyBunch()
    main(bag)


