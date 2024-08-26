
"""
# write_excel_2.py
# An example of creating Excel with logo and several charts
# using XlsxWriter.
"""

import os, sys
import xlsxwriter

# --------------------------------------------------------------
# main execution
# --------------------------------------------------------------
row0 = 10
col0 = 0
mypath = 'files'
fname = f'{mypath}/test_xlsx_demo.xlsx'
print(f"creating file {fname}")
workbook = xlsxwriter.Workbook(fname)
worksheet = workbook.add_worksheet()

# -------------------------------------
# Adjust the column width
worksheet.set_column(0,0, 20)
worksheet.set_column(1, 2, 15)

# -------------------------------------
# Insert a logo image
worksheet.insert_image(0, 0, f"{mypath}/mylogo.png", {'x_scale': 1.1, 'y_scale': 1.1}) 
worksheet.hide_gridlines(2)

# -------------------------------------
# Add the worksheet data that the charts will refer to.
headings = ['Member ID', '2013-06', '2013-07']
data = [
    ['Rubicon', 'Switch Concepts Limited', 'Google AdExchange', 
     'Redux Media', 'Heights Media', 'OpenX', 'bRealTime'],
    [0.11291 , 0.22604, 0.09389, 0.02800, 0.04635, 0.17445, 0.007512],
    [0.09444 , 0.19879, 0.08919, 0.01877, 0.04173, 0.15331, 0.003999],
]
nn = len(data[0])

# -------------------------------------
# Add styles and formats

bold                       = workbook.add_format({'bold': 1})
dol_int                    = workbook.add_format({'num_format': '$#,##0'})
dol_float6                 = workbook.add_format({'num_format': '$0.000000'})
dol_acc_int                = workbook.add_format({'num_format': '_($* #,##0_);[red]_($* (#,##0);_($* "-"??_);_(@_)'})
dol_acc_float6             = workbook.add_format({'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
fg_color_orange            = workbook.add_format()
fg_color_orange.set_fg_color('#FE9901')
fg_color_black             = workbook.add_format()
fg_color_black.set_fg_color('#000000')
# col_title                  = workbook.add_format({'bold': True, 'border': True, 'fg_color':'#FE9901'})  #orange
col_title                  = workbook.add_format({'bold':1, 'border':1, 'fg_color':'#99CCFF'})  #orange
val_row_all_borders        =   workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'right': 1, 'border_color':'#000000'})
val_row_left_right_borders =   workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE' , 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
val_row_left_right_borders_shade =   workbook.add_format({'font_size':12, 'left':1, 'right':1, 'bottom':1,'left_color':'#000000', 'right_color':'#000000', 'bottom_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)', 'fg_color':'#DCE6F1'})
val_row_all_borders        =   workbook.add_format({'font_size':12, 'border':1, 'border_color':'#CECECE', 'num_format': '_($* 0.000000_);[red]_($* (0.000000);_($* "-"??_);_(@_)'})
row_top_border             =   workbook.add_format({'top':1, 'border_color':'#000000'})

# -------------------------------------                  
# Draw orange and black horizontal lines below the logo 

for col in range(100):
    worksheet.write(6, col, '', fg_color_orange)
    worksheet.write(7, col, '', fg_color_black)
# -------------------------------------                  
# Populate data  

# worksheet.write_row('A1', headings, bold)
# worksheet.write_column('A2', data[0], dol_acc_float6)
# worksheet.write_column('B2', data[1], dol_acc_float6)
# worksheet.write_column('C2', data[2], dol_acc_float6)

worksheet.write(row0, col0  , '', fg_color_orange)
worksheet.write(row0, col0  , headings[0], col_title)
worksheet.write(row0, col0+1, headings[1], col_title)
worksheet.write(row0, col0+2, headings[2], col_title)

# Generate shading to alternate rows
for ii in range(len(data[0])):
    if ii%2 == 1:
        worksheet.write(row0+1+ii, col0  , data[0][ii],  val_row_left_right_borders_shade)
        worksheet.write(row0+1+ii, col0+1, data[1][ii],  val_row_left_right_borders_shade)
        worksheet.write(row0+1+ii, col0+2, data[2][ii],  val_row_left_right_borders_shade)
    else:
         worksheet.write(row0+1+ii, col0  , data[0][ii], val_row_left_right_borders)
         worksheet.write(row0+1+ii, col0+1, data[1][ii], val_row_left_right_borders)
         worksheet.write(row0+1+ii, col0+2, data[2][ii], val_row_left_right_borders)

# -------------------------------------
# Draw bottom table border
for col in range(col0+3):
    worksheet.write(row0+1+nn, col, '', row_top_border)
# -------------------------------------
# Create a new column bar chart.
#
chart1 = workbook.add_chart({'type': 'column'})

# Configure the first series.
chart1.add_series({
    'name':       '%s' % headings[col0+1],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+1, row0+8, col0+1],
})
#    'name':       '=Sheet1!$B$1',
#    'categories': '=Sheet1!$A$2:$A$8',
#    'values':     '=Sheet1!$B$2:$B$8',

# -------------------------------------
# Configure a second series. Note use of alternative syntax to define ranges.
# List is [ sheet_name, first_row, first_col, last_row, last_col ].

chart1.add_series({
    'name':       '%s' % headings[col0+2],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+2, row0+8, col0+2],
})

# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Matrix Type: Overall'})
chart1.set_x_axis({'name': 'RPM Groupes by Sellers'})
chart1.set_y_axis({'name': 'RPM ($)'})

# Set an Excel chart style.
chart1.set_style(10)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('D2', chart1, {'x_offset': 70, 'y_offset': 180})

# -------------------------------------
# Create a new area chart.
#
chart1 = workbook.add_chart({'type': 'area'})

# Configure the first series.
chart1.add_series({
    'name':       '%s' % headings[col0+1],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+1, row0+8, col0+1],
})

# Configure a second series. Note use of alternative syntax to define ranges.
# List is [ sheet_name, first_row, first_col, last_row, last_col ].
chart1.add_series({
    'name':       '%s' % headings[col0+2],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+2, row0+8, col0+2],
})

# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Matrix Type: Overall'})
chart1.set_x_axis({'name': 'RPM Groupes by Sellers'})
chart1.set_y_axis({'name': 'RPM ($)'})

# Set an Excel chart style.
chart1.set_style(18)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('D2', chart1, {'x_offset': 600, 'y_offset': 180}) # x=600, y=180
    

## -------------------------------------

row0 = row0 + 20
worksheet.write(row0, col0  , '', fg_color_orange)
worksheet.write(row0, col0  , headings[0], col_title)
worksheet.write(row0, col0+1, headings[1], col_title)
worksheet.write(row0, col0+2, headings[2], col_title)

for ii in range(len(data[0])):
    if ii%2 == 1:
        worksheet.write(row0+1+ii, col0  , data[0][ii], val_row_left_right_borders_shade) # dol_acc_float6
        worksheet.write(row0+1+ii, col0+1, data[1][ii], val_row_left_right_borders_shade)
        worksheet.write(row0+1+ii, col0+2, data[2][ii], val_row_left_right_borders_shade)
    else:
         worksheet.write(row0+1+ii, col0  , data[0][ii], val_row_left_right_borders)
         worksheet.write(row0+1+ii, col0+1, data[1][ii], val_row_left_right_borders)
         worksheet.write(row0+1+ii, col0+2, data[2][ii], val_row_left_right_borders)

# -------------------------------------
# Draw bottom table border
for col in range(col0+3):
    worksheet.write(row0+1+nn, col, '', row_top_border)
    
# -------------------------------------
# Create a new pie chart.
#
chart1 = workbook.add_chart({'type': 'pie'})

# Configure the first series.
chart1.add_series({
    'name':       '%s' % headings[col0+1],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+1, row0+8, col0+1],
})

# Configure a second series. Note use of alternative syntax to define ranges.
# List is [ sheet_name, first_row, first_col, last_row, last_col ].
chart1.add_series({
    'name':       '%s' % headings[col0+2],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+2, row0+8, col0+2],
})

# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Matrix Type: Overall'})
chart1.set_x_axis({'name': 'RPM Groupes by Sellers'})
chart1.set_y_axis({'name': 'RPM ($)'})

# Set an Excel chart style.
chart1.set_style(10)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('D2', chart1, {'x_offset': 70, 'y_offset': 580})
    
# =====================================
# Create a new percent stacked chart.
#
chart1 = workbook.add_chart({'type': 'column', 'subtype': 'percent_stacked'})

# Configure the first series.
chart1.add_series({
    'name':       '%s' % headings[col0+1],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+1, row0+8, col0+1],
})

# Configure a second series. Note use of alternative syntax to define ranges.
# List is [ sheet_name, first_row, first_col, last_row, last_col ].
chart1.add_series({
    'name':       '%s' % headings[col0+2],
    'categories': ['Sheet1', row0+1 , col0  , row0+8, col0],
    'values':     ['Sheet1', row0+1 , col0+2, row0+8, col0+2],
})

# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Matrix Type: Overall'})
chart1.set_x_axis({'name': 'RPM Groupes by Sellers'})
chart1.set_y_axis({'name': 'RPM ($)'})

# Set an Excel chart style.
chart1.set_style(18)

# Insert the chart into the worksheet (with an offset).

worksheet.insert_chart('D2', chart1, {'x_offset': 600, 'y_offset': 580}) # x=600, y=180

workbook.close()

