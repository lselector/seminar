
"""
# wrie_excel_colors_xlwt.py
# writing to files/xlwt_colors.xls
# using xlwt.Style.colour_map
"""

# pip install xlwt
import xlwt

wb = xlwt.Workbook()
ws = wb.add_sheet("Colors")

cc = xlwt.Style.colour_map # dictionary of 62 colors. keys are strings, values - integers

colors62 = sorted(cc.keys())

font_colors = colors62
back_colors = colors62

# 62*62 = 3844 different styles.
# Excel can not have more than ~500 styles.
# When you try to create more, it shows correctly first ~500, and the rest shows black font.
# That's why for purposes of creating a useful (but limited) pallette, you can choose
# to limit number of styles by combining only 18 dark font colors with 21 light background colors.

use_few_colors = True
if use_few_colors:
    font_colors = sorted(['black', 'blue', 'brown', 'dark_blue', 'dark_green', 'dark_purple',
        'dark_red', 'dark_teal', 'dark_yellow', 'grey50', 'grey80', 'indigo', 'light_blue',
        'magenta_ega', 'olive_green', 'purple_ega', 'teal', 'violet'])
    back_colors = sorted(['aqua', 'coral', 'cyan_ega', 'gold', 'gray25', 'green', 'ice_blue',
        'ivory', 'lavender', 'light_green', 'light_orange', 'light_turquoise', 'light_yellow', 
        'lime', 'magenta_ega', 'pale_blue', 'pink', 'sky_blue', 'tan', 'white', 'yellow'])

for row in range(len(back_colors)):
    for col in range(len(font_colors)):
        ss = 'pattern: pattern solid, fore_colour %s; font: bold True, color_index %s;' % (back_colors[row],cc[font_colors[col]]) 
        some_style = xlwt.easyxf(ss)
        ws.write(row, col, "%s/%s" % (font_colors[col],back_colors[row]), some_style)

fname = 'files/xlwt_colors.xls'
print(f"writing file {fname}")
wb.save(fname)


