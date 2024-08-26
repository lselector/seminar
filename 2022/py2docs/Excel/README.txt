
html_from_df.py 
    convert a pandas DF into HTML page 
    and into Excel binary (as io.BytesIO() object)
    and embed Excel binary (base64-encoded) into HTML
    providing a link to download this binary Excel file

html_from_report.py
    read binary excel file, base64-encodes it
    read text html of a pandas DataFrame
    combines everything into one HTML page
    which shows the table with data
    and provides a link to download Excel binary.

style0.css
    used in test_rep_to_html.py

read_excel_into_df.py
    two ways of reading the Excel into Pandas DataFrame

xlsxwriter - good module to create Excel file:
    write_excel_1_simple.py
    write_excel_2.py
    write_excel_3_simple.py
    write_excel_4.py

xlwt - limited simple library:
    write_excel_colors_xlwt.py

