from fpdf import Template

# define the ELEMENTS that will compose the template. 

elements = [

    {'name': 'company_logo', 'type': 'I', 'x1': 20.0, 'y1': 17.0, 'x2': 78.0, 'y2': 30.0, 
     'font': None, 'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 
     'background': 0, 'align': 'I', 'text': 'logo', 'priority': 2, },

    {'name': 'company_name', 'type': 'T', 'x1': 17.0, 'y1': 32.5, 'x2': 115.0, 'y2': 37.5, 
     'font': 'Arial', 'size': 12.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 
     'background': 0, 'align': 'I', 'text': '', 'priority': 2, },

    {'name': 'box', 'type': 'B', 'x1': 15.0, 'y1': 15.0, 'x2': 185.0, 'y2': 260.0, 
     'font': 'Arial', 'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 
     'background': 0, 'align': 'I', 'text': None, 'priority': 0, },

    {'name': 'box_x', 'type': 'B', 'x1': 95.0, 'y1': 15.0, 'x2': 105.0, 'y2': 25.0, 
     'font': 'Arial', 'size': 0.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 
     'background': 0, 'align': 'I', 'text': None, 'priority': 2, },

    {'name': 'line1', 'type': 'L', 'x1': 100.0, 'y1': 25.0, 'x2': 100.0, 'y2': 57.0, 
     'font': 'Arial', 'size': 0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 
     'background': 0, 'align': 'I', 'text': None, 'priority': 3, },

    {'name': 'barcode', 'type': 'BC', 'x1': 20.0, 'y1': 246.5, 'x2': 140.0, 'y2': 254.0, 
     'font': 'Interleaved 2of5 NT', 'size': 0.75, 'bold': 0, 'italic': 0, 'underline': 0, 
     'foreground': 0, 'background': 0, 'align': 'I', 
     'text': '200000000001000159053338016581200810081', 'priority': 3, },
]

# instantiate the template and define the HEADER
f = Template(format="A4", elements=elements,
             title="Sample Invoice")
f.add_page()

# fill some of the fields of the template with the information we want
# note: we access the elements treating the template instance as a "dict"
f["company_name"] = "Sample Company"
f["company_logo"] = "files/mylogo.png"

# and now we render the page
f.render("files/junbk4.pdf")

