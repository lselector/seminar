

from fpdf import FPDF

# --------------------------------------------------------------
class PDF(FPDF):
    def header(self):
        self.image('files/mylogo.png', 10, 8, 33)   # Logo
        self.set_font('Arial', 'B', 15)             # Arial bold 15
        self.cell(80)                               # Move to the right
        self.cell(30, 10, 'Title', 1, 0, 'C')       # Title
        self.ln(20)                                 # Line break

    def footer(self):
        self.set_y(-15)                 # Position at 1.5 cm from bottom
        self.set_font('Arial', 'I', 8)  # Arial italic 8
                                        # Page number
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

# --------------------------------------------------------------
pdf = PDF()

pdf.alias_nb_pages()
pdf.add_page()
pdf.set_font('Times', '', 12)

for i in range(1, 41):
    pdf.cell(0, 10, 'Printing line number ' + str(i), 0, 1)

pdf.output('files/junk2.pdf', 'F')

