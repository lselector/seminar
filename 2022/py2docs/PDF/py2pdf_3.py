from fpdf import FPDF

title = '20000 Leagues Under the Seas'

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        w = self.get_string_width(title) + 6
        self.set_x((210 - w) / 2)
        self.set_draw_color(0, 80, 180) 
        self.set_fill_color(230, 230, 0)
        self.set_text_color(220, 50, 50)
        self.set_line_width(1)
        self.cell(w, 9, title, 1, 1, 'C', 1)
        self.ln(10)   # line break

    def footer(self):
        self.set_y(-15)         # Position at 1.5 cm from bottom
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128)
        self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

    def chapter_title(self, num, label):
        self.set_font('Arial', '', 12)
        self.set_fill_color(200, 220, 255)
        self.cell(0, 6, 'Chapter %d : %s' % (num, label), 0, 1, 'L', 1)
        self.ln(4)   # Line break

    def chapter_body(self):
        txt = ""
        for ii in range(40):
            txt += f"text line {ii}\n"
        self.set_font('Times', '', 12)
        self.multi_cell(0, 5, txt)         # Read text file
        self.ln()         # Line break
        self.set_font('', 'I')
        self.cell(0, 5, '(end of excerpt)')

    def print_chapter(self, num, title):
        self.add_page()
        self.chapter_title(num, title)
        self.chapter_body()

pdf = PDF()
pdf.set_title(title)
pdf.set_author('Jules Verne')
pdf.print_chapter(1, 'A RUNAWAY REEF')
pdf.print_chapter(2, 'THE PROS AND CONS')
pdf.output('files/junk3.pdf', 'F')

