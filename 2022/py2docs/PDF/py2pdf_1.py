
"""
# py2pdf_1.py
#
# simple Hello World example
"""

from fpdf import FPDF

pdf = FPDF()
pdf.add_page()
pdf.set_font('Arial', 'B', 16)
pdf.cell(40, 10, 'Hello World!')
pdf.output('files/junk1.pdf', 'F')

