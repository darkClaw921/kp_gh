from fpdf import FPDF

pdf = FPDF()
pdf.add_page()
pdf.add_font('calibri', '', 'fonts/calibri.ttf')
pdf.set_font('calibri', size=12)
pdf.cell(text="привет")
pdf.output("hello_world.pdf")