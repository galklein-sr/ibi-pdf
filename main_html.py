import os
import pdfkit
from jinja2 import Template
from PyPDF2 import PdfMerger

config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")

pages = [
    {'title': 'Sales Overview', 'value': 20000, 'desc': 'Total Sales'},
    {'title': 'Expenses', 'value': 7500, 'desc': 'Total Expenses'},
    {'title': 'Profit', 'value': 12500, 'desc': 'Net Profit'},
]



os.makedirs('outputs', exist_ok=True)
pdf_files = []
for idx, page in enumerate(pages):
    html_template=""
    with open("templates/template1.html", "r", encoding='utf-8') as file:
        html_template= file.read()
    html_template = html_template.replace("@title", page.get('title'))
    html_content = Template(html_template).render(**page)
    html_path = f'outputs/page_{idx+1}.html'
    pdf_path = f'outputs/page_{idx+1}.pdf'
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    # Generate PDF with pdfkit
    options = {
        'page-size': 'A4',
        'orientation': 'Landscape',
        'encoding': 'UTF-8',
        'margin-top': '5mm',
        'margin-bottom': '5mm',
        'margin-left': '5mm',
        'margin-right': '5mm'
    }
    pdfkit.from_file(html_path, pdf_path, configuration=config, options=options)
    pdf_files.append(pdf_path)

#merger = PdfMerger()
#for pdf in pdf_files:
#    merger.append(pdf)
#final_pdf_path = 'outputs/PowerBI_Report.pdf'
#merger.write(final_pdf_path)
#merger.close()

#print(f"Combined PDF saved as {final_pdf_path}")
