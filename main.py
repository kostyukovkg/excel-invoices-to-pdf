import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx") # находит и помещает пути всех файлов с нужным расширением в лист
#print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1") # create new python object with extracted data from xlsx
    pdf = FPDF(orientation='P', unit='mm', format='A4') # create new clean pdf object
    pdf.add_page()
    pdf.set_font(family='Times', size=16, style='B')
    # extract filepath name for the title
    filename = Path(filepath).stem # create a filepath object and then extract the stem
    invoice_nr = filename.split('-')[0]
    # create the title
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")
    pdf.output(f'PDFs/{filename}.pdf')
