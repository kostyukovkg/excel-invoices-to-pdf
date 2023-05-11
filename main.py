import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")  # находит и помещает пути всех файлов с нужным расширением в лист
# print(filepaths)

for filepath in filepaths:
    # create new clean pdf object and add page
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # extract filepath name for the file
    filename = Path(filepath).stem  # create a filepath object and then extract the stem
    invoice_nr, date = filename.split('-')

    # create the title
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    # add date
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # creates a new pandas object with extracted data from xlsx
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # add header to the table
    columns = list(df.columns) # df.columns - index type. need to convert into list
    columns = [item.replace('_', " ").title() for item in columns]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(1)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=20, h=8, txt=columns[4], border=1, ln=1)

    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=20, h=8, txt=str(row['total_price']), border=1, ln=1)

    # add total sum into the table
    total_sum = df['total_price'].sum()
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=70, h=8, txt='', border=1)
    pdf.cell(w=40, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=20, h=8, txt=f'{total_sum}', border=1, ln=1)

    # add total sum sentence
    pdf.set_font(family='Times', size=14, style='B')
    pdf.set_text_color(1)
    pdf.cell(w=30, h=8, txt=f'Total price of your order is {total_sum}', ln=1)

    # add company name and logo
    pdf.cell(w=50, h=12, txt=f'KKostyukov Co. Ltd', align='center')
    pdf.image('mylogo.png', w=10)


    pdf.output(f'PDFs/{filename}.pdf')
