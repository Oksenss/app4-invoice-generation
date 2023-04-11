import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
#we need to also install the library openpyxl else we will get an error

filepaths = glob.glob('invoices/*.xlsx')
print(filepaths)

for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    #stem is used for getting the name of the file
    #without extension like.xlsx
    invoice_nr, date = filename.split('-')
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', ln=1)
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=16, txt=f'Date: {date}', ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    #add a header
    header = list(df.columns)
    header = [item.replace("_", " ").title() for item in header]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=8, txt=header[0], border=1 )
    pdf.cell(w=70, h=8, txt=header[1], border=1)
    pdf.cell(w=30, h=8, txt=header[2], border=1)
    pdf.cell(w=30, h=8, txt=header[3], border=1)
    pdf.cell(w=30, h=8, txt=header[4], border=1, ln=1)
    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    total_price = df['total_price'].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)
    #Add total sum sentence
    pdf.set_font(family='Times', size=10)
    pdf.cell(w = 30, h=8, txt=f'The total price is {total_price}', ln=1)
    #Add company and logo
    pdf.set_font(family='Times', size=10)
    pdf.cell(w = 30, h=8, txt=f'The company logo')
    pdf.image('pythonhow.png',w=10)


    pdf.output(f'PDFs/{filename}.pdf')