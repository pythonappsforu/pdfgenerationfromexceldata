import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    filename = Path(filepath).stem
    invoice_no = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()

    pdf.set_font(family='Times',size=18,style='B')
    pdf.cell(w=50,h=10,txt=f"Invoice no:{invoice_no}",ln=1)

    pdf.set_font(family='Times',size=18,style='B')
    pdf.cell(w=50,h=10,txt=f"Date:{date}",ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    columns = df.columns
    columns = [column.replace('_',' ').capitalize() for column in columns]
    total_price = df['total_price'].sum()

    pdf.set_font(family='Times', size=12,style='B')
    pdf.cell(w=20, h=10, txt=columns[0], border=1)
    pdf.cell(w=60, h=10, txt=columns[1], border=1)
    pdf.cell(w=40, h=10, txt=columns[2], border=1)
    pdf.cell(w=30, h=10, txt=columns[3], border=1)
    pdf.cell(w=30, h=10, txt=columns[4], border=1, ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family='Times', size=12)
        pdf.cell(w=20, h=10, txt=str(row['product_id']),border=1)
        pdf.cell(w=60, h=10, txt=str(row['product_name']),border=1)
        pdf.cell(w=40, h=10, txt=str(row['amount_purchased']),border=1)
        pdf.cell(w=30, h=10, txt=str(row['price_per_unit']),border=1)
        pdf.cell(w=30, h=10, txt=str(row['total_price']),border=1,ln=1)

    pdf.cell(w=20, h=10, txt="", border=1)
    pdf.cell(w=60, h=10, txt="", border=1)
    pdf.cell(w=40, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=f"{total_price}", border=1, ln=1)

    pdf.set_font(family='Times', size=14,style='B')
    pdf.cell(w=20, h=10, txt=f"Total price is {total_price}", border=0,ln=1)
    pdf.cell(w=20, h=10, txt="FebAsh", border=0)
    pdf.image("logo.png",w=20)

    pdf.output(f'PDFs/{filename}.pdf')


