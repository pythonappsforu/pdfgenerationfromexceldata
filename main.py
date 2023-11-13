import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    filename = Path(filepath).stem
    invoice_no = filename.split("-")[0]
    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=18,style='B')
    pdf.cell(w=50,h=10,txt=f"Invoice no:{invoice_no}")
    pdf.output(f'PDFs/{filename}.pdf')


