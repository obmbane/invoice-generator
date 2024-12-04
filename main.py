import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#create list of file names
filepaths = glob.glob('excel_templates/*.xlsx')


for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    file_name = Path(filepath).stem
    invoice_num = file_name.split('-')[0]
    Invoice_date = file_name.split('-')[1]
    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=20,style='B')
    pdf.cell(w=60,h=20,txt=f'Invoice Number: {invoice_num}',ln=1)
    pdf.set_font(family='Times',size=14,style='B')
    pdf.cell(w=60,h=14,txt=f'Date: {Invoice_date}')

    pdf_name = filepath.split('/')[1][-4]
    pdf.output(f'PDF_Invoices/{file_name}.pdf')

    #print(df)