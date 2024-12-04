import pandas as pd
import glob
from fpdf import FPDF

#create list of file names
filepaths = glob.glob('excel_templates/*.xlsx')
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    invoice_list = filepath.split('-')
    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=20,style='B')
    pdf.cell(w=60,h=20,txt=f'Invoice Number: {invoice_list[0]}')
    pdf.output(f'{filepath}.pdf')

    #print(df)