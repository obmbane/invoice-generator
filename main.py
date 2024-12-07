import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

#create list of file names
filepaths = glob.glob('excel_templates/*.xlsx')
company_name = 'Recon Pty (Ltd)'
company_address = ''' Samrand Office Park
32 XYZ Street
Midrand
2009
'''


for filepath in filepaths:
   
    #Read in data

    file_name = Path(filepath).stem
    invoice_num, Invoice_date = file_name.split('-')

    #Create PDF

    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=20,style='B')        # Add Company Information
    pdf.cell(w=60,h=8,txt=company_name,ln=0)
    pdf.image('pythonhow.png',w=9)
    pdf.cell(w=0,h=8,txt='',ln=1)
    pdf.set_font(family='Times',size=12,style='')
    pdf.multi_cell(w=60,h=8,txt=company_address)
    pdf.cell(w=0,h=8,txt='',ln=1)


    pdf.set_font(family='Times',size=14,style='B')        #add invoice number
    pdf.cell(w=60,h=8,txt=f'Invoice Number: {invoice_num}',ln=1)

    pdf.set_font(family='Times',size=16,style='B')       #add invoice date
    pdf.cell(w=60,h=10,txt=f'Date: {Invoice_date}',ln=2)

    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    column_list = list(df.columns)
    new_column_list = [item.replace('_',' ') for item in column_list]

    pdf.set_font(family='Times',size=12,style='B')       #add invoice date
    pdf.cell(w=30,h=8,txt= new_column_list[0].title(), border=1,ln=0)
    pdf.cell(w=60,h=8,txt= new_column_list[1].title(), border=1,ln=0)
    pdf.cell(w=40,h=8,txt= new_column_list[2].title(), border=1,ln=0)
    pdf.cell(w=30,h=8,txt= new_column_list[3].title(), border=1,ln=0)
    pdf.cell(w=30,h=8,txt= new_column_list[4].title(), border=1,ln=1)

    
    for index, row in df.iterrows():
            
            pdf.set_font(family='Times',size=12,style='')       #add invoice date
            pdf.cell(w=30,h=8,txt=str(row['product_id']), border=1,ln=0)
            pdf.cell(w=60,h=8,txt=str(row['product_name']), border=1,ln=0)
            pdf.cell(w=40,h=8,txt=str(row['amount_purchased']), border=1,ln=0)
            pdf.cell(w=30,h=8,txt=f'R {str(row['price_per_unit'])}', border=1,ln=0)
            pdf.cell(w=30,h=8,txt=f'R {str(row['total_price'])}', border=1,ln=1)

    total = df['total_price'].sum()
    pdf.cell(w=130,h=8,txt='', border=0,ln=0)
    pdf.set_font(family='Times',size=11,style='B')
    pdf.cell(w=30,h=8,txt='Sub-total', border=0,ln=0)
    pdf.cell(w=30,h=8,txt=f'R {str(total)}', border=1,ln=1)

    pdf.cell(w=130,h=8,txt='', border=0,ln=0)
    pdf.set_font(family='Times',size=11,style='B')
    pdf.cell(w=30,h=8,txt='VAT', border=0,ln=0)
    pdf.cell(w=30,h=8,txt=f'{str(15)} %', border=1,ln=1)

    pdf.cell(w=130,h=8,txt='', border=0,ln=0)
    pdf.set_font(family='Times',size=11,style='B')
    pdf.cell(w=30,h=8,txt='Total', border=0,ln=0)
    pdf.cell(w=30,h=8,txt=f'R {str(round(total*1.15,2))}', border=1)

    
    #Save PDF

    pdf_name = filepath.split('/')[1][-4]
    pdf.output(f'PDF_Invoices/{file_name}.pdf')
