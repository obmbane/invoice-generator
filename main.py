import pandas as pd
import glob

#create list of file names
filepaths = glob.glob('excel_templates/*.xlsx')
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    print(df)