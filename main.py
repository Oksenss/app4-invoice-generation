import pandas as pd
import glob
#we need to also install the library openpyxl else we will get an error

filepaths = glob.glob('invoices/*.xlsx')
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    print(df)