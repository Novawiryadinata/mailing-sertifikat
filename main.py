from docxtpl import DocxTemplate
# pip install docxtpl
import pandas as pd
# pip install pandas

import time
import os
# pip install os-sys

# initialization of MS Word
from win32com import client
word_app = client.Dispatch("Word.Application")

# data_frame = pd.read_csv('sumber.csv')
data_frame = pd.read_excel('sumber.xlsx')

for r_index, row in data_frame.iterrows():
    cust_name = row['Name']
    cust_Jabatan = row['Jabatan']

    tpl = DocxTemplate("target.docx")
    df_to_doct = data_frame.to_dict()
    x = data_frame.to_dict(orient='records')
    context = x
    tpl.render(context[r_index])
    tpl.save('Doc\\'+cust_name+".docx")

    # Mencari Folder Path
    time.sleep(1)
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    # print(ROOT_DIR)

    # Doc to PDF
    doc = word_app.Documents.Open(ROOT_DIR+'\\Doc\\'+cust_name+'.docx')
    print('Name File : '+cust_name+'.pdf    ')
    doc.SaveAs(ROOT_DIR+'\\PDF\\'+ cust_name + '.pdf', FileFormat=17)
print(ROOT_DIR)
word_app.Quit()