#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import pandas as pd
import pyodbc
#import ConfigParser
from pandas import Series, DataFrame

outputPath = 'C:\output'


# In[2]:


excel = pd.read_excel('Loans.xlsx')
excel.head()


# In[3]:


loanNumber = excel.LOAN_NUMBER.tolist()
print(loanNumber)


# In[5]:


from PyPDF2 import PdfFileReader, PdfFileWriter
outputPath = 'C:\\output'
def pdf_splitter(path):
    fname = os.path.splitext(os.path.basename(path))[0]
    counter = 0
    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        output_filename = '{}_Page Number_{}.pdf'.format(fname, loanNumber[page])
        
        with open(os.path.join(outputPath,output_filename), 'wb') as out:
            pdf_writer.write(out)
            
        print('Created: {}'.format(output_filename))
        counter += 1
    print (counter)
if __name__ == '__main__':
    path = 'syllabusCIT244.pdf'
    pdf_splitter(path)


# In[ ]:




