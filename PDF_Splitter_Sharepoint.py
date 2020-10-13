#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import requests_ntlm
import json
import ConfigParser
import urllib
import pandas as pd
import io
import openpyxl
import os
from IPython.display import JSON,HTML


# In[2]:


email = ''
pwd = ''
site = 'https://collab.pncint.net'
location = "/sites/RB/RetailLendingStrategyandPlanning/"
sharepoint_path = "ID1/13066/Remediation"
#file_name = 'IS-10157.xlsx'
certificate = ('/etc/wakari/paeaen02-05.pncint.net.cer', '/etc/wakari/paeaen02-05.pncint.net.key')


# In[3]:


def getFormDigest(uid, pwd, site, location):
    headers = {'Content-Type': 'application/json; odata=verbose', 'accept': 'application/json;odata=verbose'}    
    r = requests.post(site + location + "/_api/contextinfo",auth=requests_ntlm.HttpNtlmAuth(uid, pwd), headers=headers)
    print r
    print("Connection Confirmed")
    return r.json()['d']['GetContextWebInformation']['FormDigestValue']


# In[4]:


getFormDigest(email, pwd, site, location)


# In[5]:


def uploadFileToSharepoint(uid, password, site, location, path, fileName):
    try:
        url = site + location + '_api/web/GetFolderByServerRelativeUrl(\'{}\')/Files/add(url=\'{}\',overwrite=true)'            .format(urllib.quote(location + path), fileName)
        print url

        file = open(fileName, 'rb')

        fd = getFormDigest(uid, password, site, location)
        print "Acquired Form Digest"

        headers = {'Content-Type': 'application/json; odata=verbose',  'accept': 'application/json;odata=verbose', 'x-requestdigest' : fd}
        result = requests.post(url, cert=certificate, auth=requests_ntlm.HttpNtlmAuth(uid, password), headers=headers, data=file.read())
        
        if result.ok:
            print "Successfully uploaded"
        else:
            print "Unsuccessful upload. Returned status code {}".format(result.status_code)
            if result.status_code == 400:
                print "Sharepoint has maximum URL length that you may have exceeded. If the URL can't be shortened, you will need to manually upload"
    
    except (IOError):
        print("Wrong file or file path")


# In[6]:


#form = uploadFileToSharepoint(email, pwd, site, location, sharepoint_path, file_name)


# In[7]:


import os
import pandas as pd
import pyodbc
import ConfigParser

from pandas import Series, DataFrame


# In[8]:


#Read in excel 
excel = pd.read_excel('Copy of IS13066_CheckFile_1036_092520_Split.xlsx')#insert excel file name here
print("The excel contains " + str(len(excel)) + " records.")


# In[9]:


for col in excel:
    print col


# In[10]:


#Put loan numbers into a list to be used in next function for naming PDFs. 
loanNumber = excel['LOAN_NUMBER'].tolist()
print(loanNumber)


# In[11]:


from PyPDF2 import PdfFileReader, PdfFileWriter

def pdf_splitter(path):
    fname = os.path.splitext(os.path.basename(path))[0]#grab the name of the input file, minus the extension.
    counter = 0
    pdf = PdfFileReader(path)  
    for page in range(pdf.getNumPages()): 
        pdf_writer = PdfFileWriter()   
        pdf_writer.addPage(pdf.getPage(page))   
        output_filename = '{}.pdf'.format(loanNumber[page])
        
        with open(output_filename, 'wb+') as out:  
            pdf_writer.write(out)
            #Exit loop
        form = uploadFileToSharepoint(email, pwd, site, location, sharepoint_path, output_filename)
            
        print('Created: {}'.format(output_filename))
        counter += 1
    print ("\n"+ str(counter) + " PDFs split")
    print ("PDF Splitting completed.")
if __name__ == '__main__':
    path = 'IS13066 Refund Check 092820.pdf'#insert pdf file name here. 
    pdf_splitter(path)

    

