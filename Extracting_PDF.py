#!/usr/bin/env python
# coding: utf-8

# In[57]:


import PyPDF2 as pp
import glob
import pandas as pd


# In[58]:


class pdf_extraction:
    
    def __init__(self,path):
        self.path=path
        self.file_name=path.split("/")[-1][:-4]
        
    def to_list(self):
        reader = pp.PdfFileReader(open(self.path, 'rb'))
        text=reader.getPage(0).extractText()
        return text.split()
    
    def get_info(self):
        _list= self.to_list()
        self.invoice=''.join(_list[_list.index('INVOICE')+2:_list.index('DATE:')])
        self.date=''.join(_list[_list.index('DATE:')+1:_list.index('TO:')])


# In[60]:


if __name__ == "__main__":
    
    out_put=pd.DataFrame(columns=['FileName','Invoice#','Date'])
    files=glob.glob("*.pdf")

    for path in files:
        extractor=pdf_extraction(path)
        extractor.get_info()
        out_put=out_put.append({'FileName':extractor.file_name, 'Invoice#':extractor.invoice, 'Date':extractor.date}
                               ,ignore_index=True)
    out_put.to_excel("output.xlsx")

