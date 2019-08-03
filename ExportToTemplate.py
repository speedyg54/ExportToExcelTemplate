#.py
# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 07:06:24 2019

@author: OBar

Script Purpose: Create a script to output a DataFrame to a preformatted excel template 
"""

import pandas as pd
import datetime
import time
import numpy as np
import win32com.client as win32

class ETT():
    
    def __init__(self):
        self.workbook = None
        
    def open_excel_template(self, temppath):
        print("Template Path Is: {}".format(temppath))
        ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
        ExcelApp.Visible = True
        self.workbook = ExcelApp.Workbooks.Open("{}".format(temppath))
        
        return 0
    
    def push_results(self, DataFram, sheet, row, col):
        time.sleep(5) # wait for the template to open
        #first create simple shortcuts to work from
        shet = self.workbook.Sheets(sheet)
        
        colwrk = shet.Range(shet.Cells(row-1, col),
                            shet.Cells(row-1,
                                       col+len(DataFram.columns)-1)
                            )
        datwrk = shet.Range(shet.Cells(row, col),
                            shet.Cells(row+len(DataFram.index)-1,
                                       col+len(DataFram.columns)-1)
                            )
        
        #push columns to workbook
        colwrk.Value = DataFram.columns.values.tolist() #outputs c-contiguous array as list
        
        #push data to workbook
        datwrk.Value = DataFram.values.tolist()
        
        """
        This could be an area for formatting the results
        """
        return 0
    
    def save_close_template(self, outpath):
        self.workbook.SaveAs(outpath + "- {}.xlsx"
                             .format(datetime.datetime.now().strftime('%d-%m-%y %H-%M-%S'))
                             )
        self.workbook.Close(True)
        ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
        ExcelApp.Visible = False
        return 0
    
if __name__ == "__main__":
    
    x = ETT()
    
    
    a=np.ones((3,3))
    b=np.eye(3)
    c=np.concatenate((a,b))
    dfrm = pd.DataFrame(c)
    x.open_excel_template("C:\\Users\\OBar\\Documents\\Reusable Python Scripts\\New Template.xlsx")
    
    x.push_results(DataFram=dfrm, sheet='Sheet1', row=2, col=2)
    
    x.save_close_template("C:\\Users\\OBar\\Documents\\Reusable Python Scripts\\Output Template")

        
    
    
    
    
    
    
    