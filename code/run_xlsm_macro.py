#!/usr/bin/env python3
# -*- coding: utf-8-unix -*-

import os, os.path
import win32com.client as wincl
import datetime
from openpyxl import load_workbook  

def runMacro(filename, macro):
    if os.path.exists(filename):
        macropath = filename.split('\\')[-1] + '!' + macro
        # DispatchEx is required in the newest versions of Python.
        excel_macro = wincl.DispatchEx("Excel.application")
        excel_path = os.path.expanduser(filename)
        workbook = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly =1)
        excel_macro.Application.Run(macropath)
        #Save the results in case you have generated data
        workbook.Save()
        excel_macro.Application.Quit()  
        del excel_macro      

filename = 'C:\\Users\your_file.xlsm'
macroname = "Module1.update_db"

print(datetime.datetime.now(), '>> starting update:',filename)       
runMacro(filename, macroname)
print(datetime.datetime.now(), '>> update completed')
