# pyrunxlsmacro

This is a simple python3 script to run specific macros in excel macro-enabled workbooks.

## Use case:
Automate the update of database export to excel trough OLEDB link: useful when you have to feed a MS DB to python but your enterprise security rules prevent direct link. The file can then be read with ![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white) for easy manipulation.

## Description
Based on win32com.client.
runMacro function is fed with filename and macro name. 
The xlsm file is open, the macro is run and then file is saved (in case the macro makes some changes to the file itself) and closed
