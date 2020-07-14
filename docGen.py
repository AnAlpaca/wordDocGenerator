from __future__ import print_function
import xlrd
import xlwt
import time
import os
import os.path
from mailmerge import MailMerge
from datetime import date
class DocGen:
    
    def __init__(self):
        

        cwd = os.getcwd() 
        self.dir_path = cwd
        self.interface_path = self.dir_path + "\\templates\\interface.xlsm"
    
    def show_path(self):
        print(self.interface_path)

    def open_excel(self):
        os.system('start EXCEL.EXE %s' %(self.interface_path))
        input("Please input all your required information into the fields in the Excel Spreasheet.\nOnce completed please save and close EXCEL.\nPress Enter to generate Production Approval from EXCEL information.")
    
     
n = DocGen()
n.show_path()
n.open_excel()
