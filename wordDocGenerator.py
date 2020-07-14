from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import xlrd
import xlwt
import time
import os

def openExcelx():
    import os.path
    import os
    cwd = os.getcwd() + "\\templates\\interface.xlsm"
    os.system('start EXCEL.EXE %s' %(cwd))
    input("Please input all your required information into the fields in the Excel Spreasheet.\nOnce completed please save and close EXCEL.\nPress Enter to generate Production Approval from EXCEL information.")
    return cwd

x = cwd

def readExcelFile(cwd):
    workbook = xlrd.open_workbook(cwd)
    worksheet = workbook.sheet_by_index(1)

    partFullNameGen = worksheet.cell(4, 1).value
    lastUpdated = '{:%d-%b-%Y}'.format(date.today())
    partNumber= worksheet.cell(5, 1).value
    colourPartNumber = list(partNumber)
    supplierName= worksheet.cell(6, 1).value
    authorName = worksheet.cell(7, 1).value
    partShortName = worksheet.cell(9, 1).value
    machineClampForce = str(worksheet.cell(10, 1).value)
    barrelCapacity = str(worksheet.cell(16, 1).value)
    return (barrelCapacity)

openExcelx()
print(x)