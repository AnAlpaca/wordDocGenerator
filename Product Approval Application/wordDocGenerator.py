from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import xlrd
import xlwt
import time



def openExcelFile():
    import os.path
    import os
    os.system("start EXCEL.EXE Component_Approval.xlsm")
    input("Please input all your required information into the fields in the Excel Spreasheet.\nOnce completed please save and close EXCEL.\nPress Enter to generate Production Approval from EXCEL information.")
    workbook = xlrd.open_workbook('Component_Approval.xlsm')
    worksheet = workbook.sheet_by_index(1)
    return None

def readExcelFile():
    workbook = xlrd.open_workbook('Component_Approval.xlsm')
    worksheet = workbook.sheet_by_index(1)

    partFullNameGen = worksheet.cell(4, 1).value
    lastUpdated = '{:%d-%b-%Y}'.format(date.today())
    partNumber= worksheet.cell(5, 1).value
    colourPartNumber = list(partNumber)
    supplierName= worksheet.cell(6, 1).value
    authorName = worksheet.cell(7, 1).value
    partShortName = worksheet.cell(9, 1).value
    machineClampForce = str(worksheet.cell(10, 1).value)
    barrelCapacity = str(worksheet.cell(11, 1).value)
    return (barrelCapacity)


openExcelFile()
readExcelFile()
