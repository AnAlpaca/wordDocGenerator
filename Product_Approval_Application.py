from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import xlrd
import xlwt
import os.path
import time
import os



os.system("start EXCEL.EXE Component_Approval.xlsm")
input("Please input all your required information into the fields in the Excel Spreasheet.\nOnce completed please save and close EXCEL.\nPress Enter to generate Production Approval from EXCEL information.")



workbook = xlrd.open_workbook('Component_Approval.xlsm')
worksheet = workbook.sheet_by_index(0)
# Value of 1st row and 1st column
# values are read as follows worksheet.cell(0, 0).value

# Defines variable to use else where in the program and to pass through to the merge funtion.
# This allows us to use other code to define names and values such as document numbers and pass them through to other elements of the code.

#                  Main Information               # 
###################################################
partFullNameGen = worksheet.cell(4, 1).value
lastUpdated = '{:%d-%b-%Y}'.format(date.today())
partNumber= worksheet.cell(5, 1).value
colourPartNumber = list(partNumber)
supplierName= worksheet.cell(6, 1).value
authorName = worksheet.cell(7, 1).value
partShortName = worksheet.cell(9, 1).value
machineClampForce = str(int(worksheet.cell(10, 1).value))
barrelCapacity = str(int(worksheet.cell(11, 1).value))
###################################################


# Checks if the Part has Multiple Colours
if worksheet.cell(4, 2).value != None:     
    colourString = str(worksheet.cell(4, 2).value)
    colourList = colourString.split('; ',100)
    colourPartNumber = [partNumber[:7] + x + partNumber[10:] for x in colourList]
    colourPartName = [None] * len(colourList)
    for x in range(len(colourList)):
        colourPartName[x] = partFullNameGen.replace('XXX', colourList[x])
    generateColourRows = 1 # variable to say to create rows
    print(colourPartName)

else:
    pass

#Checks if the Parts has multiple various and therefore part numbers
if worksheet.cell(5, 2).value != None:
    multiPartNumberString = str(worksheet.cell(5, 2).value)
    multiPartNumberList = multiPartNumberString.split('; ',100)
    fullPartNumber = [None] * len( multiPartNumberList)
    for x in range(len( multiPartNumberList)):
        fullPartNumber[x] = [((colourPartNumber[x])[:5]) + multiPartNumberList[x] + ((colourPartNumber[x])[6:])]
         
    generateRows = 1 # variable to say to create rows
 

else:
    pass



# Start of the Merge Program
# Define the templates - assumes they are in the same directory as the code
Main_PA_Template = "PRODUCT APPROVAL TEMPLATE - MAIN.docx"
documentName = 'PA-' + partFullNameGen +'.docx'
# Points to the document template and prints the variables in the document that can be editted.
Main_PA_Doc = MailMerge(Main_PA_Template)
# Reads the variables in the document and creates a list
documentVariablesSet = Main_PA_Doc.get_merge_fields()
documentVariablesList = list(documentVariablesSet)
lengthlist = len(documentVariablesList)
print('The following variables are used in the Word Templates', documentVariablesList)

# writes a new spreadsheet and puts all the variables from the template into column one
from xlwt import Workbook
workbook = Workbook() 
sheet = workbook.add_sheet('Sheet_1')

for x in range(lengthlist):
     sheet.write(x, 0, documentVariablesList[x])

workbook.save('variables_log.xlsx')

# importing data from the excel spreadsheet


# Merge in the values
Main_PA_Doc.merge(
    PartFullName = partFullNameGen,
    LastUpdated = lastUpdated,
    PartNumber= partNumber,
    SupplierName= supplierName,
    AuthorName = authorName,
    PartShortName = partShortName,
    MachineClampForce = machineClampForce,
    BarrelCapacity = barrelCapacity)

# Save the document as example 1
Main_PA_Doc.write(documentName)
#End of the Merge Program
file_path = documentName
time_to_wait = 10
time_counter = 0
while not os.path.exists(file_path):
    time.sleep(1)
    time_counter += 1
    if time_counter > time_to_wait:
        print('File could not be saved, contact Calvin.')
        break
print(documentName, ' has been written successfully.')
 
