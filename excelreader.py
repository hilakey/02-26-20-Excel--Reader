# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from xlrd import open_workbook
import pprint
import xlrd

book = open_workbook('studentData.xlsx', on_demand=True)

 
print("How many sheets does my file have? There is", book.nsheets, " sheet")

#importing specific sheet's data into my program
sheet = book.sheet_by_name('Sheet1')
#creating an empty list
temp = []



#creating an empty dictionary
StudentInfo = {}

for r in range(sheet.nrows):
       StudentInfo[sheet.cell_value(r,2)] = {} #creating a new dictionary within the value "username"
       StudentInfo[sheet.cell_value(r,2)]["Name:"] = sheet.cell_value(r, 0)
       StudentInfo[sheet.cell_value(r,2)]["Email:"] = sheet.cell_value(r, 1)
       cell = sheet.cell(r,3)

        #converting float value to datetime in a string format
       if(cell.ctype) == xlrd.XL_CELL_DATE:
           studentBirthday = xlrd.xldate.xldate_as_datetime(cell.value ,book.datemode)
           studentBirthday = studentBirthday.strftime("%Y/%m/%d")
           StudentInfo[sheet.cell_value(r,2)]["Birthday:"] = studentBirthday
           
       StudentInfo[sheet.cell_value(r,2)]["UTEP ID:"] = int(sheet.cell_value(r,4))
        


pprint.pprint(StudentInfo)


#creating an empty list
temp = []
#going through each filled cell in excel sheet

#print cell value
