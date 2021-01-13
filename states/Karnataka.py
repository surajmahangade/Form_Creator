from states import logging,monthdict,Statefolder
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk
from functools import partial
import os
from pathlib import Path
import pandas as pd
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Alignment, Side
import calendar
import logging
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Alignment, Side, PatternFill, numbers


def Karnataka(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    karnatakafilespath = os.path.join(Statefolder,'Karnataka')
    logging.info('karnataka files path is :'+str(karnatakafilespath))
    data.reset_index(drop=True, inplace=True)

    month_num = monthdict[month]

    def create_form_A():

        formAfilepath = os.path.join(karnatakafilespath,'FormA.xlsx')
        formAfile = load_workbook(filename=formAfilepath)
        logging.info('Form A file has sheet: '+str(formAfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_formA = data.drop_duplicates(subset=['Employee Code']).copy()

        data_formA.fillna(value=0, inplace=True)

        data_formA['S.no'] = list(range(1,len(data_formA)+1))

        #data_formA["Nationality"] =''
        #data_formA['education level'] = ''
        #data_formA['Category address'] = ''
        #data_formA['type of employment'] = ''
        #data_formA['lwf'] = ''
        #data_formA['Service Book No'] = ''
        data_formA["a"] = ''
        data_formA["b"] = ''
        data_formA["c"] = ''
        formA_columns = ["S.no",'Employee Code','Employee Name','Unit','Location',"Gender","Father's Name",'Date of Birth',"Nationality","Education Level",'Date Joined','Designation','Category Address',"Type of Employment",'Mobile Tel No.','UAN Number',"PAN Number","ESIC Number","LWF EE","Aadhar Number","Bank A/c Number","Bank Name","Account Code","P","L","Service Book No","Date Left","Reason for Leaving","Identification mark","a","b","c"]
        formA_data = data_formA[formA_columns]


        formAsheet = formAfile['FORM A']

        formAsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form A is ready')

        
        rows = dataframe_to_rows(formA_data, index=False, header=False)

        logging.info('rows taken out from data')

        

        for r_idx, row in enumerate(rows, 13):
            for c_idx, value in enumerate(row, 1):
                formAsheet.cell(row=r_idx, column=c_idx, value=value)
                formAsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Bell MT', size =10)
                formAsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formAsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)
                if c_idx==15 or c_idx==16 or c_idx==20 or c_idx==21:
                    formAsheet.cell(row=r_idx, column=c_idx).number_format= numbers.FORMAT_NUMBER

        logging.info('')


        if str(data_formA['Company Name'].dtype)[0:3] != 'obj':
            data_formA['Company Name'] = data_formA['Company Name'].astype(str)

        if str(data_formA['Company Address'].dtype)[0:3] != 'obj':
            data_formA['Company Address'] = data_formA['Company Address'].astype(str)

        if str(data_formA['Contractor_name'].dtype)[0:3] != 'obj':
            data_formA['Contractor_name'] = data_formA['Contractor_name'].astype(str)

        if str(data_formA['Contractor_Address'].dtype)[0:3] != 'obj':
            data_formA['Contractor_Address'] = data_formA['Contractor_Address'].astype(str)

        if str(data_formA['Unit'].dtype)[0:3] != 'obj':
            data_formA['Unit'] = data_formA['Unit'].astype(str)

        if str(data_formA['Branch'].dtype)[0:3] != 'obj':
            data_formA['Branch'] = data_formA['Branch'].astype(str)

        establishment = formAsheet['L6'].value
        if data_formA['PE_or_contract'].unique()[0] == 'PE':
            L6_data = establishment+' '+data_formA['Company Name'].unique()[0] +', '+data_formA['Company Address'].unique()[0]  
        else:    
            L6_data = establishment+' '+data_formA['Contractor_name'].unique()[0]+', '+data_formA['Contractor_Address'].unique()[0]
        formAsheet['L6'] = L6_data


        company = formAsheet['A10'].value
        A10_data = company+' '+data_formA['Unit'].unique()[0]+', '+data_formA['Branch'].unique()[0]
        formAsheet['A10'] = A10_data

        
        formAfinalfile = os.path.join(filelocation,'FormA.xlsx')
        logging.info('Form A file is' +str(formAfinalfile))
        formAfile.save(filename=formAfinalfile)
        

    def create_form_B():
        formBfilepath = os.path.join(karnatakafilespath,'FormB.xlsx')
        formBfile = load_workbook(filename=formBfilepath)
        logging.info('Form B file has sheet: '+str(formBfile.sheetnames))

        
        logging.info('create columns which are now available')


        data_formB = data.drop_duplicates(subset=['Employee Code']).copy()
        data_formB.fillna(value=0, inplace=True)

        #data_formB['OT hours'] = 0
        #data_formB['Pay OT'] = 0
        if str(data_formB['Earned Basic'].dtype)[0:3] != 'int':
            data_formB['Earned Basic']= data_formB['Earned Basic'].astype(int)
        if str(data_formB['DA'].dtype)[0:3] != 'int':
            data_formB['DA']= data_formB['DA'].astype(int)
        data_formB['basic_and_allo'] = data_formB['Earned Basic']+ data_formB['DA']
        #data_formB['Other EAR'] = data_formB['Other Reimb']+data_formB['Arrears']+data_formB['Other Earning']+data_formB['Variable Pay']+data_formB['Stipend'] +data_formB['Consultancy Fees']
        #data_formB['VPF']=0
        data_formB['Society']="---"
        data_formB['Income Tax']="---"
        if str(data_formB['Other Deduction'].dtype)[0:3] != 'int':
            data_formB['Other Deduction']= data_formB['Other Deduction'].astype(int)
        if str(data_formB['Other Deduction'].dtype)[0:3] != 'int':
            data_formB['Other Deduction']= data_formB['Other Deduction'].astype(int)
        data_formB['Other Deduc']= data_formB['Other Deduction']+ data_formB['Salary Advance']
        data_formB['EMP PF'] = data_formB['PF']
        #data_formB['BankID'] = ''
        #data_formB['Pay Date'] = ''
        data_formB['Remarks'] =''

        formB_columns = ['Employee Code','Employee Name','FIXED MONTHLY GROSS',	'Days Paid','Total\r\nOT Hrs',	'basic_and_allo', 'Overtime',	'HRA',	'Tel and Int Reimb', 'Bonus', 'Fuel Reimb',	'Prof Dev Reimb', 'Corp Attire Reimb',	'CCA',	'Leave Encashment','Conveyance', 'Medical Allowance', 'Telephone Reimb', 'Other Allowance', 'Meal Allowance',
       'Special Allowance', 'Personal Allowance','Other Reimb', 'Arrears', 'Other Earning', 'Variable Pay','Stipend' ,'Total Earning', 'PF',	'ESIC',	'VPF', 'Loan Deduction', 'Loan Interest', 'P.Tax',	'Society',	'Income Tax', 'Insurance',	'LWF EE',	'Other Deduction',	'TDS',	'Total Deductions',	'Net Paid',	'EMP PF','Bank A/c Number','Date of payment','Remarks']

        formB_data = data_formB[formB_columns]

        formBsheet = formBfile['FORM B']

        formBsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form B is ready')

        
        rows = dataframe_to_rows(formB_data, index=False, header=False)

        logging.info('rows taken out from data')

        

        for r_idx, row in enumerate(rows, 21):
            for c_idx, value in enumerate(row, 2):
                formBsheet.cell(row=r_idx, column=c_idx, value=value)
                formBsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                formBsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formBsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)
                if c_idx==45:
                    formBsheet.cell(row=r_idx, column=c_idx).number_format= numbers.FORMAT_NUMBER

        contractline = formBsheet['B10'].value
        B10_data = contractline+' '+contractor_name+', '+contractor_address
        formBsheet['B10'] = B10_data

        if str(data_formB['Nature of work'].dtype)[0:3] != 'obj':
            data_formB['Nature of work'] = data_formB['Nature of work'].astype(str)

        if str(data_formB['Location'].dtype)[0:3] != 'obj':
            data_formB['Location'] = data_formB['Location'].astype(str)

        if str(data_formB['Company Name'].dtype)[0:3] != 'obj':
            data_formB['Company Name'] = data_formB['Company Name'].astype(str)

        if str(data_formB['Company Address'].dtype)[0:3] != 'obj':
            data_formB['Company Address'] = data_formB['Company Address'].astype(str)

        if str(data_formB['Unit'].dtype)[0:3] != 'obj':
            data_formB['Unit'] = data_formB['Unit'].astype(str)

        if str(data_formB['Address'].dtype)[0:3] != 'obj':
            data_formB['Address'] = data_formB['Address'].astype(str)

        locationline = formBsheet['B11'].value
        B11_data = locationline+' '+data_formB['Nature of work'].unique()[0]+', '+data_formB['Location'].unique()[0]
        formBsheet['B11'] = B11_data

        establine = formBsheet['B12'].value
        if data_formB['PE_or_contract'].unique()[0]== 'PE':
            B12_data = establine+' '+data_formB['Company Name'].unique()[0]+', '+data_formB['Company Address'].unique()[0]  
        else:    
            B12_data = establine+' '+data_formB['Unit'].unique()[0]+', '+data_formB['Address'].unique()[0]
        formBsheet['B12'] = B12_data

        peline = formBsheet['B13'].value
        if data_formB['PE_or_contract'].unique()[0]== 'PE':
            B13_data = peline+' '+data_formB['Company Name'].unique()[0]+', '+data_formB['Company Address'].unique()[0]  
        else:    
            B13_data = peline+' '+data_formB['Unit'].unique()[0]+', '+data_formB['Address'].unique()[0]
        formBsheet['B13'] = B13_data

        monthstart = datetime.date(year,month_num,1)
        monthend = datetime.date(year,month_num,calendar.monthrange(year,month_num)[1])
        formBsheet['B16'] = 'Wage period From: '+str(monthstart)+' to '+str(monthend)

        formBfinalfile = os.path.join(filelocation,'FormB.xlsx')
        formBfile.save(filename=formBfinalfile)

    def create_form_XXI():
        formXXIfilepath = os.path.join(karnatakafilespath,'FormXXI.xlsx')
        formXXIfile = load_workbook(filename=formXXIfilepath)
        logging.info('Form XXI file has sheet: '+str(formXXIfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_formXXI = data.drop_duplicates(subset=['Employee Code']).copy()

        data_formXXI.fillna(value=0, inplace=True)

        data_formXXI['S.no'] = list(range(1,len(data_formXXI)+1))

        data_formXXI['a'] ='---'
        data_formXXI['b'] ='---'
        data_formXXI['c'] ='---'
        #data_formXXI['e'] ='---'
        data_formXXI['f'] ='---'
        data_formXXI['g'] =''

        formXXI_columns = ['S.no','Employee Name',"Father's Name",'Designation','a','b','c','FIXED MONTHLY GROSS','Fine','f','g']

        formXXI_data = data_formXXI[formXXI_columns]

        formXXIsheet = formXXIfile['FORM XXI']

        formXXIsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form XXI is ready')

        
        rows = dataframe_to_rows(formXXI_data, index=False, header=False)

        logging.info('rows taken out from data')


        for r_idx, row in enumerate(rows, 14):
            for c_idx, value in enumerate(row, 3):
                formXXIsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXIsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                formXXIsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formXXIsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)

        contractline = formXXIsheet['C7'].value
        C7_data = contractline+' '+contractor_name+', '+contractor_address
        formXXIsheet['C7'] = C7_data

        if str(data_formXXI['Nature of work'].dtype)[0:3] != 'obj':
            data_formXXI['Nature of work'] = data_formXXI['Nature of work'].astype(str)

        if str(data_formXXI['Location'].dtype)[0:3] != 'obj':
            data_formXXI['Location'] = data_formXXI['Location'].astype(str)

        if str(data_formXXI['Company Name'].dtype)[0:3] != 'obj':
            data_formXXI['Company Name'] = data_formXXI['Company Name'].astype(str)

        if str(data_formXXI['Company Address'].dtype)[0:3] != 'obj':
            data_formXXI['Company Address'] = data_formXXI['Company Address'].astype(str)

        if str(data_formXXI['Unit'].dtype)[0:3] != 'obj':
            data_formXXI['Unit'] = data_formXXI['Unit'].astype(str)

        if str(data_formXXI['Address'].dtype)[0:3] != 'obj':
            data_formXXI['Address'] = data_formXXI['Address'].astype(str)

        locationline = formXXIsheet['C8'].value
        C8_data = locationline+' '+data_formXXI['Nature of work'].unique()[0]+', '+data_formXXI['Location'].unique()[0]
        formXXIsheet['C8'] = C8_data

        establine = formXXIsheet['C9'].value
        if data_formXXI['PE_or_contract'].unique()[0]== 'PE':
            C9_data = establine+' '+data_formXXI['Company Name'].unique()[0]+', '+data_formXXI['Company Address'].unique()[0]
        else:
            C9_data = establine+' '+data_formXXI['Unit'].unique()[0]+', '+data_formXXI['Address'].unique()[0]
        formXXIsheet['C9'] = C9_data

        peline = formXXIsheet['C10'].value
        if data_formXXI['PE_or_contract'].unique()[0]== 'PE':
            C10_data = peline+' '+data_formXXI['Company Name'].unique()[0]+', '+data_formXXI['Company Address'].unique()[0]
        else:
            C10_data = peline+' '+data_formXXI['Unit'].unique()[0]+', '+data_formXXI['Address'].unique()[0]
        formXXIsheet['C10'] = C10_data

        #border the region
        count1 = len(data_formXXI)
        border_1 = Side(style='thick')
        for i in range(2,15):
            formXXIsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
            formXXIsheet.cell(row=count1+15, column=i).border = Border(outline= True, bottom=border_1)
        for i in range(2,count1+16):
            formXXIsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
            formXXIsheet.cell(row=i, column=14).border = Border(outline= True, right=border_1)

        formXXIfinalfile = os.path.join(filelocation,'FormXXI.xlsx')
        formXXIfile.save(filename=formXXIfinalfile)



    def create_form_XXII():
        formXXIIfilepath = os.path.join(karnatakafilespath,'FormXXII.xlsx')
        formXXIIfile = load_workbook(filename=formXXIIfilepath)
        logging.info('Form XXII file has sheet: '+str(formXXIIfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_formXXII = data.drop_duplicates(subset=['Employee Code']).copy()

        data_formXXII.fillna(value=0, inplace=True)

        data_formXXII['S.no'] = list(range(1,len(data_formXXII)+1))

        data_formXXII['b'] ='---'
        data_formXXII['c'] ='---'
        data_formXXII['d'] ='---'
        data_formXXII['e'] ='---'
        data_formXXII['f'] ='---'
        data_formXXII['g'] =''

        formXXII_columns = ['S.no','Employee Name',"Father's Name",'Designation','FIXED MONTHLY GROSS','b','c','d','e','f','g']

        formXXII_data = data_formXXII[formXXII_columns]

        formXXIIsheet = formXXIIfile['FORM XXII']

        formXXIIsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form XXII is ready')

        
        rows = dataframe_to_rows(formXXII_data, index=False, header=False)

        logging.info('rows taken out from data')


        for r_idx, row in enumerate(rows, 15):
            for c_idx, value in enumerate(row, 3):
                formXXIIsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXIIsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                formXXIIsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formXXIIsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)

        contractline = formXXIIsheet['C7'].value
        C7_data = contractline+' '+contractor_name+', '+contractor_address
        formXXIIsheet['C7'] = C7_data


        if str(data_formXXII['Nature of work'].dtype)[0:3] != 'obj':
            data_formXXII['Nature of work'] = data_formXXII['Nature of work'].astype(str)

        if str(data_formXXII['Location'].dtype)[0:3] != 'obj':
            data_formXXII['Location'] = data_formXXII['Location'].astype(str)

        if str(data_formXXII['Company Name'].dtype)[0:3] != 'obj':
            data_formXXII['Company Name'] = data_formXXII['Company Name'].astype(str)

        if str(data_formXXII['Company Address'].dtype)[0:3] != 'obj':
            data_formXXII['Company Address'] = data_formXXII['Company Address'].astype(str)

        if str(data_formXXII['Unit'].dtype)[0:3] != 'obj':
            data_formXXII['Unit'] = data_formXXII['Unit'].astype(str)

        if str(data_formXXII['Address'].dtype)[0:3] != 'obj':
            data_formXXII['Address'] = data_formXXII['Address'].astype(str)

        locationline = formXXIIsheet['C8'].value
        C8_data = locationline+' '+data_formXXII['Nature of work'].unique()[0]+', '+data_formXXII['Location'].unique()[0]
        formXXIIsheet['C8'] = C8_data

        establine = formXXIIsheet['C9'].value
        if data_formXXII['PE_or_contract'].unique()[0]== 'PE':
            C9_data = establine+' '+data_formXXII['Company Name'].unique()[0]+', '+data_formXXII['Company Address'].unique()[0]
        else:
            C9_data = establine+' '+data_formXXII['Unit'].unique()[0]+', '+data_formXXII['Address'].unique()[0]
        formXXIIsheet['C9'] = C9_data

        peline = formXXIIsheet['C10'].value
        if data_formXXII['PE_or_contract'].unique()[0]== 'PE':
            C10_data = peline+' '+data_formXXII['Company Name'].unique()[0]+', '+data_formXXII['Company Address'].unique()[0]
        else:
            C10_data = peline+' '+data_formXXII['Unit'].unique()[0]+', '+data_formXXII['Address'].unique()[0]
        formXXIIsheet['C10'] = C10_data

        #border the region
        count1 = len(data_formXXII)
        border_1 = Side(style='thick')
        for i in range(2,15):
            formXXIIsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
            formXXIIsheet.cell(row=count1+16, column=i).border = Border(outline= True, bottom=border_1)
        for i in range(2,count1+17):
            formXXIIsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
            formXXIIsheet.cell(row=i, column=14).border = Border(outline= True, right=border_1)

        formXXIIfinalfile = os.path.join(filelocation,'FormXXII.xlsx')
        formXXIIfile.save(filename=formXXIIfinalfile)


    def create_form_XXIII():
        formXXIIIfilepath = os.path.join(karnatakafilespath,'FormXXIII.xlsx')
        formXXIIIfile = load_workbook(filename=formXXIIIfilepath)
        logging.info('Form XXIII file has sheet: '+str(formXXIIIfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_formXXIII = data.drop_duplicates(subset=['Employee Code']).copy()
        data_formXXIII.fillna(value=0, inplace=True)

        data_formXXIII['S.no'] = list(range(1,len(data_formXXIII)+1))

        data_formXXIII['b'] ='---'
        data_formXXIII['c'] ='---'
        data_formXXIII['d'] ='---'
        data_formXXIII['e'] ='---'
        data_formXXIII['f'] ='---'
        data_formXXIII['g'] =''

        formXXIII_columns = ['S.no','Employee Name',"Father's Name",'Designation','FIXED MONTHLY GROSS','b','c','d','e','f','g']

        formXXIII_data = data_formXXIII[formXXIII_columns]

        formXXIIIsheet = formXXIIIfile['FORM XXIII']

        formXXIIIsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form XXIII is ready')

        
        rows = dataframe_to_rows(formXXIII_data, index=False, header=False)

        logging.info('rows taken out from data')


        for r_idx, row in enumerate(rows, 12):
            for c_idx, value in enumerate(row, 3):
                formXXIIIsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXIIIsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                formXXIIIsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formXXIIIsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)

        contractline = formXXIIIsheet['C5'].value
        C5_data = contractline+' '+contractor_name+', '+contractor_address
        formXXIIIsheet['C5'] = C5_data

        if str(data_formXXIII['Nature of work'].dtype)[0:3] != 'obj':
            data_formXXIII['Nature of work'] = data_formXXIII['Nature of work'].astype(str)

        if str(data_formXXIII['Location'].dtype)[0:3] != 'obj':
            data_formXXIII['Location'] = data_formXXIII['Location'].astype(str)

        if str(data_formXXIII['Company Name'].dtype)[0:3] != 'obj':
            data_formXXIII['Company Name'] = data_formXXIII['Company Name'].astype(str)

        if str(data_formXXIII['Company Address'].dtype)[0:3] != 'obj':
            data_formXXIII['Company Address'] = data_formXXIII['Company Address'].astype(str)

        if str(data_formXXIII['Unit'].dtype)[0:3] != 'obj':
            data_formXXIII['Unit'] = data_formXXIII['Unit'].astype(str)

        if str(data_formXXIII['Address'].dtype)[0:3] != 'obj':
            data_formXXIII['Address'] = data_formXXIII['Address'].astype(str)

        locationline = formXXIIIsheet['C6'].value
        C6_data = locationline+' '+data_formXXIII['Nature of work'].unique()[0]+', '+data_formXXIII['Location'].unique()[0]
        formXXIIIsheet['C6'] = C6_data

        establine = formXXIIIsheet['C7'].value
        if data_formXXIII['PE_or_contract'].unique()[0]== 'PE':
            C7_data = establine+' '+data_formXXIII['Company Name'].unique()[0]+', '+data_formXXIII['Company Address'].unique()[0]
        else:
            C7_data = establine+' '+data_formXXIII['Unit'].unique()[0]+', '+data_formXXIII['Address'].unique()[0]
        formXXIIIsheet['C7'] = C7_data

        peline = formXXIIIsheet['C8'].value
        if data_formXXIII['PE_or_contract'].unique()[0]== 'PE':
            C8_data = peline+' '+data_formXXIII['Company Name'].unique()[0]+', '+data_formXXIII['Company Address'].unique()[0]
        else:
            C8_data = peline+' '+data_formXXIII['Unit'].unique()[0]+', '+data_formXXIII['Address'].unique()[0]
        formXXIIIsheet['C8'] = C8_data


        #border the region
        count1 = len(data_formXXIII)
        border_1 = Side(style='thick')
        for i in range(2,15):
            formXXIIIsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
            formXXIIIsheet.cell(row=count1+13, column=i).border = Border(outline= True, bottom=border_1)
        for i in range(2,count1+14):
            formXXIIIsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
            formXXIIIsheet.cell(row=i, column=14).border = Border(outline= True, right=border_1)


        formXXIIIfinalfile = os.path.join(filelocation,'FormXXIII.xlsx')
        formXXIIIfile.save(filename=formXXIIIfinalfile)


    def create_form_XX():
        formXXfilepath = os.path.join(karnatakafilespath,'FormXX.xlsx')
        formXXfile = load_workbook(filename=formXXfilepath)
        logging.info('Form XX file has sheet: '+str(formXXfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_formXX = data.drop_duplicates(subset=['Employee Code']).copy()
        data_formXX.fillna(value=0, inplace=True)

        data_formXX['S.no'] = list(range(1,len(data_formXX)+1))

        data_formXX['a'] ='---'
        data_formXX['b'] ='---'
        data_formXX['c'] ='---'
        data_formXX['d'] ='---'
        data_formXX['e'] ='---'
        data_formXX['f'] ='---'
        data_formXX['g'] ='---'
        data_formXX['h'] ='---'
        data_formXX['i'] =''

        formXX_columns = ['S.no','Employee Name',"Father's Name",'Designation','a','b','c','d','e','f','g','h','i']

        formXX_data = data_formXX[formXX_columns]

        formXXsheet = formXXfile['FORM XX']

        formXXsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form XX is ready')

        
        rows = dataframe_to_rows(formXX_data, index=False, header=False)

        logging.info('rows taken out from data')


        for r_idx, row in enumerate(rows, 15):
            for c_idx, value in enumerate(row, 3):
                formXXsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                formXXsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formXXsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)

        contractline = formXXsheet['C6'].value
        C6_data = contractline+' '+contractor_name+', '+contractor_address
        formXXsheet['C6'] = C6_data

        if str(data_formXX['Nature of work'].dtype)[0:3] != 'obj':
            data_formXX['Nature of work'] = data_formXX['Nature of work'].astype(str)

        if str(data_formXX['Location'].dtype)[0:3] != 'obj':
            data_formXX['Location'] = data_formXX['Location'].astype(str)

        if str(data_formXX['Company Name'].dtype)[0:3] != 'obj':
            data_formXX['Company Name'] = data_formXX['Company Name'].astype(str)

        if str(data_formXX['Company Address'].dtype)[0:3] != 'obj':
            data_formXX['Company Address'] = data_formXX['Company Address'].astype(str)

        if str(data_formXX['Unit'].dtype)[0:3] != 'obj':
            data_formXX['Unit'] = data_formXX['Unit'].astype(str)

        if str(data_formXX['Address'].dtype)[0:3] != 'obj':
            data_formXX['Address'] = data_formXX['Address'].astype(str)

        locationline = formXXsheet['C7'].value
        C7_data = locationline+' '+data_formXX['Nature of work'].unique()[0]+', '+data_formXX['Location'].unique()[0]
        formXXsheet['C7'] = C7_data

        establine = formXXsheet['C8'].value
        if data_formXX['PE_or_contract'].unique()[0]== 'PE':
            C8_data = establine+' '+data_formXX['Company Name'].unique()[0]+', '+data_formXX['Company Address'].unique()[0]
        else:
            C8_data = establine+' '+data_formXX['Unit'].unique()[0]+', '+data_formXX['Address'].unique()[0]
        formXXsheet['C8'] = C8_data

        peline = formXXsheet['C9'].value
        if data_formXX['PE_or_contract'].unique()[0]== 'PE':
            C9_data = peline+' '+data_formXX['Company Name'].unique()[0]+', '+data_formXX['Company Address'].unique()[0]
        else:
            C9_data = peline+' '+data_formXX['Unit'].unique()[0]+', '+data_formXX['Address'].unique()[0]
        formXXsheet['C9'] = C9_data

        #border the region
        count1 = len(data_formXX)
        border_1 = Side(style='thick')
        for i in range(2,17):
            formXXsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
            formXXsheet.cell(row=count1+16, column=i).border = Border(outline= True, bottom=border_1)
        for i in range(2,count1+17):
            formXXsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
            formXXsheet.cell(row=i, column=16).border = Border(outline= True, right=border_1)

        formXXfinalfile = os.path.join(filelocation,'FormXX.xlsx')
        formXXfile.save(filename=formXXfinalfile)

    def create_wages():
        wagesfilepath = os.path.join(karnatakafilespath,'Wages.xlsx')
        wagesfile = load_workbook(filename=wagesfilepath)
        logging.info('wages file has sheet: '+str(wagesfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_wages = data.drop_duplicates(subset=['Employee Code']).copy()
        data_wages.fillna(value=0, inplace=True)

        data_wages['S.no'] = list(range(1,len(data_wages)+1))


        data_wages['fixed_wage'] = '---'
        #data_wages['OT hours'] = 0

        if str(data_wages['Earned Basic'].dtype)[0:3] != 'int':
            data_wages['Earned Basic']= data_wages['Earned Basic'].astype(int)
        if str(data_wages['DA'].dtype)[0:3] != 'int':
            data_wages['DA']= data_wages['DA'].astype(int)
        data_wages['basic_and_allo'] = data_wages['Earned Basic']+ data_wages['DA']
        data_wages['NFH'] = '---'
        data_wages['maturity'] = '---'
        data_wages['Sub Allow'] = '---'
        data_wages['Society'] = '---'
        data_wages['Fines']= '---'
        data_wages['Damages']= '---'
        data_wages['Pay mode'] = 'Bank Transfer'
        data_wages['Remarks'] =''

        wages_columns = ['S.no','Employee Code','Employee Name',"Father's Name",'Gender',
                                'Designation','Department','Address','Date Joined','ESIC Number',
                                'PF Number','fixed_wage','Days Paid','Total\r\nOT Hrs',
                                'basic_and_allo','HRA','Conveyance','Medical Allowance',
                                'Tel and Int Reimb', 'Bonus', 'Fuel Reimb',	'Prof Dev Reimb', 
                                'Corp Attire Reimb','Special Allowance','Overtime','NFH',
                                'maturity','Other Reimb', 'CCA', 'Medical Allowance', 
                                'Telephone Reimb', 'Other Allowance', 'Meal Allowance',
                                'Special Allowance', 'Personal Allowance', 'Arrears', 
                                'Other Earning', 'Variable Pay','Stipend','Sub Allow',
                                'Leave Encashment', 'Total Earning','ESIC', 'PF','P.Tax',
                                'TDS','Society','Insurance','Salary Advance','Fines','Damages',
                                'Other Deduction',	'Total Deductions',	'Net Paid','Pay mode',
                                'Bank A/c Number','Remarks']

        wages_data = data_wages[wages_columns]

        wagessheet = wagesfile['Wages']

        wagessheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for wages is ready')

        
        rows = dataframe_to_rows(wages_data, index=False, header=False)

        logging.info('rows taken out from data')

        

        for r_idx, row in enumerate(rows, 18):
            for c_idx, value in enumerate(row, 1):
                wagessheet.cell(row=r_idx, column=c_idx, value=value)
                wagessheet.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                wagessheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                wagessheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)
                if c_idx==56:
                    wagessheet.cell(row=r_idx, column=c_idx).number_format= numbers.FORMAT_NUMBER

        contractline = wagessheet['A10'].value
        A10_data = contractline+' '+contractor_name+', '+contractor_address
        wagessheet['A10'] = A10_data

        if str(data_wages['Nature of work'].dtype)[0:3] != 'obj':
            data_wages['Nature of work'] = data_wages['Nature of work'].astype(str)

        if str(data_wages['Location'].dtype)[0:3] != 'obj':
            data_wages['Location'] = data_wages['Location'].astype(str)

        if str(data_wages['Company Name'].dtype)[0:3] != 'obj':
            data_wages['Company Name'] = data_wages['Company Name'].astype(str)

        if str(data_wages['Company Address'].dtype)[0:3] != 'obj':
            data_wages['Company Address'] = data_wages['Company Address'].astype(str)

        if str(data_wages['Unit'].dtype)[0:3] != 'obj':
            data_wages['Unit'] = data_wages['Unit'].astype(str)

        if str(data_wages['Address'].dtype)[0:3] != 'obj':
            data_wages['Address'] = data_wages['Address'].astype(str)

        locationline = wagessheet['A11'].value
        A11_data = locationline+' '+data_wages['Nature of work'].unique()[0]+', '+data_wages['Location'].unique()[0]
        wagessheet['A11'] = A11_data

        establine = wagessheet['A12'].value
        if data_wages['PE_or_contract'].unique()[0]== 'PE':
            A12_data = establine+' '+data_wages['Company Name'].unique()[0]+', '+data_wages['Company Address'].unique()[0]
        else:
            A12_data = establine+' '+data_wages['Unit'].unique()[0]+', '+data_wages['Address'].unique()[0]
        wagessheet['A12'] = A12_data

        peline = wagessheet['A13'].value
        if data_wages['PE_or_contract'].unique()[0]== 'PE':
            A13_data = peline+' '+data_wages['Company Name'].unique()[0]+', '+data_wages['Company Address'].unique()[0]
        else:
            A13_data = peline+' '+data_wages['Unit'].unique()[0]+', '+data_wages['Address'].unique()[0]
        wagessheet['A13'] = A13_data

        wagessheet['F4'] = 'Combined Muster Roll-cum-Register of Wages in lieu of '+month+' '+str(year)

        wagesfinalfile = os.path.join(filelocation,'Wages.xlsx')
        wagesfile.save(filename=wagesfinalfile)

    
    def create_form_H_F(form):
        if form=='FORM H':
            formHfilepath = os.path.join(karnatakafilespath,'FormH.xlsx')
        if form=='FORM F':
            formHfilepath = os.path.join(karnatakafilespath,'FormF.xlsx')
        formHfile = load_workbook(filename=formHfilepath)
        logging.info('file has sheet: '+str(formHfile.sheetnames))
        sheetformh = formHfile[form]

        
        logging.info('create columns which are now available')

        data_formH = data[data['Leave Type']=='PL'].copy()
        data_formH.fillna(value=0, inplace=True)

        def attandance_data(employee_attendance,i):

            leavelist = list(employee_attendance.columns[(employee_attendance=='PL').iloc[i]])
            empcodeis = employee_attendance.iloc[i]['Employee Code']
            logging.info(empcodeis)
            if 'Leave Type' in leavelist:
                leavelist.remove('Leave Type')
            emp1 = pd.DataFrame(leavelist)
            
            
            if len(emp1.index)==0:
                defaultemp = {'emp':(employee_attendance).iloc[i]['Employee Code'],'startdate':0,'enddate':0,'days':0,'start_date':'-------','end_date':'-------'}
                emp1 = pd.DataFrame(defaultemp, index=[0])
                emp1.index = np.arange(1, len(emp1) + 1)
                emp1['s.no'] = emp1.index
                emp1.reset_index(drop=True, inplace=True)
                emp1['from'] = datetime.date(year,month_num,1)
                emp1['to'] = datetime.date(year,month_num,calendar.monthrange(year,month_num)[1])
                emp1['totaldays'] = (employee_attendance).iloc[i]['Days Paid']
                emp1['leavesearned'] = float(employee_attendance.iloc[i]['Monthly Increment'])
                emp1['leavesstart'] = float(employee_attendance.iloc[i]['Opening'])
                emp1['leavesend'] = float(employee_attendance.iloc[i]['Closing'])
                emp1['Date of Payment and fixed'] = str(employee_attendance.iloc[i]['Date of payment'])+' and '+str(employee_attendance.iloc[i]['FIXED MONTHLY GROSS'])
                emp1['a']='---'
                emp1['b']='---'
                emp1['c']='---'
                emp1['d']='---'
                emp1['e']='---'
                emp1['f']='---'
                emp1 = emp1[["s.no","from","to","totaldays","leavesearned","leavesstart","start_date","end_date","days","leavesend","Date of Payment and fixed",'a','b','c','d','e','f']]
            else:
                logging.info(emp1)
                emp1.columns = ['Leaves']
                emp1['emp'] = (employee_attendance).iloc[i]['Employee Code']
                emp1['Leavesdays'] = emp1.Leaves.str[5:7].astype(int)
                emp1['daysdiff'] = (emp1.Leavesdays.shift(-1) - emp1.Leavesdays).fillna(0).astype(int)
                emp1['startdate'] = np.where(emp1.daysdiff.shift() != 1, emp1.Leavesdays, 0)
                emp1['enddate'] = np.where(emp1.daysdiff!=1, emp1.Leavesdays, 0)
                emp1.drop(emp1[(emp1.startdate==0) & (emp1.enddate==0)].index, inplace=True)
                emp1['startdate'] = np.where(emp1.startdate ==0, emp1.startdate.shift(), emp1.startdate).astype(int)
                emp1['enddate'] = np.where(emp1.enddate ==0, emp1.enddate.shift(-1), emp1.enddate).astype(int)
                emp1 = emp1[['emp','startdate','enddate']]
                emp1.drop_duplicates(subset='startdate', inplace=True)
                emp1['days'] = emp1.enddate -emp1.startdate +1
                emp1['start_date'] = [datetime.date(year,month_num,x) for x in emp1.startdate]
                emp1['end_date'] = [datetime.date(year,month_num,x) for x in emp1.enddate]
                emp1.index = np.arange(1, len(emp1) + 1)
                emp1['s.no'] = emp1.index
                emp1.reset_index(drop=True, inplace=True)
                emp1['from'] = datetime.date(year,month_num,1)
                emp1['to'] = datetime.date(year,month_num,calendar.monthrange(year,month_num)[1])
                emp1['totaldays'] = (employee_attendance).iloc[i]['Days Paid']
                emp1['leavesearned'] = float((employee_attendance).iloc[i]['Monthly Increment'])
                emp1['totalleaves']= float(employee_attendance.iloc[i]['Opening'])
                emp1['leavesend'] = float(employee_attendance.iloc[i]['Closing'])
                emp1['leavesstart'] =emp1['totalleaves']
                emp1['Date of Payment and fixed'] = str(employee_attendance.iloc[i]['Date of payment'])+' and '+str(employee_attendance.iloc[i]['FIXED MONTHLY GROSS'])
                emp1['a']='---'
                emp1['b']='---'
                emp1['c']='---'
                emp1['d']='---'
                emp1['e']='---'
                emp1['f']='---'
                emp1 = emp1[["s.no","from","to","totaldays","leavesearned","leavesstart","start_date","end_date","days","leavesend","Date of Payment and fixed",'a','b','c','d','e','f']]
                
            
            return emp1

        def prepare_emp_sheet(emp1,sheet_key,key,name,fathername):
            
            sheet1 = formHfile.copy_worksheet(sheetformh)
            sheet1.title = sheet_key
            lastline = sheet1['B18'].value
            sheet1['B18'] =''

            if len(emp1)>3:
                lastlinerow = 'B'+str(18+len(emp1))
            else:
                lastlinerow = 'B18'

            
            logging.info(lastlinerow)
            sheet1[lastlinerow] = lastline

            
            
            rows = dataframe_to_rows(emp1, index=False, header=False)

            for r_idx, row in enumerate(rows, 14):
                for c_idx, value in enumerate(row, 2):
                    sheet1.cell(row=r_idx, column=c_idx, value=value)
                    sheet1.cell(row=r_idx, column=c_idx).font =Font(name ='Verdana', size =8)
                    sheet1.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                    border_sides = Side(style='thin')
                    sheet1.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)
            sheet1['H5']=key
            sheet1['F7']=name
            sheet1['F8']=fathername

            sheet1.sheet_properties.pageSetUpPr.fitToPage = True


        emp_count = len(data_formH.index)
        emp_dic = dict()
        for i in range(0,emp_count):
            key = (data_formH).iloc[i]['Employee Code']
            emp_dic[key] = attandance_data(data_formH,i)
            sheet_key = form+'_'+str(key)
            name= data_formH[data_formH['Employee Code']==key]['Employee Name'].values[0]
            fathername= data_formH[data_formH['Employee Code']==key]["Father's Name"].values[0]
            logging.info(name)
            logging.info(fathername)
            prepare_emp_sheet(emp_dic[key],sheet_key,key,name,fathername)
            logging.info(key)
            logging.info(sheet_key)
        if form=='FORM H':
            formHfinalfile = os.path.join(filelocation,'FormH.xlsx')
        if form=='FORM F':
            formHfinalfile = os.path.join(filelocation,'FormF.xlsx')
        
        formHfile.remove(sheetformh)
        formHfile.save(filename=formHfinalfile)

    
    def create_muster():

        musterfilepath = os.path.join(karnatakafilespath,'Muster.xlsx')
        musterfile = load_workbook(filename=musterfilepath)
        logging.info('muster file has sheet: '+str(musterfile.sheetnames))

        
        logging.info('create columns which are now available')

        data_muster = data.drop_duplicates(subset=['Employee Code']).copy()
        data_muster.fillna(value=0, inplace=True)

        data_muster['S.no'] = list(range(1,len(data_muster)+1))

        data_muster['datelast'] ='---'

        first3columns = ["S.no",'Employee Code','Employee Name']
        last2columns = ["datelast","Days Paid"]

        columnstotake =[]
        days = ['01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
        for day in days:
            for col in data_muster.columns:
                if col[5:7]==day:
                    columnstotake.append(col)
        if len(columnstotake)==28:

            columnstotake.append('29')
            columnstotake.append('30')
            columnstotake.append('31')
            data_muster['29'] = ''
            data_muster['30'] = ''
            data_muster['31'] = ''
            
        elif len(columnstotake)==29:
            columnstotake.append('30')
            columnstotake.append('31')
            data_muster['30'] = ''
            data_muster['31'] = ''

        elif len(columnstotake)==30:
            columnstotake.append('31')
            data_muster['31'] = ''

        muster_columns = first3columns+columnstotake+last2columns

        muster_data = data_muster[muster_columns]

        mustersheet = musterfile['Muster']

        mustersheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for muster is ready')

        
        rows = dataframe_to_rows(muster_data, index=False, header=False)

        logging.info('rows taken out from data')

        

        for r_idx, row in enumerate(rows, 18):
            for c_idx, value in enumerate(row, 2):
                mustersheet.cell(row=r_idx, column=c_idx, value=value)
                mustersheet.cell(row=r_idx, column=c_idx).font =Font(name ='Bell MT', size =10)
                mustersheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                mustersheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)

        logging.info('')

        contractline = mustersheet['B10'].value
        B10_data = contractline+' '+contractor_name+', '+contractor_address
        mustersheet['B10'] = B10_data

        if str(data_muster['Nature of work'].dtype)[0:3] != 'obj':
            data_muster['Nature of work'] = data_muster['Nature of work'].astype(str)

        if str(data_muster['Location'].dtype)[0:3] != 'obj':
            data_muster['Location'] = data_muster['Location'].astype(str)

        if str(data_muster['Company Name'].dtype)[0:3] != 'obj':
            data_muster['Company Name'] = data_muster['Company Name'].astype(str)

        if str(data_muster['Company Address'].dtype)[0:3] != 'obj':
            data_muster['Company Address'] = data_muster['Company Address'].astype(str)

        if str(data_muster['Unit'].dtype)[0:3] != 'obj':
            data_muster['Unit'] = data_muster['Unit'].astype(str)

        if str(data_muster['Address'].dtype)[0:3] != 'obj':
            data_muster['Address'] = data_muster['Address'].astype(str)

        locationline = mustersheet['B11'].value
        B11_data = locationline+' '+data_muster['Nature of work'].unique()[0]+', '+data_muster['Location'].unique()[0]
        mustersheet['B11'] = B11_data

        establine = mustersheet['B12'].value
        if data_muster['PE_or_contract'].unique()[0]== 'PE':
            B12_data = establine+' '+data_muster['Company Name'].unique()[0]+', '+data_muster['Company Address'].unique()[0]
        else:
            B12_data = establine+' '+data_muster['Unit'].unique()[0]+', '+data_muster['Address'].unique()[0]
        mustersheet['B12'] = B12_data

        peline = mustersheet['B13'].value
        if data_muster['PE_or_contract'].unique()[0]== 'PE':
            B13_data = peline+' '+data_muster['Company Name'].unique()[0]+', '+data_muster['Company Address'].unique()[0]
        else:
            B13_data = peline+' '+data_muster['Unit'].unique()[0]+', '+data_muster['Address'].unique()[0]
        mustersheet['B13'] = B13_data

        mustersheet['B4'] = 'Combined Muster Roll-cum-Register of Wages in lieu of '+month+' '+str(year)

        musterfinalfile = os.path.join(filelocation,'Muster.xlsx')
        musterfile.save(filename=musterfinalfile)

    def create_formXIX():

        formXIXfilepath = os.path.join(karnatakafilespath,'FormXIX.xlsx')
        formXIXfile = load_workbook(filename=formXIXfilepath)
        logging.info('Form XIX file has sheet: '+str(formXIXfile.sheetnames))
        sheetformXIX = formXIXfile['FORM XIX']

        
        logging.info('create columns which are now available')

        data_formXIX = data.drop_duplicates(subset=['Employee Code']).copy()

        data_formXIX.fillna(value=0, inplace=True)

        emp_count = len(data_formXIX.index)
        
        for i in range(0,emp_count):
            key = (data_formXIX).iloc[i]['Employee Code']
            sheet_key = 'FORM XIX_'+str(key)

            emp_data = (data_formXIX).iloc[i]

            sheet1 = formXIXfile.copy_worksheet(sheetformXIX)
            sheet1.title = sheet_key
            sheet1['D7'] = contractor_name+', '+contractor_address
            sheet1['D8'] = str(emp_data['Nature of work'])+', '+str(emp_data['Location'])
            if emp_data['PE_or_contract'][0]== 'PE':
                sheet1['D9'] = emp_data['Company Name']+', '+emp_data['Company Address']
                sheet1['D10'] = emp_data['Company Name']+', '+emp_data['Company Address']
            else:
                sheet1['D9'] = emp_data['Unit']+', '+emp_data['Address']
                sheet1['D10'] = emp_data['Unit']+', '+emp_data['Address']
            sheet1['D11'] = emp_data['Employee Name']
            sheet1['D12'] = emp_data['Gender']
            sheet1['D13'] = month+'-'+str(year)
            sheet1['D14'] = key
            sheet1['D15'] = emp_data['Days Paid']
            sheet1['D16'] = emp_data['Earned Basic']
            sheet1['D17'] = emp_data['DA']
            sheet1['D18'] = emp_data['HRA']
            sheet1['D19'] = emp_data['Tel and Int Reimb']
            sheet1['D20'] = emp_data['Bonus']
            sheet1['D21'] = emp_data['Fuel Reimb']
            sheet1['D22'] = emp_data['Corp Attire Reimb']
            sheet1['D23'] = emp_data['CCA']
            sheet1['D24'] = emp_data['Conveyance']+emp_data['Medical Allowance']+emp_data['Telephone Reimb']+emp_data['Other Allowance']+emp_data['Prof Dev Reimb']+emp_data['Meal Allowance']+emp_data['Special Allowance']+emp_data['Personal Allowance']+emp_data['Other Reimb']+emp_data['Arrears']+emp_data['Variable Pay']+emp_data['Other Earning']+emp_data['Leave Encashment']+emp_data['Stipend']
            sheet1['D25'] = emp_data['Total Earning']
            sheet1['D26'] = emp_data['Insurance']
            sheet1['D27'] = emp_data['PF']
            sheet1['D28'] = emp_data['ESIC']
            sheet1['D29'] = emp_data['P.Tax']
            sheet1['D30'] = emp_data['TDS']
            sheet1['D31'] = emp_data['CSR']+emp_data['VPF']+emp_data['LWF EE']+emp_data['Salary Advance']+emp_data['Loan Deduction']+emp_data['Loan Interest']+emp_data['Other Deduction']
            sheet1['D32'] = emp_data['Total Deductions']
            sheet1['D33'] = emp_data['Net Paid']

        formXIXfinalfile = os.path.join(filelocation,'FormXIX.xlsx')
        formXIXfile.remove(sheetformXIX)
        formXIXfile.save(filename=formXIXfinalfile)

    def create_ecard():

        ecardfilepath = os.path.join(karnatakafilespath,'Employment card.xlsx')
        ecardfile = load_workbook(filename=ecardfilepath)
        logging.info('Employment card file has sheet: '+str(ecardfile.sheetnames))
        sheetecard = ecardfile['Employment card']

        
        logging.info('create columns which are now available')

        data_ecard = data.drop_duplicates(subset=['Employee Code']).copy()
        data_ecard.fillna(value=0, inplace=True)

        emp_count = len(data_ecard.index)
        
        for i in range(0,emp_count):
            key = (data_ecard).iloc[i]['Employee Code']
            sheet_key = 'Employment card_'+str(key)

            emp_data = (data_ecard).iloc[i]
            emp_data.fillna(value='', inplace=True)

            sheet1 = ecardfile.copy_worksheet(sheetecard)
            sheet1.title = sheet_key
            sheet1['B4'] = contractor_name
            sheet1['B5'] = str(emp_data['Contractor_LIN'])+' / '+str(emp_data['Contractor_PAN'])
            sheet1['B6'] = emp_data['Contractor_email']
            sheet1['B7'] = emp_data['Contractor_mobile']
            sheet1['B7'].number_format= numbers.FORMAT_NUMBER
            sheet1['B8'] = emp_data['Nature of work']
            sheet1['B9'] = contractor_address
            if emp_data['PE_or_contract'][0]== 'PE':
                sheet1['B10'] = emp_data['Company Name']
            else:
                sheet1['B10'] = emp_data['Unit']
            sheet1['B11'] = str(emp_data['Unit_LIN'])+' / '+str(emp_data['Unit_PAN'])
            sheet1['B12'] = emp_data['Unit_email']
            sheet1['B13'] = emp_data['Unit_mobile']
            sheet1['B13'].number_format= numbers.FORMAT_NUMBER
            sheet1['B14'] = emp_data['Employee Name']
            sheet1['B15'] = emp_data['Aadhar Number']
            sheet1['B15'].number_format= numbers.FORMAT_NUMBER
            sheet1['B16'] = emp_data['Mobile Tel No.']
            sheet1['B16'].number_format= numbers.FORMAT_NUMBER
            sheet1['B17'] = key
            sheet1['B18'] = emp_data['Designation']
            sheet1['B19'] = emp_data['FIXED MONTHLY GROSS']
            sheet1['B20'] = emp_data['Date Joined']
            sheet1['B21'] = '-'
            

        ecardfinalfile = os.path.join(filelocation,'Employment card.xlsx')
        ecardfile.remove(sheetecard)
        ecardfile.save(filename=ecardfinalfile)
            
    try:
        create_form_A()
        create_form_B()
        create_form_XXI()
        create_form_XXII()
        create_form_XXIII()
        create_form_XX()
        create_wages()
        create_form_H_F('FORM H')
        create_form_H_F('FORM F')
        create_muster()
        create_formXIX()
        create_ecard()
    except KeyError as e:
        logging.info("Key error : Check if {} column exsists".format(e))
        print("Key error {}".format(e))
        report.configure(text="Failed: Check input file format  \n column {} not found".format(e))
        master.update()
        raise KeyError
    except FileNotFoundError as e:
        logging.info("File not found : Check if {} exsists".format(e))
        report.configure(text="Failed: File {} not found".format(e))
        master.update()
        raise FileNotFoundError