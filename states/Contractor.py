from states import logging, monthdict, Statefolder
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
from states.utils import forms_template


def Contractor_Process(data, contractor_name, contractor_address, filelocation, month, year, report, master):
    Contractorfilespath = os.path.join(Statefolder, 'CLRA')
    logging.info('Contractor files path is :'+str(Contractorfilespath))
    data.reset_index(drop=True, inplace=True)
    month_num = monthdict[month]
    
    data = data.drop_duplicates(subset='Employee Code', keep="last")
    templates = forms_template.Templates(
        to_read=Contractorfilespath, to_write=filelocation, month=month, year=year, report=report, master=master)

    def create_form_A():
        data_formA = data.copy(deep=True)
        data_formA = data.drop_duplicates(subset=['Employee Code']).copy()

        data_formA['S.no'] = list(range(1, len(data_formA)+1))

        formA_columns = ["S.no", 'Employee Code', 'Employee Name', "Gender", "Father's Name", 'Date of Birth', "Nationality", "Education Level", 'Date Joined',
                         'Designation', 'Category Address', "Type of Employment", 'Mobile Tel No.', 'UAN Number', "PAN Number", "ESIC Number", "LWF EE", "Aadhar Number",
                         "Bank A/c Number", "Bank Name", 'Branch', "Present_Address", "Permanent_Address", 'Service Book No', "Date Left", "Reason for Leaving", 'Identification mark',
                         "photo", "sign", "remarks"]

        data_formA[["photo", "sign", "remarks"]] = ""

        persent_addr_columns = [
            'Local Address 1', 'Local Address 2', 'Local Address 3', 'Local Address 4']
        permanent_addr_columns = [
            'Permanent Address 1', 'Permanent Address 2', 'Permanent Address 3', 'Permanent Address 4']
        data_formA["Present_Address"] = templates.combine_columns_of_dataframe(
            data_formA, persent_addr_columns, " ")
        data_formA["Permanent_Address"] = templates.combine_columns_of_dataframe(
            data_formA, permanent_addr_columns, " ")

        formA_data = templates.get_data(data_formA, formA_columns)

        if data_formA['PE_or_contract'].unique()[0] == 'PE':
            A5_data = templates.combine_columns_of_dataframe(
                data_formA, ['Company Name', 'Company Address']).unique()[0]
        else:
            A5_data = templates.combine_columns_of_dataframe(
                data_formA, ['Contractor_name', 'Contractor_Address']).unique()[0]

        data_once_per_sheet = {'A5': A5_data}
        templates.create_basic_form(filename='Form A Employee register.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formA_data, start_row=11, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

        # company = formAsheet['A10'].value
        # A10_data = company+' '+data_formA['Unit'].unique()[0]+', '+data_formA['Branch'].unique()[0]
        # formAsheet['A10'] = A10_data

    def create_form_B():
        data_formB = data.copy(deep=True)
        data_formB = data.drop_duplicates(subset=['Employee Code']).copy()
        
        data_formB['basic_and_allo'] = templates.sum_columns_of_dataframe(data_formB,['Earned Basic','DA'])
        #data_formB['Other EAR'] = data_formB['Other Reimb']+data_formB['Arrears']+data_formB['Other Earning']+data_formB['Variable Pay']+data_formB['Stipend'] +data_formB['Consultancy Fees']
        # data_formB['VPF']=0
        data_formB['Society'] = "---"
        data_formB['Income Tax'] = "---"

        data_formB['Other Deduc'] = templates.sum_columns_of_dataframe(data_formB,['Other Deduction','Salary Advance'])
        data_formB['EMP PF'] = data_formB['PF']
        data_formB['Remarks'] = ''

        formB_columns = ['Employee Code', 'Employee Name', 'FIXED MONTHLY GROSS',	'Days Paid', 'Total\r\nOT Hrs',	'basic_and_allo', 'Overtime',
                         'HRA',	'Tel and Int Reimb', 'Bonus', 'Fuel Reimb',	'Prof Dev Reimb', 'Corp Attire Reimb',	'CCA',
                         'all_Other_Allowance', 'Total Earning', 'PF', 'VPF', 'P.Tax', 'Society', 'LWF EE', 'Insurance',
                         'TDS', 'advance+deductions', 'Recoveries', "Total Deductions", 'Net Paid', "PF", "Bank A/c Number",
                         'Date of payment', 'Remarks']

        
        all_other_allowance_columns = ['Other Allowance', 'OtherAllowance1',
                                       'OtherAllowance2', 'OtherAllowance3', 'OtherAllowance4', 'OtherAllowance5']

        
        
        data_formB['all_Other_Allowance'] = templates.sum_columns_of_dataframe(data_formB,all_other_allowance_columns)

        all_Other_deductions_advance_columns = ['Other Deduction', 'OtherDeduction1',
                                        'OtherDeduction2', 'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5','Salary Advance']
        
        data_formB['advance+deductions'] = templates.sum_columns_of_dataframe(data_formB,all_Other_deductions_advance_columns)
        
        data_formB['Recoveries'] = ""

        formB_data = templates.get_data(data_formB,formB_columns)

        if data_formB['PE_or_contract'].unique()[0] == 'PE':
            A10_data = templates.combine_columns_of_dataframe(data_formB,['Company Name','Company Address']).unique()[0]
        else:
            A10_data = templates.combine_columns_of_dataframe(data_formB,['Unit','Address']).unique()[0]

        A11_data = templates.combine_columns_of_dataframe(data_formB,['Unit','Address']).unique()[0]
            
        monthstart = datetime.date(year, month_num, 1)
        monthend = datetime.date(
            year, month_num, calendar.monthrange(year, month_num)[1])

        data_once_per_sheet = {'A8': contractor_name+', '+contractor_address,
                               'A9': templates.combine_columns_of_dataframe(data_formB, ['Nature of work', 'Location'], ', ').unique()[0],
                               'A10': A10_data,'A11':A11_data,
                               'A12': str(monthstart)+" "+str(monthend)}
        
        templates.create_basic_form(filename='Form B wage register equal remuniration.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formB_data, start_row=17, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_C():
        data_formC = data.copy(deep=True)
        data_formC = data.drop_duplicates(subset=['Employee Code']).copy()

        columns = ['Employee Code', "Employee Name", "Recovery_Type", "Particulars", "Date of payment and damage loss", 
                                    'Damage or Loss', "whether_show_cause_issue", "explaination_heard_in_presence_of",
                                    "num_installments", "first_month_year", "last_month_year", "Date_of_complete_recovery", "remarks"]

        Recovery_Type_columns_name = ['Other Deduction', 'OtherDeduction1', 'OtherDeduction2',
                                      'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5', 'Damage or Loss', 'Fine', 'Salary Advance']

        data_formC["Recovery_Type"] = templates.sum_columns_of_dataframe(data_formC,Recovery_Type_columns_name)

        data_formC["amount"] = data_formC["Recovery_Type"]
        data_formC[["Particulars", "whether_show_cause_issue", "explaination_heard_in_presence_of",
                    "num_installments", "first_month_year", "last_month_year", "Date_of_complete_recovery"]] = "---"

        data_formC["remarks"] = ""

        data_formC["Date of payment and damage loss"] = templates.combine_columns_of_dataframe(
                                            data_formC,["Date of payment",'Damage or Loss']," ")

        formC_data = templates.get_data(data_formC,columns)
        data_once_per_sheet = {'A4': str(data_formC['UnitName'].unique()[0])
                               }
        
        templates.create_basic_form(filename='Form C register of loan or recoveries.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formC_data, start_row=9, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_D():
        data_formD = data.copy(deep=True)
        data_formD = data.drop_duplicates(subset=['Employee Code']).copy()

        columns = ['S.no', "Employee Name", "Relay_or_set_work", 'Branch']

        columns.extend(templates.get_attendance_columns(data_formD))

        columns.extend(['in', 'out', 'Total\r\nDP', 'num_hours', 'sign'])

        data_formD[["Relay_or_set_work", "in",
                    "out", 'num_hours', 'sign']] = ""

        data_formD['S.no'] = list(range(1, len(data_formD)+1))

        formD_data = templates.get_data(data_formD,columns)
        

        data_once_per_sheet = {'A4': str(data_formD['UnitName'].unique()[0]),
                                'A5': str(data_formD['UnitName'].unique()[0])
                               }
        
        templates.create_basic_form(filename='Form D Register of attendance.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formD_data, start_row=12, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)


    def Form_E():
        
        data_formE = data.copy(deep=True)
        data_formE = data_formE.drop_duplicates(subset=['Employee Code']).copy()

        columns = ['Employee Code', "Employee Name", "Days Paid", "opening_bal", "added", "rest_allowed", "rest_availed",
                   "closing_bal", "Opening", "Monthly Increment", "Leave Accrued", "Closing",
                   "Monthly Increment", "Leave Accrued", "Closing", "openeing_bal", "added", "leave_availed", "closing_bal",
                   "remarks"]

        data_formE[["opening_bal", "added", "rest_allowed", "rest_availed",
                    "closing_bal", "openeing_bal", "added", "leave_availed", "closing_bal"]] = "---"
        data_formE["remarks"] = ""
        
        formE_data = templates.get_data(data_formE,columns)
        

        data_once_per_sheet = {'A5': str(formE_data['UnitName'].unique()[0]),
                                'A6': str(formE_data['UnitName'].unique()[0]),
                                'A8':str(month)+" "+str(year)
                               }
        
        templates.create_basic_form(filename='Form E Register of Rest,Leave,leave wages.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formE_data, start_row=13, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)


    def create_formXIX():
        data_formXIX = data.copy(deep=True)
        data_formXIX = data_formXIX.drop_duplicates(subset=['Employee Code']).copy()

        data_formXIX['contractor_name_and_address'] = templates.combine_columns_of_dataframe(data,['Contractor_name','Contractor_Address'],", ")
        data_formXIX['nature_location'] = templates.combine_columns_of_dataframe(data,['Nature of work','Location'],", ")
        if data_formXIX['PE_or_contract'].unique()[0] == 'PE':
            data_formXIX['unit_or_company'] = templates.combine_columns_of_dataframe(data,['Nature of work','Location'],", ")
        else:
            data_formXIX['unit_or_company'] = templates.combine_columns_of_dataframe(data,['Unit','Address'],", ")

        data_formXIX[['num_units_worked','rate_daily_wages']]=""
        mapping={'B4':'contractor_name_and_address',
                        'B5':'nature_location',
                        'B6':"unit_or_company",
                        'B7':'unit_or_company',
                        'B8':'month_year',
                        'B9':'Employee Code','B10':'Employee Name','B11':'Days Paid',
                        'B12':'num_units_worked','B13':'rate_daily_wages',
                        'B14':'Earned Basic','B15':'HRA','B16':'Tel and Int Reimb',
                        'B17':'Bonus','B18':'Fuel Reimb','B19':'Corp Attire Reimb','B20':'CCA',
                        'B21':'Total Earning',
                        'B22':'Insurance','B23':'PF','B24':'P.Tax','B25':'Total Deductions','B26':'Net Paid'
                        }

        data_once_per_sheet=templates.get_data_once_persheet_peremployee(data_formXIX,mapping)
        templates.create_per_employee_basic_form(filename='Form XIX Wages slip.xlsx',sheet_name='Sheet1',start_row=0,start_column=0,
                            employee_codes=data_formXIX['Employee Code'],data_once_per_sheet=data_once_per_sheet,
                            per_employee_diff_data=True
                                    )
        
    def create_formXV():
        data_formXV = data.copy(deep=True)
        data_formXV = data_formXV.drop_duplicates(subset=['Employee Code']).copy()

        data_formXV['contractor_name_and_address'] = templates.combine_columns_of_dataframe(data,['Contractor_name','Contractor_Address'],", ")
        data_formXV['nature_location'] = templates.combine_columns_of_dataframe(data,['Nature of work','Location'],", ")
        data_formXV['unit_address'] = templates.combine_columns_of_dataframe(data,['Unit','Address'],", ")

        data_formXV[['num_units_worked','rate_daily_wages']]=""
        data_formXV["one"]="1"
        monthstart = datetime.date(year, month_num, 1)
        monthend = datetime.date(
                year, month_num, calendar.monthrange(year, month_num)[1])
        data_formXV["monthstart"]=monthstart
        data_formXV["monthend"]=monthend

        mapping={'A5':'contractor_name_and_address',
                'A6':'unit_address','A7':'Location','A8':'unit_address','A9':'unit_address',
                'A10':'Age','A11':'Identification mark','A12':"Father's Name",'A18':"one",'B18':"monthstart",'C18':"monthend",
                'D18':'Designation'}

        data_once_per_sheet=templates.get_data_once_persheet_peremployee(data_formXV,mapping)
        templates.create_per_employee_basic_form(filename='Form XV Service certificate.xlsx',sheet_name='Sheet1',start_row=0,start_column=0,
                            employee_codes=data_formXV['Employee Code'],data_once_per_sheet=data_once_per_sheet,
                            per_employee_diff_data=True
                                    )


    def create_form_XX():
        
        data_formXX = data.copy(deep=True)
        data_formXX = data_formXX.drop_duplicates(subset=['Employee Code']).copy()

        data_formXX['S.no'] = list(range(1, len(data_formXX)+1))

        data_formXX['c'] = '---'
        data_formXX['d'] = '---'
        data_formXX['f'] = '---'
        data_formXX['g'] = '---'
        data_formXX['h'] = '---'
        data_formXX['i'] = ''

        formXX_columns = ['S.no', 'Employee Name', "Father's Name", 'Designation', 'Damage or Loss',
                          "Date of payment", 'c', 'd', "all_Other_Deduction_sum", 'f', 'g', 'h', 'i']

        other_deductions_columns_name = ['Other Deduction', 'OtherDeduction1', 'OtherDeduction2',
                                         'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5']


        data_formXX["all_Other_Deduction_sum"] = templates.sum_columns_of_dataframe(data_formXX,
                                                                other_deductions_columns_name)

        
        formXX_data = templates.get_data(data_formXX,formXX_columns)
        
        data_formXX['contractor_name_and_address'] = templates.combine_columns_of_dataframe(data,['Contractor_name','Contractor_Address'],", ")
        data_formXX['nature_location'] = templates.combine_columns_of_dataframe(data,['Nature of work','Location'],", ")
        data_formXX['unit_address'] = templates.combine_columns_of_dataframe(data,['Unit','Address'],", ")

        data_once_per_sheet = {'A5': 'contractor_name_and_address',
                                'A6': 'nature_location','A7':'unit_address',
                                'A8': 'unit_address'
                               }
        
        templates.create_basic_form(filename='Form XX Register of Deduction for damage or loss.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formXX_data, start_row=13, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def create_form_XXI():
        data_formXXI = data.copy(deep=True)
        data_formXXI = data_formXXI.drop_duplicates(subset=['Employee Code']).copy()

        data_formXXI['S.no'] = list(range(1, len(data_formXXI)+1))

        data_formXXI['a'] = '---'
        data_formXXI['b'] = '---'
        data_formXXI['c'] = '---'
        #data_formXXI['e'] ='---'
        data_formXXI['f'] = '---'
        data_formXXI['g'] = ''

        formXXI_columns = ['S.no', 'Employee Name', "Father's Name", 'Designation', 'a', 'Date of payment',
                           'c', 'Employee Name', 'Date of payment and FIXED MONTHLY GROSS', 'Fine', 'Date of payment', 'g']

        data_formXXI['Date of payment and FIXED MONTHLY GROSS'] = templates.combine_columns_of_dataframe(data_formXXI,['Date of payment','FIXED MONTHLY GROSS']," /")
        formXXI_data = templates.get_data(data_formXXI,formXXI_columns)

        data_formXXI['contractor_name_and_address'] = templates.combine_columns_of_dataframe(data,['Contractor_name','Contractor_Address'],", ")
        data_formXXI['nature_location'] = templates.combine_columns_of_dataframe(data,['Nature of work','Location'],", ")
        data_formXXI['unit_address'] = templates.combine_columns_of_dataframe(data,['Unit','Address'],", ")

        data_once_per_sheet = {'A5': 'contractor_name_and_address','A6':'nature_location',
                                'A7':'unit_address','A8':'unit_address'}

        templates.create_basic_form(filename='Form XXI register of fine.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formXXI_data, start_row=12, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def create_form_XXII():
        'Form XXII Register of Advances.xlsx'
        
        data_formXXII = data.copy(deep=True)

        data_formXXII['S.no'] = list(range(1, len(data_formXXII)+1))

        data_formXXII['c'] = '---'
        data_formXXII['d'] = '---'
        data_formXXII['e'] = '---'
        data_formXXII['f'] = '---'
        data_formXXII['g'] = ''

        formXXII_columns = ['S.no', 'Employee Name', "Father's Name", 'Designation',
                            'FIXED MONTHLY GROSS', 'Date of payment', 'c', 'd', 'e', 'f', 'g']

        formXXII_data = templates.get_data(data_formXXII,formXXII_columns)

        for r_idx, row in enumerate(rows, 12):
            for c_idx, value in enumerate(row, 1):
                formXXIIsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXIIsheet.cell(row=r_idx, column=c_idx).font = Font(
                    name='Verdana', size=8)
                formXXIIsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    horizontal='center', vertical='center', wrap_text=True)
                border_sides = Side(style='thin')
                formXXIIsheet.cell(row=r_idx, column=c_idx).border = Border(
                    outline=True, right=border_sides, bottom=border_sides)

        contractline = formXXIIsheet['A5'].value
        A5_data = contractline+' '+contractor_name+', '+contractor_address
        formXXIIsheet['A5'] = A5_data

        if str(data_formXXII['Nature of work'].dtype)[0:3] != 'obj':
            data_formXXII['Nature of work'] = data_formXXII['Nature of work'].astype(
                str)

        if str(data_formXXII['Location'].dtype)[0:3] != 'obj':
            data_formXXII['Location'] = data_formXXII['Location'].astype(str)

        if str(data_formXXII['Company Name'].dtype)[0:3] != 'obj':
            data_formXXII['Company Name'] = data_formXXII['Company Name'].astype(
                str)

        if str(data_formXXII['Company Address'].dtype)[0:3] != 'obj':
            data_formXXII['Company Address'] = data_formXXII['Company Address'].astype(
                str)

        if str(data_formXXII['Unit'].dtype)[0:3] != 'obj':
            data_formXXII['Unit'] = data_formXXII['Unit'].astype(str)

        if str(data_formXXII['Address'].dtype)[0:3] != 'obj':
            data_formXXII['Address'] = data_formXXII['Address'].astype(str)

        locationline = formXXIIsheet['A6'].value
        A6_data = locationline+' ' + \
            data_formXXII['Nature of work'].unique()[0]+', ' + \
            data_formXXII['Location'].unique()[0]
        formXXIIsheet['A6'] = A6_data

        establine = formXXIIsheet['A7'].value
        A7_data = establine+' ' + \
            data_formXXII['Unit'].unique()[0]+', ' + \
            data_formXXII['Address'].unique()[0]
        formXXIIsheet['A7'] = A7_data

        establine = formXXIIsheet['A8'].value
        A8_data = establine+' ' + \
            data_formXXII['Unit'].unique()[0]+', ' + \
            data_formXXII['Address'].unique()[0]
        formXXIIsheet['A8'] = A8_data

        # border the region
        count1 = len(data_formXXII)
        border_1 = Side(style='thick')
        # for i in range(1,12):
        #     formXXIIsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
        #     formXXIIsheet.cell(row=count1+16, column=i).border = Border(outline= True, bottom=border_1)
        # for i in range(1,count1+14):
        #     formXXIIsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
        #     formXXIIsheet.cell(row=i, column=14).border = Border(outline= True, right=border_1)

        formXXIIfinalfile = os.path.join(
            filelocation, 'Form XXII Register of Advances.xlsx')
        formXXIIfile.save(filename=formXXIIfinalfile)

    def create_form_XXIII():
        formXXIIIfilepath = os.path.join(
            Contractorfilespath, 'Form XXIII register of overtime.xlsx')
        formXXIIIfile = load_workbook(filename=formXXIIIfilepath)
        logging.info('Form XXIII file has sheet: ' +
                     str(formXXIIIfile.sheetnames))

        logging.info('create columns which are now available')
        data_formXXIII = data.copy(deep=True)

        data_formXXIII.fillna(value=0, inplace=True)

        data_formXXIII['S.no'] = list(range(1, len(data_formXXIII)+1))

        data_formXXIII['a'] = '---'
        data_formXXIII['g'] = ''

        formXXIII_columns = ['S.no', 'Employee Name', "Father's Name", 'Gender', 'Designation',
                             'Date of payment', 'a', 'FIXED MONTHLY GROSS', "overtime rate", "Overtime", 'Date of payment', 'g']

        formXXIII_data = data_formXXIII[formXXIII_columns]

        formXXIIIsheet = formXXIIIfile['Sheet1']

        formXXIIIsheet.sheet_properties.pageSetUpPr.fitToPage = True

        logging.info('data for form XXIII is ready')

        rows = dataframe_to_rows(formXXIII_data, index=False, header=False)

        logging.info('rows taken out from data')

        for r_idx, row in enumerate(rows, 12):
            for c_idx, value in enumerate(row, 1):
                formXXIIIsheet.cell(row=r_idx, column=c_idx, value=value)
                formXXIIIsheet.cell(row=r_idx, column=c_idx).font = Font(
                    name='Verdana', size=8)
                formXXIIIsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    horizontal='center', vertical='center', wrap_text=True)
                border_sides = Side(style='thin')
                formXXIIIsheet.cell(row=r_idx, column=c_idx).border = Border(
                    outline=True, right=border_sides, bottom=border_sides)

        contractline = formXXIIIsheet['A5'].value
        A5_data = contractline+' '+contractor_name+', '+contractor_address
        formXXIIIsheet['A5'] = A5_data

        if str(data_formXXIII['Nature of work'].dtype)[0:3] != 'obj':
            data_formXXIII['Nature of work'] = data_formXXIII['Nature of work'].astype(
                str)

        if str(data_formXXIII['Location'].dtype)[0:3] != 'obj':
            data_formXXIII['Location'] = data_formXXIII['Location'].astype(str)

        if str(data_formXXIII['Company Name'].dtype)[0:3] != 'obj':
            data_formXXIII['Company Name'] = data_formXXIII['Company Name'].astype(
                str)

        if str(data_formXXIII['Company Address'].dtype)[0:3] != 'obj':
            data_formXXIII['Company Address'] = data_formXXIII['Company Address'].astype(
                str)

        if str(data_formXXIII['Unit'].dtype)[0:3] != 'obj':
            data_formXXIII['Unit'] = data_formXXIII['Unit'].astype(str)

        if str(data_formXXIII['Address'].dtype)[0:3] != 'obj':
            data_formXXIII['Address'] = data_formXXIII['Address'].astype(str)

        locationline = formXXIIIsheet['A6'].value
        A6_data = locationline+' ' + \
            data_formXXIII['Nature of work'].unique()[0]+', ' + \
            data_formXXIII['Location'].unique()[0]
        formXXIIIsheet['A6'] = A6_data

        establine = formXXIIIsheet['A7'].value
        A7_data = establine+' ' + \
            data_formXXIII['Unit'].unique()[0]+', ' + \
            data_formXXIII['Address'].unique()[0]
        formXXIIIsheet['A7'] = A7_data

        establine = formXXIIIsheet['A8'].value
        A8_data = establine+' ' + \
            data_formXXIII['Unit'].unique()[0]+', ' + \
            data_formXXIII['Address'].unique()[0]
        formXXIIIsheet['A8'] = A8_data

        # border the region
        count1 = len(data_formXXIII)
        border_1 = Side(style='thick')
        # for i in range(1,13):
        #     formXXIIIsheet.cell(row=2, column=i).border = Border(outline= True, top=border_1)
        #     formXXIIIsheet.cell(row=count1+13, column=i).border = Border(outline= True, bottom=border_1)
        # for i in range(1,count1+12):
        #     formXXIIIsheet.cell(row=i, column=2).border = Border(outline= True, left=border_1)
        #     formXXIIIsheet.cell(row=i, column=14).border = Border(outline= True, right=border_1)

        formXXIIIfinalfile = os.path.join(
            filelocation, 'Form XXIII register of overtime.xlsx')
        formXXIIIfile.save(filename=formXXIIIfinalfile)

    def create_ecard():

        ecardfilepath = os.path.join(
            Contractorfilespath, 'FormXII Employment Card.xlsx')
        ecardfile = load_workbook(filename=ecardfilepath)
        logging.info('Employment card file has sheet: ' +
                     str(ecardfile.sheetnames))
        sheetecard = ecardfile['Sheet1']

        logging.info('create columns which are now available')
        data_ecard = data.copy(deep=True)
        data_ecard.fillna(value=0, inplace=True)

        emp_count = len(data_ecard.index)

        for i in range(0, emp_count):
            key = (data_ecard).iloc[i]['Employee Code']
            sheet_key = 'Employment card_'+str(key)

            emp_data = (data_ecard).iloc[i]

            sheet1 = ecardfile.copy_worksheet(sheetecard)
            sheet1.title = sheet_key
            sheet1['B4'] = contractor_name
            sheet1['B5'] = str(emp_data['Contractor_LIN']) + \
                ' / '+str(emp_data['Contractor_PAN'])
            sheet1['B6'] = emp_data['Contractor_email']
            sheet1['B7'] = emp_data['Contractor_mobile']
            sheet1['B7'].number_format = numbers.FORMAT_NUMBER
            sheet1['B8'] = emp_data['Nature of work']
            sheet1['B9'] = contractor_address
            sheet1['B10'] = emp_data['Unit']
            sheet1['B11'] = str(emp_data['Unit_LIN']) + \
                ' / '+str(emp_data['Unit_PAN'])
            sheet1['B12'] = emp_data['Unit_email']
            sheet1['B13'] = emp_data['Unit_mobile']
            sheet1['B13'].number_format = numbers.FORMAT_NUMBER
            sheet1['B14'] = emp_data['Employee Name']
            sheet1['B15'] = emp_data['Aadhar Number']
            sheet1['B15'].number_format = numbers.FORMAT_NUMBER
            sheet1['B16'] = emp_data['Mobile Tel No.']
            sheet1['B16'].number_format = numbers.FORMAT_NUMBER
            sheet1['B17'] = key
            sheet1['B18'] = emp_data['Designation']
            sheet1['B19'] = ""  # emp_data['FIXED MONTHLY GROSS']
            sheet1['B20'] = emp_data['Date Joined']
            sheet1['B21'] = '-'

        ecardfinalfile = os.path.join(
            filelocation, 'FormXII Employment Card.xlsx')
        ecardfile.remove(sheetecard)
        ecardfile.save(filename=ecardfinalfile)

    try:
        create_form_A()
        create_form_B()
        Form_C()
        Form_D()
        Form_E()
        create_formXIX()
        create_formXV()
        create_form_XX()
        create_form_XXI()
        create_form_XXII()
        create_form_XXIII()
        create_ecard()
    except KeyError as e:
        logging.info("Key error : Check if {} column exsists".format(e))
        print("Key error {}".format(e))
        report.configure(
            text="Failed: Check input file format  \n column {} not found".format(e))
        master.update()
        raise KeyError
    except FileNotFoundError as e:
        logging.info("File not found : Check if {} exsists".format(e))
        report.configure(text="Failed: File {} not found".format(e))
        master.update()
        raise FileNotFoundError
