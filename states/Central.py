import calendar
import datetime
import logging
import os
import tkinter as tk
from functools import partial
from pathlib import Path
from tkinter import *
from tkinter import filedialog, ttk

import numpy as np
import pandas as pd

from states import Statefolder, logging, monthdict
from states.utils import forms_template

employee_name_column = "Employee Name"
fathers_name_column = "Father's Name"
gender_column = "Gender"
employee_code_column = "Employee Code"
contractor_name_column = 'Contractor_name'
department_column = 'Department'
fix_monthly_gross_column = "FIXED MONTHLY GROSS"
date_of_payment_column = 'Date of payment'
company_name_column = "Company Name"
address_column = "Address"
PE_or_contract_column = 'PE_or_contract'


def Central_Process(data, contractor_name, contractor_address, filelocation, month, year, report, master):
    Centralfilespath = os.path.join(Statefolder, 'Central')
    logging.info('Central files path is :'+str(Centralfilespath))
    data.reset_index(drop=True, inplace=True)
    month_num = monthdict[month]
    # Comment this line if in future leave file data is needed in any of functions below
    data = data.drop_duplicates(subset='Employee Code', keep="last")
    templates = forms_template.Templates(
            to_read=Centralfilespath, to_write=filelocation, month=month,
            year=year, report=report, master=master)

    def Form_C():

        data_formC = data.copy(deep=True)
        data_formC = data_formC.drop_duplicates(
            subset=employee_code_column, keep="last")
        columns = ['Employee Code', "Employee Name", "Recovery_Type", "Particulars", "Date of payment", "amount", "whether_show_cause_issue",
                   "explaination_heard_in_presence_of",
                                    "num_installments", "first_month_year", "last_month_year", "Date_of_complete_recovery", "remarks"]

        Recovery_Type_columns_name = ['Other Deduction', 'OtherDeduction1', 'OtherDeduction2',
                                      'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5', 'Damage or Loss', 'Fine', 'Salary Advance']

        data_formC["amount"] = templates.sum_columns_of_dataframe(
            data_formC, Recovery_Type_columns_name)
        data_formC["Recovery_Type"]=data_formC["amount"]
        data_formC[["Particulars", "whether_show_cause_issue", "explaination_heard_in_presence_of",
                    "num_installments", "first_month_year", "last_month_year", "Date_of_complete_recovery"]] = "---"
        data_formC["remarks"] = ""

        data_formC['S.no'] = list(range(1, len(data_formC)+1))
        data_formC['date of suspension'] = ""
        formC_data = templates.get_data(data_formC,columns)
        data_once_per_sheet = {'A4': str(data_formC['UnitName'].unique()[0])}

        templates.create_basic_form(filename='Form C Format of register of loan.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formC_data, start_row=8, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_I():
        data_formI = data.copy(deep=True)
        data_formI = data_formI.drop_duplicates(
            subset=employee_code_column, keep="last")
        columns = ['S.no', "Employee Name", "Father's Name", "Gender", "Department", "name&date_of_offence", "cause_against_fine", "FIXED MONTHLY GROSS",
                   "Date of payment_fine_released", "Date of payment_fine_imposed", "remarks"]

        data_formI['S.no'] = list(range(1, len(data_formI)+1))
        data_formI[["name&date_of_offence", "cause_against_fine"]] = "---"

        data_formI["Date of payment_fine_released"] = data_formI['Date of payment']
        data_formI["Date of payment_fine_imposed"] = data_formI['Date of payment']
        data_formI[["FIXED MONTHLY GROSS", "Date of payment_fine_released",
                    "Date of payment_fine_imposed"]] = "---"

        data_formI["remarks"] = ""
        formI_data = templates.get_data(data_formI, columns)
        data_once_per_sheet = {'A4': str(data_formI['UnitName'].unique()[0])}

        templates.create_basic_form(filename='Form I register of Fine.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formI_data, start_row=8, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_II_reg_damage_loss():
        data_formII = data.copy(deep=True)
        data_formII = data_formII.drop_duplicates(
            subset=employee_code_column, keep="last")
        columns = ['S.no', "Employee Name", "Father's Name", "Gender", "Department", "Damage or Loss", "whether_work_showed_cause",
                   "Date of payment & amount of deduction", "num_instalments", "Date of payment", "remarks"]

        data_formII['S.no'] = list(range(1, len(data_formII)+1))
        data_formII[["Damage or Loss", "whether_work_showed_cause", "Date of payment & amount of deduction",
                     "num_instalments", "Date of payment", "remarks"]] = "---"
        formII_data = templates.get_data(data_formII,columns)
        data_once_per_sheet = {'A4': str(data_formII['UnitName'].unique()[
                                         0]), 'A5': str(month)+" "+str(year)}

        templates.create_basic_form(filename='Form II Register of deductions for damage or loss.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formII_data, start_row=8, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_IV():
        data_formIV = data.copy(deep=True)
        data_formIV = data_formIV.drop_duplicates(
            subset=employee_code_column, keep="last")
        columns = ['S.no', "Employee Name", "Father's Name", "Gender", "Designation_Dept", "Date_overtime_worked",
                   "Extent of over-time", "Total over-time", "Normal hrs ",
                   "FIXED MONTHLY GROSS", "overtime rate",
                   "normal_earning", "Overtime", 'Total Earning', "date_overtime_paid"]

        # data_formIV['Total\r\nOT Hrs']=data_formIV[['Total\r\nOT Hrs',"Overtime",'Total Earning']].astype(float)
        # data_formIV["Total over-time"]=data_formIV['Total\r\nOT Hrs']
        # data_formIV["normal_earning"]=data_formIV['FIXED MONTHLY GROSS']-data_formIV["Overtime"]
        # data_formIV.loc[data_formIV['Total\r\nOT Hrs']==0,["Total over-time","Normal hrs ",
        #                                 "FIXED MONTHLY GROSS","overtime rate",
        #                                 "normal_earning","Overtime",'Total Earning']]="---"

        # data_formIV["date_overtime_paid"]=data_formIV['Date of payment']
        # data_formIV.loc[data_formIV["Overtime"]==0,"date_overtime_paid"]="---"
        # data_formIV.loc[data_formIV['Total\r\nOT Hrs']==0,"date_overtime_paid"]="---"
        # data_formIV["Extent of over-time"]="-----"
        # data_formIV["Date_overtime_worked"]="-----"
        data_formIV["Date_overtime_worked", "Extent of over-time", "Total over-time", "Normal hrs ",
                    "FIXED MONTHLY GROSS", "overtime rate", "normal_earning", "Overtime", 'Total Earning', "date_overtime_paid"] = "---"

        data_formIV['S.no'] = list(range(1, len(data_formIV)+1))
        data_formIV['Designation_Dept'] = templates.combine_columns_of_dataframe(data_formIV,["Designation","Department"])
        # data_formIV["Date of payment & amount of deduction"]=data_formIV['Date of payment']+"\n"+data_formIV["Total Deductions"]
        data_formIV[['Extent of over-time', 'Date_overtime_worked',
                     'date_overtime_paid', 'normal_earning', 'Total over-time']] = "---"

        formIV_data = templates.get_data(data_formIV,columns)
        data_once_per_sheet = {'A4': str(month)+" "+str(year)}

        templates.create_basic_form(filename='Form IV Overtime register.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formIV_data, start_row=8, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

        # formIVsheet['A4']=formIVsheet['A4'].value+"  "+data_formIV['Company Name'].unique()[0]+"  "+data_formIV['Company Address'].unique()[0]+"                                Month Ending: "+month+" "+str(year)
        
    def Form_V():
        data_formV = data.copy(deep=True)
        data_formV = data_formV.drop_duplicates(
            subset=employee_code_column, keep="last")
        
        columns = ['S.no', "Employee Name",
                   "Father's Name", "Gender", 'Nature of work']
        
        columns.extend(templates.get_attendance_columns(data_formV))
        columns.append('Total\r\nDP')
        data_formV['S.no'] = list(range(1, len(data_formV)+1))

        formV_data = templates.get_data(data_formV,columns)

        monthstart = datetime.date(year, month_num, 1)
        monthend = datetime.date(
            year, month_num, calendar.monthrange(year, month_num)[1])

        data_once_per_sheet = {'A4': str(data_formV['UnitName'].unique()[
                                         0]), 'A5': data_formV['Branch'].unique()[0],'A6':monthstart,'B6':monthend}

        templates.create_basic_form(filename='Form V Muster Roll.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formV_data, start_row=9, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

        #formPsheet['AE4']=formPsheet['AE4'].value+"   "+str(data_formP['Registration_no'].unique()[0])
        
    def Form_X():
        data_formX = data.copy(deep=True)
        data_formX = data_formX.drop_duplicates(
            subset=employee_code_column, keep="last")

        columns = ['S.no', "Employee Name", "Father's Name", "Designation", 'Earned Basic', 'DA',
                   'Earned Basic', 'DA', "Days Paid", "Overtime", "FIXED MONTHLY GROSS", "PF", "HRA",
                   "all_Other_Deduction_sum", 'Total Deductions', 'Net Paid', 'Date of payment',
                   'sign']

        other_deductions_columns_name = ['Other Deduction', 'OtherDeduction1', 'OtherDeduction2',
                                         'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5']
        
        data_formX["all_Other_Deduction_sum"] = templates.sum_columns_of_dataframe(data_formX,other_deductions_columns_name)

        data_formX[["sign"]] = ""

        data_formX['S.no'] = list(range(1, len(data_formX)+1))

        formX_data = templates.get_data(data_formX,columns)
        # formXsheet['P4']=formXsheet['P4'].value+"   "+str(data_formX['Registration_no'].unique()[0])
        # formXsheet['P5']=formXsheet['P5'].value+"   "+month
        monthstart = datetime.date(year, month_num, 1)
        monthend = datetime.date(
            year, month_num, calendar.monthrange(year, month_num)[1])
        data_once_per_sheet = {'A3': str(data_formX['UnitName'].unique()[0]),'A4':str(data_formX['Branch'].unique()[0]),
                                        'B5': monthstart,'C5':monthend}

        templates.create_basic_form(filename='Form X register of wages.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formX_data, start_row=9, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)
        
    def create_ecard():

        data_ecard = data.copy(deep=True)
        data_ecard = data_ecard.drop_duplicates(
            subset=employee_code_column, keep="last")

        data_ecard["employee_name_fathers_name"]=templates.combine_columns_of_dataframe(data_ecard,['Employee Name',"Father's Name"],"/")
        data_ecard["month_year"]=str(month)+" "+str(year)
        mapping={'B4':'UnitName','B5':'Location','B6':"employee_name_fathers_name",'B7':'Designation',
                        'B8':'month_year','B9':'FIXED MONTHLY GROSS','B10':'Earned Basic','B11':'DA',
                        'B12':'Days Paid','B13':"Overtime",'B14':"FIXED MONTHLY GROSS",'B15':'Total Deductions','B16':'Net Paid'}
        data_once_per_sheet = templates.get_data_once_persheet_peremployee(data_ecard,mapping=mapping)
        templates.create_per_employee_basic_form(filename='Form XI wages slip.xlsx',sheet_name="Sheet1",start_row=0,start_column=0,
                                        employee_codes=data_ecard[employee_code_column],data_once_per_sheet=data_once_per_sheet,per_employee_diff_data=True)
        
    try:
        Form_C()
        Form_I()
        Form_II_reg_damage_loss()
        Form_IV()
        Form_V()
        Form_X()
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
    finally:
        del templates