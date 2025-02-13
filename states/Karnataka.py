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


def Karnataka(data, contractor_name, contractor_address, filelocation, month, year, report, master):
    karnatakafilespath = os.path.join(Statefolder, 'Karnataka')
    logging.info('karnataka files path is :'+str(karnatakafilespath))
    data.reset_index(drop=True, inplace=True)

    month_num = monthdict[month]

    templates = forms_template.Templates(
        to_read=karnatakafilespath, to_write=filelocation, month=month, year=year, report=report, master=master)

    def Form_F():
        leave_file_data = data[[employee_code_column, employee_name_column, "Leave Type", "Opening",
                                "Monthly Increment", "Leave Accrued", "Used", "Encash", "Closing"]]

        data_formF = data.drop_duplicates(subset=['Employee Code']).copy()

        columns = ['S.no', 'month_start', 'month_end', "Days Paid", "Monthly Increment", "open_balance", "from", "to", "numdays", "colsing_bal",
                   "Date_of_payement_fixed_monthly", "remarks", "slno", "year", "of_credit", "availed", "Balance"]

        for employee_name_leave_file in data_formF[employee_name_column]:
            # opening
            emp_details = leave_file_data.loc[leave_file_data[employee_name_column]
                                              == employee_name_leave_file, :]
            opening_pl = emp_details["Opening"].loc[emp_details["Leave Type"] == "PL"].replace("",0).astype(
                float)
            opening_cl = emp_details["Opening"].loc[emp_details["Leave Type"] == "CL"].replace("",0).astype(
                float)
            opening_sl = emp_details["Opening"].loc[emp_details["Leave Type"] == "SL"].replace("",0).astype(
                float)
            prev_bal = opening_pl.add(opening_cl.add(
                opening_sl, fill_value=0), fill_value=0).sum()

            data_formF.loc[data_formF[employee_name_column] ==
                           employee_name_leave_file, "open_balance"] = prev_bal

            # closing
            Closing_pl = emp_details["Closing"].loc[emp_details["Leave Type"] == "PL"].replace("",0).astype(
                float)
            Closing_cl = emp_details["Closing"].loc[emp_details["Leave Type"] == "CL"].replace("",0).astype(
                float)
            Closing_sl = emp_details["Closing"].loc[emp_details["Leave Type"] == "SL"].replace("",0).astype(
                float)
            closing = Closing_cl.add(Closing_pl.add(
                Closing_sl, fill_value=0), fill_value=0).sum()

            data_formF.loc[data_formF[employee_name_column] ==
                           employee_name_leave_file, "colsing_bal"] = closing

        start_month = datetime.date(year, month_num, 1)
        end_month = datetime.date(
            year, month_num, calendar.monthrange(year, month_num)[1])
        data_formF['S.no'] = list(range(1, len(data_formF)+1))
        data_formF['month_start'] = start_month
        data_formF['month_end'] = end_month
        data_formF["Date_of_payement_fixed_monthly"] = templates.combine_columns_of_dataframe(data_formF,
                                                                                              [date_of_payment_column, fix_monthly_gross_column])
        data_formF["remarks"] = ""
        data_formF[["slno", "year", "of_credit", "availed", "Balance"]] = "---"
        data_formF["permanent_address"] = templates.combine_columns_of_dataframe(data_formF, ["Permanent Address 1", "Permanent Address 2",
                                                                                              "Permanent Address 3", "Permanent Address 4"], " ")

        data_once_per_sheet = templates.get_data_once_persheet_peremployee(data_formF, {"A4": employee_code_column, "A5": "Date Joined",
                                                                                        "A6": employee_name_column, "A7": fathers_name_column, "A8": "permanent_address"})

        data_with_attendance = templates.get_from_to_dates_attendance(
            data_formF, "PL", sno_column="S.no")
        templates.create_attendance_form_per_employee(filename="Form F register of leave with wages.xlsx", sheet_name="Sheet1",
                                                      start_row=13, start_column=1,
                                                      data_with_attendance=data_with_attendance, columns=columns, data_once_per_sheet=data_once_per_sheet, per_employee_diff_data=True)

    def Form_H():
        leave_file_data = data[[employee_code_column, employee_name_column, "Leave Type", "Opening",
                                "Monthly Increment", "Leave Accrued", "Used", "Encash", "Closing"]]

        data_formH = data.drop_duplicates(subset=['Employee Code']).copy()

        columns = ['S.no', 'month_start', 'month_end', "Days Paid", "Monthly Increment", "open_balance", "from", "to", "numdays", "colsing_bal",
                   "Date_of_payement_fixed_monthly", "remarks", "slno", "year", "of_credit", "availed", "Balance"]

        for employee_name_leave_file in data_formH[employee_name_column]:
            # opening
            emp_details = leave_file_data.loc[leave_file_data[employee_name_column]
                                              == employee_name_leave_file, :]
            opening_pl = emp_details["Opening"].loc[emp_details["Leave Type"] == "PL"].replace("",0).astype(
                float)
            opening_cl = emp_details["Opening"].loc[emp_details["Leave Type"] == "CL"].replace("",0).astype(
                float)
            opening_sl = emp_details["Opening"].loc[emp_details["Leave Type"] == "SL"].replace("",0).astype(
                float)
            prev_bal = opening_pl.add(opening_cl.add(
                opening_sl, fill_value=0), fill_value=0).sum()

            data_formH.loc[data_formH[employee_name_column] ==
                           employee_name_leave_file, "open_balance"] = prev_bal

            # closing
            Closing_pl = emp_details["Closing"].loc[emp_details["Leave Type"] == "PL"].replace("",0).astype(
                float)
            Closing_cl = emp_details["Closing"].loc[emp_details["Leave Type"] == "CL"].replace("",0).astype(
                float)
            Closing_sl = emp_details["Closing"].loc[emp_details["Leave Type"] == "SL"].replace("",0).astype(
                float)
            closing = Closing_cl.add(Closing_pl.add(
                Closing_sl, fill_value=0), fill_value=0).sum()

            data_formH.loc[data_formH[employee_name_column] ==
                           employee_name_leave_file, "colsing_bal"] = closing

        start_month = datetime.date(year, month_num, 1)
        end_month = datetime.date(
            year, month_num, calendar.monthrange(year, month_num)[1])
        data_formH['S.no'] = list(range(1, len(data_formH)+1))
        data_formH['month_start'] = start_month
        data_formH['month_end'] = end_month
        data_formH["Date_of_payement_fixed_monthly"] = templates.combine_columns_of_dataframe(data_formH,
                                                                                              [date_of_payment_column, fix_monthly_gross_column])
        data_formH["remarks"] = ""
        data_formH[["slno", "year", "of_credit", "availed", "Balance"]] = "---"
        data_formH["permanent_address"] = templates.combine_columns_of_dataframe(data_formH, ["Permanent Address 1", "Permanent Address 2",
                                                                                              "Permanent Address 3", "Permanent Address 4"], " ")

        data_once_per_sheet = templates.get_data_once_persheet_peremployee(data_formH, {"A4": employee_code_column, "A5": "Date Joined",
                                                                                        "A6": employee_name_column, "A7": fathers_name_column, "A8": "permanent_address"})
        data_with_attendance = templates.get_from_to_dates_attendance(
            data_formH, "PL", sno_column="S.no")
        templates.create_attendance_form_per_employee(filename="Form H leave with wages.xlsx", sheet_name="Sheet1",
                                                      start_row=13, start_column=1,
                                                      data_with_attendance=data_with_attendance, columns=columns, data_once_per_sheet=data_once_per_sheet, per_employee_diff_data=True)

    def Form_T():
        logging.info('create columns which are now available')
        data_formT = data.copy(deep=True)
        data_formT = data_formT.drop_duplicates(
            subset=employee_code_column, keep="last")

        columns = ['S.no', employee_code_column, employee_name_column, fathers_name_column, gender_column, "Designation",
                   department_column, address_column, "Date Joined", "ESIC Number", 'PF Number', "VDA", "Days Paid",
                   'Total\r\nOT Hrs', 'basic_da', 'Earned Basic', 'HRA', 'Bonus', 'Special Allowance', 'Overtime',
                   'NFH', 'maternity', 'Telephone Reimb', 'Bonus', 'Fuel Reimb', 'Prof Dev Reimb', 'Corp Attire Reimb',
                   'CCA', 'Others', 'subsistence', 'Leave Encashment', 'Total Earning', 'ESIC Number', 'PF', 'P.Tax', 'TDS', 'Society', 'Insurance',
                   'FIXED MONTHLY GROSS', 'Fine', 'Damage or Loss', 'other_deductions', 'Total Deductions', 'Net Paid', 'BankName', 'Bank A/c Number',
                   'Account Code', 'fixed', 'sign']

        data_formT['S.no'] = list(range(1, len(data_formT)+1))
        other_deductions = ['Other Deduction', 'OtherDeduction1', 'OtherDeduction2',
                            'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5']
        data_formT['other_deductions'] = templates.sum_columns_of_dataframe(
            data_formT, other_deductions)
        data_formT['basic_da'] = templates.sum_columns_of_dataframe(
            data_formT, ['Earned Basic', 'DA'])
        data_formT[['Society', 'Others', 'subsistence',
                    'NFH', 'VDA', 'maternity', 'fixed', 'sign']] = ""
        formI_data = templates.get_data(data_formT,columns)
        data_once_per_sheet = {'A8': templates.combine_columns_of_dataframe(data_formT, ['Unit', address_column], ', ').unique()[0],
                               'A9': templates.combine_columns_of_dataframe(data_formT, ['Unit', address_column], ', ').unique()[0],
                               'A10': contractor_name,
                               'A11': templates.combine_columns_of_dataframe(data_formT, ['Nature of work', 'Location'], ', ').unique()[0]}
        templates.create_basic_form(filename='Form T Combine Muster roll cum register of wages.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formI_data, start_row=16, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    def Form_T_muster():
        data_formT = data.copy(deep=True)
        data_formT = data_formT.drop_duplicates(
            subset=employee_code_column, keep="last")

        columns = ['S.no', employee_code_column, employee_name_column]
        columns.extend(templates.get_attendance_columns(data_formT))
        columns.extend(['date of suspension', 'Days Paid', 'Total\r\nOT Hrs'])
        data_formT['S.no'] = list(range(1, len(data_formT)+1))
        data_formT['date of suspension'] = ""
        formI_data = templates.get_data(data_formT,columns)
        data_once_per_sheet = {'A8': templates.combine_columns_of_dataframe(data_formT, ['Unit', address_column], ', ').unique()[0],
                               'A9': templates.combine_columns_of_dataframe(data_formT, ['Unit', address_column], ', ').unique()[0],
                               'A10': contractor_name,
                               'AI4': month+" "+str(year)}
        templates.create_basic_form(filename='Form T muster roll.xlsx',
                                    sheet_name='Sheet1', all_employee_data=formI_data, start_row=14, start_column=1,
                                    data_once_per_sheet=data_once_per_sheet)

    try:
        Form_F()
        master.update()
        Form_H()
        master.update()
        Form_T()
        master.update()
        Form_T_muster()
        master.update()
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
