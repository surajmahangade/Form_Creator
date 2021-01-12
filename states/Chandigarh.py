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


def Chandigarh(data,contractor_name,contractor_address,filelocation,month,year):
    
    Chandigarhfilespath = os.path.join(Statefolder,'Chandigarh')
    logging.info('Chandigarh files path is :'+str(Chandigarhfilespath))
    data.reset_index(drop=True, inplace=True)
    month_num = monthdict[month]
    
    def Form_A():
    
        formAfilepath = os.path.join(Chandigarhfilespath,'Form A.xlsx')
        formAfile = load_workbook(filename=formAfilepath)
        logging.info('Form A file has sheet: '+str(formAfile.sheetnames))
        logging.info('create columns which are now available')

        data_formA = data.copy()
        columns=['S.no',"Employee Name","start_time","end_time",'interval_for_reset_from','interval_for_reset_to']
        
        data_formA['interval_for_reset_to']=data_formA.rest_interval.str.split("-",expand=True)[1]
        data_formA['interval_for_reset_from']=data_formA.rest_interval.str.split("-",expand=True)[0]


        data_formA['S.no'] = list(range(1,len(data_formA)+1))
        formA_data=data_formA[columns]
        formAsheet = formAfile['Sheet1']
        formAsheet.sheet_properties.pageSetUpPr.fitToPage = True
        logging.info('data for form A is ready')

        from openpyxl.utils.dataframe import dataframe_to_rows
        rows = dataframe_to_rows(formA_data, index=False, header=False)

        logging.info('rows taken out from data')
        formAsheet.delete_rows(15,4)
        formAsheet.insert_rows(15,len(data_formA))
        for r_idx, row in enumerate(rows, 15):
            for c_idx, value in enumerate(row, 1):
                formAsheet.cell(row=r_idx, column=c_idx, value=value)
                formAsheet.cell(row=r_idx, column=c_idx).font =Font(name ='Bell MT', size =10)
                formAsheet.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal='center', vertical='center', wrap_text = True)
                border_sides = Side(style='thin')
                formAsheet.cell(row=r_idx, column=c_idx).border = Border(outline= True, right=border_sides, bottom=border_sides)
        formAsheet['A4']=formAsheet['A4'].value+" : "+data_formA['Unit'][0]
        formAfinalfile = os.path.join(filelocation,'Form A.xlsx')
        formAfile.save(filename=formAfinalfile)

    Form_A()