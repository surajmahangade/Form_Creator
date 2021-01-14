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
import logging

master = Tk()
master.title("Form Creator")
master.minsize(640,400)
Testing=True
from states import Register_folder,logfolder,Statefolder,State_forms,dbfolder,systemdrive,monthdict
from states import Goa,Karnataka,Delhi,Maharashtra,Kerala,Gujarat,Madhya_Pradesh,Haryana,Chandigarh,Central,Contractor,Hyderabad,Tamilnadu
Tamilnadu=Tamilnadu.Tamilnadu
Madhya_Pradesh=Madhya_Pradesh.Madhya_Pradesh
Goa=Goa.Goa
Karnataka=Karnataka.Karnataka
Chandigarh=Chandigarh.Chandigarh
Delhi=Delhi.Delhi
Maharashtra=Maharashtra.Maharashtra
Kerala=Kerala.Kerala
Gujarat=Gujarat.Gujarat
Haryana=Haryana.Haryana
Central_Process=Central.Central_Process
Contractor_Process=Contractor.Contractor_Process
Hyderabad=Hyderabad.Hyderabad

#backend code starts here
log_filename = datetime.datetime.now().strftime(os.path.join(logfolder,'logfile_%d_%m_%Y_%H_%M_%S.log'))
logging.basicConfig(filename=log_filename, level=logging.INFO)

def create_pdf(folderlocation,file_name):
    import win32com.client
    from pywintypes import com_error
    excel_filename = file_name
    pdf_filename = file_name.split('.')[0]+'.pdf'
    # Path to original excel file
    WB_PATH=os.path.join(folderlocation,excel_filename)
    # PDF path when saving
    PATH_TO_PDF =os.path.join(folderlocation,pdf_filename)

    logging.info(WB_PATH)
    logging.info(PATH_TO_PDF)

    excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = False

    try:
        logging.info('Start conversion to PDF')

        # Open
        wb = excel.Workbooks.Open(WB_PATH)

        sheetnumbers= len(pd.ExcelFile(WB_PATH).sheet_names)

        # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        ws_index_list = list(range(1,sheetnumbers+1))
        wb.WorkSheets(ws_index_list).Select()

        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        logging.info('failed.')
    else:
        logging.info('Succeeded.')
    finally:
        wb.Close()
        excel.Quit()

def Rajasthan(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    logging.info("Rajasthan form creation")


def Telangana(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    logging.info('Telangana forms')

def Uttar_Pradesh(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    logging.info('Uttar Pradesh forms')


def West_Bengal(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    logging.info("West_Bengal form creation")

def Uttarakhand(data,contractor_name,contractor_address,filelocation,month,year,report,master):
    logging.info("Uttarakhand form creation")

State_Process = {'delhi':Delhi,'telangana':Telangana,'uttar pradesh':Uttar_Pradesh,'goa':Goa,
                'gujarat':Gujarat,'kerala':Kerala,'madhya pradesh':Madhya_Pradesh,'rajasthan':Rajasthan,'haryana':Haryana,
                'west bengal':West_Bengal,'uttarakhand':Uttarakhand,'hyderabad':Hyderabad,'karnataka':Karnataka,'maharashtra':Maharashtra}#'tamilnadu':Tamilnadu,


companylist = ['SVR LTD','PRY Wine Ltd','CDE Technology Ltd']

def Type5(inputfolder,month,year):
    logging.info('type5 data process running')

def Type4(inputfolder,month,year):
    logging.info('type4 data process running')

def Type3(inputfolder,month,year):
    logging.info('type3 data process running')

def Type2(inputfolder,month,year):
    logging.info('type2 data process running')

def Type1(inputfolder,month,year):
    global output_text
    output_text=''
    logging.info('type1 data process running')

    emp_df_columns = ['Employee Code_master', 'Employee Name_master', 'Company Name','Company Address', 'Grade', 'Branch_master',
       'Department', 'Designation_master', 'Division', 'Group', 'Category', 'Unit',
       'Location Code', 'State', 'Date of Birth', 'Date Joined_master',
       'Date of Confirmation', 'Date Left', 'Title', 'Last Inc. Date',
       'Ticket Number', 'Local Address 1', 'Local Address 2',
       'Local Address 3', 'Local Address 4', 'Local City Name',
       'Local District Name', 'Local PinCode', 'Local State Name',
       'Residence Tel No.', 'Permanent Address 1', 'Permanent Address 2',
       'Permanent Address 3', 'Permanent Address 4', 'UAN Number_master',
       'Permanent Tel No.', 'Office Tel No.', 'Extension Tel No.',
       'Mobile Tel No.', "Father's Name", 'Gender', 'Age', 'Number of Months',
       'Marital Status', 'PT Number', 'PF Number (Old Version)', 'PF Number',
       'PF Number (WithComPrefix)', 'PAN Number', 'ESIC Number (Old Version)',
       'ESIC Number_master', 'ESIC Number (CompPrefix)', 'FPF Number', 'PF Flag',
       'ESIC Flag', 'PT Flag', 'Bank A/c Number_master', 'Bank Name', 'Mode',
       'Account Code_master', 'E-Mail_master', 'Remarks_master', 'PF Remarks', 'ESIC Remarks',
       'ESIC IMP Code', 'ESIC IMP Name', 'Employee Type (For PF)',
       'Freeze Account', 'Freeze Date', 'Freeze Reason', 'Type of House (In)',
       'Comp. adn.', 'Staying In (Metro Type)', 'Children (For CED)',
       'TDS Rate', 'Resignation Date', 'Reason for Leaving', 'Bank A/C No.1',
       'Bank A/C No.2', 'Bank A/C No.3', 'Alt.Email', 'Emp Status',
       'Probation Date', 'Surcharge Flag', 'Gratuity Code',
       'Resign Offer Date', 'Permanent City', 'Permanent District',
       'Permanent Pin Code', 'Permanent State', 'Spouse Name',
       'PF Joining Date', 'PRAN Number', 'Group Joining Date', 'Aadhar Number',
       'Child in Hostel (For CED)', 'Total Exp in Years', 'P', 'L',
       'Identification mark','Nationality',	'Education Level',	'Category Address',
       'Type of Employment',	'Service Book No',	'Nature of work']

    salary_df_columns = ['Sr', 'DivisionName', 'Sal Status', 'Emp Code_salary', 'Emp Name_salary', 'Designation_salary',
       'Date Joined_salary', 'UnitName', 'Branch_salary', 'Days Paid', 'Earned Basic','DA', 'HRA',
       'Conveyance', 'Medical Allowance', 'Telephone Reimb',
       'Tel and Int Reimb', 'Bonus', 'Other Allowance', 'Fuel Reimb',
       'Prof Dev Reimb', 'Corp Attire Reimb', 'Meal Allowance',
       'Special Allowance', 'Personal Allowance','Overtime', 'CCA', 'Other Reimb',
       'Arrears', 'Other Earning', 'Variable Pay', 'Leave Encashment',
       'Stipend', 'Consultancy Fees', 'OtherAllowance1', 'OtherAllowance2', 'OtherAllowance3', 'OtherAllowance4', 'OtherAllowance5'
       'Total Earning', 'Insurance', 'CSR',
       'PF', 'ESIC','VPF', 'P.Tax', 'LWF EE', 'Salary Advance', 'Loan Deduction',
       'Loan Interest', 'Fine',	'Damage or Loss','Other Deduction', 'TDS', 'OtherDeduction1', 'OtherDeduction2', 'OtherDeduction3', 'OtherDeduction4', 'OtherDeduction5'
       'Total Deductions','Net Paid', 'BankName', 'Bank A/c Number_salary', 'Account Code_salary', 'Remarks_salary',
       'PF Number (Old)', 'UAN Number_salary', 'ESIC Number_salary', 'Personal A/c Number',
       'E-Mail_salary', 'Mobile No.', 'FIXED MONTHLY GROSS', 'CHECK CTC Gross','Date of payment',	'Arrears salary', 'Cheque No - NEFT date']

    atten_df_columns = ['Emp Code', 'Employee Name', 'Branch', 'Designation', 'Sat\r\n01/02',
       'Sun\r\n02/02', 'Mon\r\n03/02', 'Tue\r\n04/02', 'Wed\r\n05/02',
       'Thu\r\n06/02', 'Fri\r\n07/02', 'Sat\r\n08/02', 'Sun\r\n09/02',
       'Mon\r\n10/02', 'Tue\r\n11/02', 'Wed\r\n12/02', 'Thu\r\n13/02',
       'Fri\r\n14/02', 'Sat\r\n15/02', 'Sun\r\n16/02', 'Mon\r\n17/02',
       'Tue\r\n18/02', 'Wed\r\n19/02', 'Thu\r\n20/02', 'Fri\r\n21/02',
       'Sat\r\n22/02', 'Sun\r\n23/02', 'Mon\r\n24/02', 'Tue\r\n25/02',
       'Wed\r\n26/02', 'Thu\r\n27/02', 'Fri\r\n28/02', 'Sat\r\n29/02',
       'Total\r\nDP', 'Total\r\nABS', 'Total\r\nLWP', 'Total\r\nCL',
       'Total\r\nSL', 'Total\r\nPL', 'Total\r\nL1', 'Total\r\nL2',
       'Total\r\nL3', 'Total\r\nL4', 'Total\r\nL5', 'Total\r\nCO-',
       'Total\r\nCO+', 'Total\r\nOL', 'Total\r\nWO', 'Total\r\nPH',
       'Total\r\nEO', 'Total\r\nWOP', 'Total\r\nPHP', 'Total\r\nOT Hrs',
       'Total\r\nLT Hrs']

    leave_df_columns = ['Emp. Code', 'Emp. Name', 'Leave Type', 'Opening', 'Monthly Increment',
       'Used', 'Closing', 'Leave Accrued', 'Encash']

    leftemp_df_columns = ['Employee Name', 'Employee Code_left', 'Date Joined', 'Date Left',
       'UAN Number']

    unit_df_columns = ['Unit', 'Location Code','Location', 'Address', 'Registration_no','Unit_PAN','Unit_LIN','Unit_email','Unit_mobile', 'PE_or_contract',
       'State_or_Central', 'start_time', 'end_time', 'rest_interval','Contractor_name','Contractor_Address','Contractor_PAN', 
        'Contractor_LIN', 'Contractor_email',	'Contractor_mobile','Normal hrs', 'overtime rate']

    logging.info('column variables set')

    

    
    file_list = os.listdir(inputfolder)
    logging.info('input folder is '+str(inputfolder))
    for f in file_list:
        if f[0:6].upper()=='MASTER':
            masterfilename = f
            logging.info('masterfilename is :'+f)
        if f[0:6].upper()=='SALARY':
            salaryfilename = f
            logging.info('salaryfilename is :'+f)
        if f[0:10].upper()=='ATTENDANCE':
            attendancefilename = f
            logging.info('attendancefilename is :'+f)
        if f[0:5].upper()=='LEAVE':
            leavefilename = f
            logging.info('leavefilename is :'+f)
        if f[0:14].upper()=='LEFT EMPLOYEES':
            leftempfilename = f
            logging.info('leftempfilename is :'+f)
        if f[0:5].upper()=='UNITS':
            unitfilename = f
            logging.info('unitfilename is :'+f)
    
    logging.info('file names set')
    try:
        if 'masterfilename' in locals():
            masterfile = os.path.join(inputfolder,masterfilename)
            employee_data = pd.read_excel(masterfile)
            employee_data.dropna(subset=['Employee Code','Location Code'], inplace=True)
            employee_data.dropna(how='all', inplace=True)
            employee_data.reset_index(drop=True, inplace=True)
            employee_data.rename(columns={"Employee Code": "Employee Code_master", "Employee Name": "Employee Name_master", "Designation": "Designation_master", "Branch": "Branch_master", "Date Joined": "Date Joined_master", "UAN Number": "UAN Number_master",
                            "ESIC Number": "ESIC Number_master", "Bank A/c Number": "Bank A/c Number_master", "Account Code": "Account Code_master",
                            "E-Mail": "E-Mail_master", "Remarks": "Remarks_master"}, inplace=True)
            logging.info('employee data loaded')
        else:
            employee_data = pd.DataFrame(columns = emp_df_columns)
            logging.error('employee data not available setting empty dataset')
        if 'salaryfilename' in locals():
            salaryfile = os.path.join(inputfolder,salaryfilename)
            salary_data = pd.read_excel(salaryfile)
            salary_data.dropna(subset=['Emp Code'], inplace=True)
            salary_data.dropna(how='all', inplace=True)
            salary_data.reset_index(drop=True, inplace=True)
            salary_data.rename(columns={"Emp Code": "Emp Code_salary", "Emp Name": "Emp Name_salary","DesigName": "Designation_salary", "Branch": "Branch_salary", "Date Joined": "Date Joined_salary", "UAN Number": "UAN Number_salary",
                            "ESIC Number": "ESIC Number_salary", "Bank A/c Number": "Bank A/c Number_salary", "Account Code": "Account Code_salary",
                            "E-Mail": "E-Mail_salary", "Remarks": "Remarks_salary"}, inplace=True)
            logging.info('salary data loaded')
        else:
            salary_data = pd.DataFrame(columns = salary_df_columns)
            logging.info('salary data not available setting empty dataset')
        if 'attendancefilename' in locals():
            attendancefile = os.path.join(inputfolder,attendancefilename)
            attendance_data = pd.read_excel(attendancefile)
            attendance_data.dropna(subset=['Emp Code'], inplace=True)
            attendance_data.dropna(how='all', inplace=True)
            attendance_data.reset_index(drop=True, inplace=True)
            logging.info('attendance data loaded')
        else:
            attendance_data = pd.DataFrame(columns = atten_df_columns)
            logging.info('attendance data not available setting empty dataset')
        if 'leavefilename' in locals():
            leavefile = os.path.join(inputfolder,leavefilename)
            leave_data = pd.read_excel(leavefile)
            leave_data.dropna(subset=['Emp. Code'], inplace=True)
            leave_data.dropna(how='all', inplace=True)
            leave_data.reset_index(drop=True, inplace=True)
            logging.info('leave data loaded')
        else:
            leave_data = pd.DataFrame(columns = leave_df_columns)
            logging.info('leave data not available setting empty dataset')
        if 'leftempfilename' in locals():
            leftempfile = os.path.join(inputfolder,leftempfilename)
            leftemp_data = pd.read_excel(leftempfile)
            leftemp_data.dropna(subset=['Employee Code'], inplace=True)
            leftemp_data.dropna(how='all', inplace=True)
            leftemp_data.reset_index(drop=True, inplace=True)
            leftemp_data.rename(columns={"Employee Code": "Employee Code_left"}, inplace=True)
            logging.info('left employees data loaded')
        else:
            leftemp_data = pd.DataFrame(columns = leftemp_df_columns)
            logging.info('left employees data not available setting empty dataset')
        if 'unitfilename' in locals():
            unitfile = os.path.join(inputfolder,unitfilename)
            unit_data = pd.read_excel(unitfile)
            unit_data.dropna(subset=['Location Code'], inplace=True)
            unit_data.dropna(how='all', inplace=True)
            unit_data.reset_index(drop=True, inplace=True)
            logging.info('unit data loaded')
        else:
            unit_data = pd.DataFrame(columns = unit_df_columns)
            logging.info('unit data not available setting empty dataset')

        employee_data.drop(columns='Date Left', inplace=True)
        
        logging.info(type(employee_data['Location Code'].unique()[0]))
        logging.info(unit_data.head())

        if str(employee_data['Location Code'].dtype)[0:3] != 'int':
            employee_data['Location Code'] = employee_data['Location Code'].astype(int)
        
        if str(employee_data['Employee Code_master'].dtype)[0:3] != 'obj':
            employee_data['Employee Code_master'] = employee_data['Employee Code_master'].astype(str)

        if str(unit_data['Location Code'].dtype)[0:3] != 'int':
            unit_data['Location Code'] = unit_data['Location Code'].astype(int)


        employee_data.drop(columns=list(employee_data.columns.intersection(salary_data.columns)), inplace=True)

        if str(salary_data['Emp Code_salary'].dtype)[0:3] != 'obj':
            salary_data['Emp Code_salary'] = salary_data['Emp Code_salary'].astype(str)

        attendance_data.drop(columns=['Employee Name', 'Branch', 'Designation'], inplace=True)

        if str(attendance_data['Emp Code'].dtype)[0:3] != 'obj':
            attendance_data['Emp Code'] = attendance_data['Emp Code'].astype(str)

        if str(leave_data['Emp. Code'].dtype)[0:3] != 'obj':
            leave_data['Emp. Code'] = leave_data['Emp. Code'].astype(str)

        leftemp_data.drop(columns=['Employee Name', 'Date Joined', 'UAN Number'],inplace=True)

        if str(leftemp_data['Employee Code_left'].dtype)[0:3] != 'obj':
            leftemp_data['Employee Code_left'] = leftemp_data['Employee Code_left'].astype(str)

        CDE_Data = salary_data.merge(employee_data,how='left',left_on='Emp Code_salary', right_on='Employee Code_master').merge(
            unit_data,how='left',on='Location Code').merge(
                attendance_data,how='left',left_on='Emp Code_salary', right_on='Emp Code').merge(
                    leave_data, how='left', left_on='Emp Code_salary', right_on='Emp. Code').merge(
                        leftemp_data, how='left', left_on='Emp Code_salary', right_on='Employee Code_left')
        
        '''
        CDE_Data = employee_data.merge(unit_data,how='left',on='Location Code').merge(
            salary_data,how='left',left_on='Employee Code',right_on='Emp Code').merge(
                attendance_data,how='left',left_on='Employee Code', right_on='Emp Code').merge(
                    leave_data, how='left', left_on='Employee Code', right_on='Emp. Code').merge(
                        leftemp_data, how='left', on='Employee Code')
        '''
        

        CDE_Data['Employee Code'] = CDE_Data['Emp Code_salary']
        CDE_Data['Employee Name'] = CDE_Data['Emp Name_salary'].combine_first(CDE_Data['Employee Name_master'])
        CDE_Data['Designation'] = CDE_Data['Designation_salary'].combine_first(CDE_Data['Designation_master'])
        CDE_Data['Branch'] = CDE_Data['Branch_salary'].combine_first(CDE_Data['Branch_master'])
        CDE_Data['Date Joined'] = CDE_Data['Date Joined_salary'].combine_first(CDE_Data['Date Joined_master'])
        CDE_Data['UAN Number'] = CDE_Data['UAN Number_salary'].combine_first(CDE_Data['UAN Number_master'])
        CDE_Data['ESIC Number'] = CDE_Data['ESIC Number_salary'].combine_first(CDE_Data['ESIC Number_master'])
        CDE_Data['Bank A/c Number'] = CDE_Data['Bank A/c Number_salary'].combine_first(CDE_Data['Bank A/c Number_master'])
        CDE_Data['Account Code'] = CDE_Data['Account Code_salary'].combine_first(CDE_Data['Account Code_master'])
        CDE_Data['E-Mail'] = CDE_Data['E-Mail_salary'].combine_first(CDE_Data['E-Mail_master'])
        CDE_Data['Remarks'] = CDE_Data['Remarks_salary'].combine_first(CDE_Data['Remarks_master'])

        
        logging.info('merged all data sets')

        logging.info(len(salary_data))
        logging.info(len(CDE_Data))



        rename_list=[]
        renamed=[]
        drop_list=[]
        for x in list(CDE_Data.columns):
            if x[-2:]=='_x':
                rename_list.append(x)
                renamed.append(x[0:-2])
            if x[-2:]=='_y':
                drop_list.append(x)
        
        rename_dict = dict(zip(rename_list,renamed))

        CDE_Data.rename(columns=rename_dict, inplace=True)

        logging.info('columns renamed correctly')

        CDE_Data.drop(columns=drop_list, inplace=True)

        logging.info('dropped duplicate columns')

        print(CDE_Data['Date of payment'].dtype)

        if str(CDE_Data['Date of payment'].dtype)[0:8] == 'datetime':
            CDE_Data['Date of payment'] = CDE_Data['Date of payment'].dt.date
            
    except KeyError as e:
        logging.info("Key error : Check if {} column exsists".format(e))
        print("Key error {}".format(e))
        output_text="Failed: Check input file format  \n column {} not found".format(e)
        return

    monthyear = month+' '+str(year)
    print(monthyear)
    print(masterfilename.upper())
    if monthyear.upper() in masterfilename.upper():
        progress['maximum']=calculate_num_loop(CDE_Data)
        logging.info('month year matches with data')
        #for all state employees(PE+contractor)
        statedata = CDE_Data.loc[CDE_Data['State_or_Central']=='State'].copy(deep=True)
        
        statedata.State=statedata.State.str.lower()
        CDE_States = list(statedata['State'].unique())
        implemented_state_list=[x.lower() for x in State_Process.keys()]
        if Testing==True:
            print("In testing Mode")
            CDE_States=implemented_state_list
        
        for state in CDE_States:
            report.configure(text="Creating state forms for {}".format(state))
            master.update()
            print("-----------------------------")
            state=state.lower()
            print(state)
            if Testing==True:
                statedata.State=state
            # continue
            if state not in implemented_state_list:
                logging.info('State {} not implemented in our set,that is {} hence continuing'.format(state,implemented_state_list))
                print('State {} not implemented in our set,that is {} hence continuing'.format(state,implemented_state_list))
                continue
            
            unit_with_location = list((statedata[statedata.State==state]['Unit']+';'+statedata[statedata.State==state]['Location']).unique())
            
            for UL in unit_with_location:
                inputdata = statedata[(statedata['State']==state) & (statedata['Unit']==UL.split(';')[0]) & (statedata['Location']==UL.split(';')[1])].copy(deep=True)
                inputdata['Contractor_name'] = inputdata['Contractor_name'].fillna(value='')
                inputdata['Contractor_Address'] = inputdata['Contractor_Address'].fillna(value='')
                inputdata.fillna(value=0, inplace=True)
                if UL.strip()[-1] == '.':
                    ULis = UL.strip()[0:-1]
                else:
                    ULis = UL.strip()
                inpath = os.path.join(inputfolder,Register_folder,'States',state,ULis)
                logging.info('folder for forms path is'+str(inpath))
                if os.path.exists(inpath):
                    logging.info('running state process')
                    logging.info(inputdata)
                    contractor_name= inputdata['Contractor_name'].unique()[0]
                    contractor_address= inputdata['Contractor_Address'].unique()[0]
                    State_Process[state](data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year,report=report,master=master)
                else:
                    logging.info('making directory')
                    os.makedirs(inpath)
                    logging.info('directory created')
                    logging.info(inputdata)
                    contractor_name= inputdata['Contractor_name'].unique()[0]
                    contractor_address= inputdata['Contractor_Address'].unique()[0]
                    State_Process[state](data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year,report=report,master=master)
                progress["value"]+=1
                percent.configure(text=str(progress["value"]*100//progress["maximum"])+"%")
                progress.update()
                master.update()
        #for contractors form
        contractdata = CDE_Data.loc[(CDE_Data['State_or_Central']=='State') & (CDE_Data['PE_or_contract']=='Contract')].copy(deep=True)
        contractor_units = list((contractdata['Unit']+';'+contractdata['Location']).unique())
        report.configure(text="Creating contractor Forms")
        master.update()
        
        for UL in contractor_units:
            inputdata = contractdata.loc[(contractdata['Unit']==UL.split(';')[0]) & (contractdata['Location']==UL.split(';')[1])].copy(deep=True)
            contractor_name= inputdata['Contractor_name'].unique()[0]
            contractor_address= inputdata['Contractor_Address'].unique()[0]
            inputdata.fillna(value=0, inplace=True)
                    
            if UL.strip()[-1] == '.':
                ULis = UL.strip()[0:-1]
            else:
                ULis = UL.strip()
            inpath = os.path.join(inputfolder,Register_folder,'Contractors',ULis)
            if os.path.exists(inpath):
                logging.info('running contractor process')
            else:
                logging.info('making directory')
                os.makedirs(inpath)
                logging.info('directory created')
            
            if not inputdata.empty:
                Contractor_Process(data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year,report=report,master=master)
            progress["value"]+=1
            percent.configure(text=str(progress["value"]*100//progress["maximum"])+"%")
            progress.update()
            master.update()
            
        #for central form
        centraldata = CDE_Data.loc[CDE_Data['State_or_Central']=='Central'].copy(deep=True)
        central_units = list((centraldata['Unit']+','+centraldata['Location']).unique())
        report.configure(text="Creating central Forms")
        master.update()
        
        for UL in central_units:
            inputdata = centraldata.loc[(centraldata['Unit']==UL.split(',')[0]) & (centraldata['Location']==UL.split(',')[1])].copy(deep=True)
            contractor_name= inputdata['Contractor_name'].unique()[0]
            contractor_address= inputdata['Contractor_Address'].unique()[0]
            inputdata.fillna(value=0, inplace=True)
                    
            inpath = os.path.join(inputfolder,Register_folder,'Central',UL)
            if os.path.exists(inpath):
                logging.info('running contractor process')
            else:
                logging.info('making directory')
                os.makedirs(inpath)
                logging.info('directory created')
            if not inputdata.empty:
                Central_Process(data=inputdata,contractor_name=contractor_name,contractor_address=contractor_address,filelocation=inpath,month=month,year=year,report=report,master=master)    
            progress["value"]+=1
            percent.configure(text=str(progress["value"]*100//progress["maximum"])+"%")
            progress.update()
            master.update()
    else:
        output_text = "Date you mentioned doesn't match with Input data"
        logging.error(output_text)

    
def calculate_num_loop(CDE_Data):
    count=0
    statedata = CDE_Data[CDE_Data['State_or_Central']=='State'].copy()
    statedata.State=statedata.State.str.lower()
    CDE_States = list(statedata['State'].unique())
    implemented_state_list=[x.lower() for x in State_Process.keys()]
    if Testing==True:
        CDE_States=implemented_state_list
    for state in CDE_States:
        state=state.lower()
        if Testing==True:
            statedata.State=state
        if state not in implemented_state_list:
            continue
        count +=len(list((statedata[statedata.State==state]['Unit']+';'+statedata[statedata.State==state]['Location']).unique()))
    
    contractdata = CDE_Data[(CDE_Data['State_or_Central']=='State') & (CDE_Data['PE_or_contract']=='Contract')].copy()
    count += len(list((contractdata['Unit']+';'+contractdata['Location']).unique()))
    
    centraldata = CDE_Data[CDE_Data['State_or_Central']=='Central'].copy()
    count += len(list((centraldata['Unit']+','+centraldata['Location']).unique()))
    return count



DataProcess = {'Type1':Type1,'Type2':Type2,'Type3':Type3,'Type4':Type4,'Type5':Type5}


def CompanyDataProcessing(companytype,inputfolder,month,year):
    inputfolder = Path(inputfolder)
    yr = int(year)
    DataProcess[companytype](inputfolder,month,yr)

#backend code ends here

Types = ['Type1','Type2','Type3','Type4','Type5']


Months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']

Years = ['2017','2018','2019','2020','2021']

Typeis = tk.StringVar()

companyname = tk.StringVar()

month = tk.StringVar()
year = tk.StringVar()


folderLabel = ttk.LabelFrame(master, text="Select the Company")
folderLabel.grid(column=0,row=1,padx=20,pady=20)

TypeLabel = Label(folderLabel,text="Company Type")
TypeLabel.grid(column=1,row=0,padx=20,pady=20)

TypeEntry = ttk.Combobox(folderLabel,values=Types,textvariable=Typeis)
TypeEntry.grid(column=2, row=0, padx=20,pady=20)

companynameLabel = Label(folderLabel, text="Company Name")
companynameLabel.grid(column=1, row=1, padx=20,pady=20)

comapnynameEntry = tk.Entry(folderLabel,textvariable=companyname)
comapnynameEntry.grid(column=2, row=1, padx=20,pady=20)

MonthLabel = Label(folderLabel, text="Month and Year")
MonthLabel.grid(column=1, row=3, padx=20,pady=20)

MonthEntry = ttk.Combobox(folderLabel,values=Months,textvariable=month)
MonthEntry.grid(column=2, row=3, padx=20,pady=20)

YearEntry = ttk.Combobox(folderLabel,values=Years,textvariable=year)
YearEntry.grid(column=3, row=3, padx=20,pady=20)

def disfo():
    foldername = filedialog.askdirectory()
    logging.info(foldername)
    logging.info(type(foldername))
    foldernamelabel.configure(text=foldername)


button = ttk.Button(folderLabel, text = "Select Company Folder", command=disfo)
button.grid(column=1, row=2, columnspan=2,padx=20, pady=20)

foldernamelabel = Label(folderLabel, text="")
foldernamelabel.grid(column=1, row=4, columnspan=2,padx=20,pady=20)




def generateforms(comptype,mn,yr):
    companytype=comptype.get()

    month = mn.get()
    year = yr.get()


    getfolder = foldernamelabel.cget("text")



    logging.info(type(companytype))
    logging.info(companytype)

    logging.info(type(getfolder))
    logging.info(getfolder)

    if (companytype =="" and getfolder =="" and (month =="" or year =="")):
        report.configure(text="Please select month year, company folder and company type")
    elif (companytype=="" and getfolder =="" and not(month =="" or year =="")):
        report.configure(text="Please select company folder and company type")
    elif (companytype=="" and getfolder !="" and not(month =="" or year =="")):
        report.configure(text="Please select company type")
    elif (companytype!="" and getfolder =="" and not(month =="" or year =="")):
        report.configure(text="Please select company folder")
    elif (companytype =="" and getfolder !="" and (month =="" or year=="")):
        report.configure(text="Please select month year and company type")
    elif (companytype!="" and getfolder=="" and (month =="" or year=="")):
        report.configure(text="Please select month year and company folder")
    elif (companytype!="" and getfolder!="" and (month =="" or year=="")):
        report.configure(text="Please select month year")
    else:
        logging.info("{} , {} , {} , {}".format(companytype, getfolder,  month,  year))
        report.configure(text="Processing")
        try:
            CompanyDataProcessing(companytype,getfolder,month,year)
        except Exception as e:
            logging.info('Failed')
            report.configure('Failed')
        else:
            if output_text=='':
                logging.info('Completed Form Creation')
                report.configure(text='Completed Form Creation')
            else:
                logging.info(output_text)
                report.configure(text=output_text)
        finally:
            logging.info('done')
        
def convert_forms_to_pdf():

    getfolder = foldernamelabel.cget("text")

    if getfolder=="":
        report.configure(text="Please select company folder")
    else:
        registerfolder = os.path.join(Path(getfolder),Register_folder)
        if os.path.exists(registerfolder):
            for root, dirs, files in os.walk(registerfolder):
                for fileis in files:
                    if fileis.endswith(".xlsx"):
                        try:
                            create_pdf(root,fileis)
                        except Exception as e:
                            logging.info('Failed pdf Conversion')
                            report.configure(text="Failed")
                        else:
                            logging.info('Completed pdf Conversion')
                            report.configure(text="Completed")
                        finally:
                            logging.info('done')
        else:
            report.configure(text="Registers not available")
                        



generateforms = partial(generateforms,Typeis,month,year)

button = ttk.Button(master, text = "Generate Forms", command=generateforms)
button.grid(column=1, row=1, columnspan=1,padx=20, pady=20)

Detailbox = ttk.LabelFrame(master, text="")
Detailbox.grid(column=0,row=2,padx=20,pady=20)

report = Label(Detailbox, text="                                                            ")
report.grid(column=0, row=0, padx=20,pady=20)

if Testing==True:
    report.configure(text="In testing Mode,will override state variable.\n Forms will be created for all implemented states")
    master.update()
if not os.path.isdir(State_forms):
    report.configure(text="Directory  {} not found, Need to have Form formats in the folder".format(State_forms))
    master.update()

button2 = ttk.Button(master, text = "Convert forms to PDF", command=convert_forms_to_pdf)
button2.grid(column=0, row=3, columnspan=2,padx=20, pady=20)

progress = ttk.Progressbar(master, orient = HORIZONTAL, 
              length = 200, mode = 'determinate') 


progress.grid(column=2, row=1, padx=20,pady=2)
progress["value"]=0

percent = Label(master, text="0 %")
percent.grid(column=2, row=1, padx=10,pady=10)

mainloop()


