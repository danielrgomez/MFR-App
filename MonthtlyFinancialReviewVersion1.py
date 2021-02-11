# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 16:24:50 2020

@author: dgomezpe
"""

from tkinter import *
import os
from tempfile import NamedTemporaryFile
import shutil
import csv
import pandas as pd
import os
import glob
import win32com.client as win32
from win32com.client.gencache import EnsureDispatch
import re
import numpy as np
import xlsxwriter
from send2trash import send2trash
from tkinter import ttk
import time
import tkinter.messagebox
import sys


#Variables
global query_file_path
global es_reports_file_path

# #Query Variables
# global payroll_query
# global kk_enc_query
# global kk_exp_crefn_query
# global kk_exp_uflor_query
# global budget_query
# global ledger_query

#Cost Center Variables
global department
global fund
global sof
global period
global account

#Tkinter My Progress Bar
#global my_progress
#global second_root
    
    
# #Set variables
# query_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\2021_PER_4_Queries_UFLOR.xlsx'


# #Query Sheets Set as variables
# payroll_query = pd.read_excel(query_file_path,'PAYROLL',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
# kk_enc_query = pd.read_excel(query_file_path,'KK_ENC',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
# kk_exp_crefn_query = pd.read_excel(query_file_path,'KK_EXP_CREFN',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
# kk_exp_uflor_query = pd.read_excel(query_file_path,'KK_EXP_UFLOR',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
# budget_query = pd.read_excel(query_file_path,'BUDGET',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
# ledger_query = pd.read_excel(query_file_path,'LEDGER',converters={'ACCOUNT':str,'DEPTID':str,'ACCOUNTING_PERIOD':str,'FUND_CODE':str,'DEPTID.1':str})



def main():
    print('MFR has been imported')    
   
def setting_query_file_path(file_path_for_query,my_progress,second_root,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    #Progress Bar Variables Set as Global
    #global my_progress
    #global second_root
    #global my_label
    global validation_to_assign_query
    query_file_path = file_path_for_query
    
    #Query Sheets Set as variables
    payroll_query = pd.read_excel(query_file_path,'PAYROLL',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,1,second_root,my_label,"Extracted Query Data: Payroll")
    kk_enc_query = pd.read_excel(query_file_path,'KK_ENC',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,1,second_root,my_label,"Extracted Query Data: KK_ENC")
    kk_exp_crefn_query = pd.read_excel(query_file_path,'KK_EXP_CREFN',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,2,second_root,my_label,"Extracted Query Data: KK_EXP_CREFN")
    kk_exp_uflor_query = pd.read_excel(query_file_path,'KK_EXP_UFLOR',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,2,second_root,my_label,"Extracted Query Data: KK_EXP_UFLOR")
    budget_query = pd.read_excel(query_file_path,'BUDGET',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,2,second_root,my_label,"Extracted Query Data: Budget")
    ledger_query = pd.read_excel(query_file_path,'LEDGER',converters={'ACCOUNT':str,'DEPTID':str,'ACCOUNTING_PERIOD':str,'FUND_CODE':str,'DEPTID.1':str})
    step_progress_bar(my_progress,1,second_root,my_label,"Extracted Query Data: Ledger")
    print("Query Data has been assigned because the validation is 1")
    
   
def step_progress_bar(my_progress,increase_value,second_root,my_label,current_status):
    my_progress['value'] += increase_value
    current_status_as_text = current_status
    running_as_text = "Running: " + str(int(my_progress['value'])) + "%"
    running_as_text = "{:>10}".format(running_as_text)
    current_status_as_text = "{:<80}".format(current_status_as_text)
    #my_label = Label(second_root,text = running_as_text)
    my_label.config(text = current_status_as_text + running_as_text)
    my_label.pack(pady = 1)
    second_root.update_idletasks()
    second_root.update()
    own_working_directory = os.path.abspath(os.getcwd())
    second_root.iconbitmap(own_working_directory + "\logo.ico")
    time.sleep(.2)
    #second_root.mainloop()
    if increase_value == 100:
        tkinter.messagebox.showinfo("Completed MFR Report","Completed MFR")
        
    
def script_control(report,department,fund,sof,period,project,flex,filepath,file_path_for_query,file_path_for_saving,validation,second_root,my_progress,my_label):
    #global my_progress
    #global second_root
    #global my_label
    global excel_sheet_open_error
    # second_root = Tk()
    # second_root.title("Progress To Complete MFR Report")
    # second_root.geometry("420x100")
    # my_progress = ttk.Progressbar(second_root, orient = HORIZONTAL, length = 400, mode = 'determinate')
    # my_progress.pack(pady = 20)
    # running_as_text = "Running: "
    # running_as_text = "{:>100}".format(running_as_text)
    # my_label = Label(second_root,text = running_as_text)
    # my_label.pack(pady = 1)
    # #my_label.config(text = my_progress['value'])
    # step_progress_bar(my_progress,1,second_root,my_label,"Script control initiated for report: " + report)
      
    step_progress_bar(my_progress,1,second_root,my_label,"Script control initiated for report: " + report)
    
    try:
        if validation == 1:
            setting_query_file_path(file_path_for_query,my_progress,second_root,my_label)
            
                    
        if report == "DEPT_APPROP":
            #print("THIS IS THE ONE WE ARE LOOKING AT: " + str(filepath))
            dept_approp_uflor(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "DEPT_CASH":
            dept_cash_uflor(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "DEPT_STUGOV":
            print(report + department + fund + sof + period + project + flex + filepath)
            dept_stugov(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "FUND_HOUSE":
            fund_house(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "FUND_CASH_DSO":
            fund_cash_dso(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "FUND_CASH":
            fund_cash_uflor(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "RESIDUAL":
            residual(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        elif report == "FLEX_FUND":
            flex_fund(department,fund,sof,period,project,flex,filepath,file_path_for_saving,second_root,my_progress,my_label)
        
        if excel_sheet_open_error == "ExcelSheetOpenError":
            return 'ExcelSheetOpenError'
        else: return 'NoError'
    
    except FileNotFoundError as error_as_string:
        #second_root.quit()
        print("Error as String: " + str(error_as_string))
        tkinter.messagebox.showinfo("Query or ES File or Path Not Found","Please make sure the Query or ES file and/or directory path are correct \n\n" + str(error_as_string).replace("[Errno 2] ","").replace("[WinError 3] ",""))
        #print("FileNotFoundError")
        return 'FileNotFoundError'
            
    
#1
def dept_approp_uflor(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_APPROP"
    excel_file_open_validation = False  
    
    
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    print(es_reports_file_path)
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1

    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
        
    
    if excel_file_open_validation == False:
        
               
        print("All are saved!")
        
        #Report Sheets - Appropriations uses a different header number
        #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_APPROP\Temporary xlsx'
        temporary_file_path = newpath
        #Progress Bar
        #step_progress_bar(my_progress,10,second_root,my_label)
        
        appropriations_report = pd.read_excel(temporary_file_path+'\Appropriations_Summary_Excel.xlsx','Sheet1',header=1,converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        budget_transaction_report = pd.read_excel(temporary_file_path+'\Budget_Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Budget_Transaction_Detail_Excel.xlsx")
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Open_Encumbrance_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
        #Progress Bar
        
    
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_open_encum_report,
                                      unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        #print("Total in Length: " + str(len(all_accounts)))
        print("All unique accounts accross all raw data: \n" + str(len(all_unique_accounts)))
        print(all_unique_accounts)
        
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        # print("This is the Department: " + str(department))
        # print("This is the Fund: " + str(fund))
        # print("This is the Period: " + str(period))
        # print("*********************asdfasdffasdffasdf****************************************************")
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['DEPTID'] == department) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['DEPTID'] == department) &
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
               
            payroll_totals_mtd_yes_1 = payroll_query.loc[
                    (payroll_query['ACCOUNT'] == acct) &
                    (payroll_query['DEPTID'] == department) &
                    (payroll_query['FUND_CODE'] == fund),
                    'MONETARY_AMOUNT'].sum()
                          
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_expense_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = appropriations_report.loc[
                            (appropriations_report['DeptID'] == department) &
                            (appropriations_report['Account Code'] == acct) &
                            (appropriations_report['Fund Code'] == fund),
                             'YTD Expenses'].sum()
            mtd_summary_report = appropriations_report.loc[
                            (appropriations_report['DeptID'] == department) &
                            (appropriations_report['Account Code'] == acct) &
                            (appropriations_report['Fund Code'] == fund),
                             'MTD Expenses'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['DeptID'] == department) &
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['DeptID'] == department) &
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Fund Code'] == fund),
                             'Open Amount'].sum()
            payroll_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund),
                             'Salary'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_expense_report[acct] = ytd_summary_expense_report
            each_mtd_summary_report[acct] = mtd_summary_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = payroll_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_expense_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
            
            #Create new Python MFR Results Folder
            #newpath = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results'
            newpath = str(""r""+file_path_for_saving)
            if not os.path.exists(newpath):
                os.makedirs(newpath)
                
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"ES Report Data to various dictionaries")
                
        #Create new workbook and add results onto it
        workbook = xlsxwriter.Workbook(newpath + '\DEPT_APPROP_MFR.xlsx')
        worksheet = workbook.add_worksheet("DEPT_APPROP_RECON")
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        


        for acct in all_unique_accounts:
            print(str(acct) + ":" + str(each_ytd_summary_expense_report[acct]))
            print(str(acct) + ":" + str(each_mtd_summary_report[acct]))
            print(str(acct) + ":" + str(each_tran_detail_report[acct]))
            print(str(acct) + ":" + str(each_kk_totals_ytd_report[acct]))
            print(str(acct) + ":" + str(each_gl_totals_ytd_report[acct]))
            print(str(acct) + ":" + str(each_open_encumbrance_report[acct]))
            print(str(acct) + ":" + str(each_payroll_totals_report[acct]))
            print(str(acct) + ":" + str(each_ytd_summary[acct]))
            print(str(acct) + ":" + str(each_mtd_summary[acct]))
            print(str(acct) + ":" + str(each_glkk_totals_ytd[acct]))
            print(str(acct) + ":" + str(each_open_enc[acct]))
            print(str(acct) + ":" + str(each_payroll_totals[acct]))
            
            if (each_ytd_summary_expense_report[acct] + each_mtd_summary_report[acct] + each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + each_gl_totals_ytd_report[acct] + each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] + each_ytd_summary[acct] + each_mtd_summary[acct] + each_glkk_totals_ytd[acct] +
                each_open_enc[acct] + each_payroll_totals[acct] != 0):
                    
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_expense_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                    worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
            
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
        
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        
        worksheet.set_column('A:Z', 22)
        
        
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        
        workbook.close()
    
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        
        
        # worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        # worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        # worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        # worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        
        
        
        
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")
  

#2   
def dept_cash_uflor(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_CASH"
    excel_file_open_validation = False  
    
    
    
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
   
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    
    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
        
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
        
    print("All are saved!")
    
    if excel_file_open_validation == False:
         
        #Report Sheets - Appropriations uses a different header number
        #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_CASH\Temporary xlsx'
        temporary_file_path = newpath
        
        appropriations_report = pd.read_excel(temporary_file_path+'\Appropriations_Summary_Excel.xlsx','Sheet1',header=1,converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        budget_transaction_report = pd.read_excel(temporary_file_path+'\Budget_Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Budget_Transaction_Detail_Excel.xlsx")
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Open_Encumbrance_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
    
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_open_encum_report,
                                      unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        #print("Total in Length: " + str(len(all_accounts)))
        print("All unique accounts accross all raw data: \n" + str(len(all_unique_accounts)))
        print(all_unique_accounts)
    
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
    
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['OPERATING_UNIT'] == sof) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['OPERATING_UNIT'] == sof) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['DEPTID'] == department) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['DEPTID'] == department) &
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
    
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund) &
                            (cash_summary_report['Source of Funds Code'] == sof),
                             'YTD Expense'].sum()
            ytd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund) &
                            (cash_summary_report['Source of Funds Code'] == sof),
                             'YTD Revenue'].sum()
            mtd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund)&
                            (cash_summary_report['Source of Funds Code'] == sof),
                             'MTD Expense'].sum()
            mtd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund)&
                            (cash_summary_report['Source of Funds Code'] == sof),
                             'MTD Revenue'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['DeptID'] == department) &
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund)&
                            (transaction_detail_report['Source of Funds Code'] == sof),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund)&
                            (kk_to_gl_summary_report['Source of Funds Code'] == sof),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund)&
                            (kk_to_gl_summary_report['Source of Funds Code'] == sof),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['DeptID'] == department) &
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Fund Code'] == fund)&
                            (open_encum_report['Source of Funds Code'] == sof),
                             'Open Amount'].sum()
            salary_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Source of Funds Code'] == sof),
                             'Salary'].sum()
            fringe_pool_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Source of Funds Code'] == sof),
                             'Fringe Pool Amount'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_report[acct] = ytd_summary_expense_report +  ytd_summary_revenue_report
            each_mtd_summary_report[acct] = mtd_summary_expense_report + mtd_summary_revenue_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = salary_totals_report + fringe_pool_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
            
         
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        #Create new workbook and add results onto it
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\DEPT_CASH_MFR.xlsx")
        worksheet = workbook.add_worksheet("DEPT_CASH_RECON")
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            if (each_ytd_summary_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] + 
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] + 
                each_payroll_totals[acct] != 0):
            
            
                row += 1
                #Write Accounts
                worksheet.write(row, 2,acct,border)
                
                #Write Report Results
                worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
                worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                
             
                #Write Query Results
                worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                worksheet.write(row, 14,each_open_enc[acct],number_format)
                worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                
                #Variance Results
                worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
            
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format_bold)
        
        worksheet.set_column('A:Z', 22)
        
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        
        workbook.close()
        
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")  
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")
        
    
#3
def dept_stugov(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_STUGOV"
    excel_file_open_validation = False 
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
            
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    
    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
    
    #Report Sheets - Appropriations uses a different header number
    #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_STUGOV\Temporary xlsx'
    if excel_file_open_validation == False:
        temporary_file_path = newpath
    
    
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        stugov_report = pd.read_excel(temporary_file_path+'\StuGov_Summary_Excel.xlsx','Sheet1',header=1,converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted StuGov_Summary_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
        
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        unique_stugov_report = stugov_report['Account Code'].dropna()
        len(unique_stugov_report)
        
        
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_payroll_recon_report,
                                      unique_stugov_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
        print("All unique accounts accross all raw data: " + str(len(all_unique_accounts)))
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['DEPTID'] == department) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['DEPTID'] == department) &
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_ytd_summary.get(95451000))
        
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        
        #2) Report Data into dictionaries format is -> account:amount
        
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_expense_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = stugov_report.loc[
                            (stugov_report['DeptID'] == department) &
                            (stugov_report['Account Code'] == acct) &
                            (stugov_report['Fund Code'] == fund),
                             'YTD Expenses'].sum()
            mtd_summary_report = stugov_report.loc[
                            (stugov_report['DeptID'] == department) &
                            (stugov_report['Account Code'] == acct) &
                            (stugov_report['Fund Code'] == fund),
                             'MTD Expenses'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['DeptID'] == department) &
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            #No Open Encumbrance Report In This Process So all of the open encumbrance lines will be equal to zero
            open_encumbrance_report = 0.0
            
            payroll_totals_report_1 = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund),
                             'Salary'].sum()
            payroll_totals_report_2 = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund),
                             'Fringe Pool Amount'].sum()
        
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_expense_report[acct] = ytd_summary_expense_report
            each_mtd_summary_report[acct] = mtd_summary_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = payroll_totals_report_1 + payroll_totals_report_2
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_expense_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
        
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        #Create new workbook and add results onto it
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\DEPT_STUGOV_MFR.xlsx")
        worksheet = workbook.add_worksheet("DEPT_STUGOV_RECON")
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\DEPT_STUGOV_MFR.xlsx')
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            if (each_ytd_summary_expense_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] + 
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] + 
                each_payroll_totals[acct] != 0):
            
            
            
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_expense_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                    worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
            
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
            
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        
        worksheet.set_column('A:Z', 22)
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        
        workbook.close()
    
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")

#4    
def fund_house(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_HOUSE"
    excel_file_open_validation = False
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
    
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)
        
    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
     temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
     print(temporary_directory)
     fname = file_list[i]
     #excel = win32.Dispatch()
     #ORIGINAL vv
     #excel = win32.gencache.EnsureDispatch('Excel.Application')
     try:
         excel = win32.dynamic.Dispatch("Excel.Application")
         wb = excel.Workbooks.Open(fname)
         wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
         wb.Close()                               #FileFormat = 56 is for .xls extension
         excel.Application.Quit()
         excel_sheet_open_error = "Nothing"
     except AttributeError as error_as_string_two:
         #excel.Application.Quit()
         print("Error as String: " + str(error_as_string_two))
         tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
         excel_file_open_validation = True
         print(str(excel_file_open_validation) + " is the Error Validtion")
         excel_sheet_open_error = "ExcelSheetOpenError"
         send2trash(base_dir + "\Temporary xlsx")
         break
         #print("FileNotFoundError")
         #return 'FileNotFoundError'
     i += 1
    
    #Report Sheets - Appropriations uses a different header number
    #dtemporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_HOUSE\Temporary xlsx'
    if excel_file_open_validation == False:
        temporary_file_path = newpath
        
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        housing_summary_report = pd.read_excel(temporary_file_path+'\Housing_Fund_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Housing_Fund_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Open_Encumbrance_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
        
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_housing_summary_report = housing_summary_report['Account Code'].dropna()
        len(unique_housing_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query,
                                       unique_housing_summary_report,
                                       unique_open_encum_report,
                                      unique_kk_to_gl_summary_report,
                                       unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['DEPTID.1'] == department) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['DEPTID.1'] == department) &
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        
        
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_expense_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = housing_summary_report.loc[
                            (housing_summary_report['DeptID'] == department) &
                            (housing_summary_report['Account Code'] == acct) &
                            (housing_summary_report['Fund Code'] == fund),
                             'YTD Expense'].sum()
            mtd_summary_report = housing_summary_report.loc[
                            (housing_summary_report['DeptID'] == department) &
                            (housing_summary_report['Account Code'] == acct) &
                            (housing_summary_report['Fund Code'] == fund),
                             'MTD Expense'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['DeptID'] == department) &
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['DeptID'] == department) &
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Fund Code'] == fund),
                             'Open Amount'].sum()
            payroll_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund),
                             'Salary'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_expense_report[acct] = ytd_summary_expense_report
            each_mtd_summary_report[acct] = mtd_summary_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = payroll_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_expense_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
    
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        #Create new workbook and add results onto it
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\FUND_HOUSE_MFR.xlsx')
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\FUND_HOUSE_MFR.xlsx")
        worksheet = workbook.add_worksheet("DEPT_HOUSE_RECON")
    
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            
            if (each_ytd_summary_expense_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] + 
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] + 
                each_payroll_totals[acct] != 0):
                           
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_expense_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                    worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
            
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
        
        
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format)
        
        worksheet.set_column('A:Z', 22)
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        
        workbook.close()
    
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        #Progress Bar
        step_progress_bar(my_progress,1,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")

#5
def fund_cash_dso(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_CASH_DSO"
    excel_file_open_validation = False
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
            
   #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)  
        
    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
    
    #Report Sheets - Appropriations uses a different header number
    #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_CASH_DSO\Temporary xlsx'
    
    if excel_file_open_validation == False:
        temporary_file_path = newpath
        
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,4,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,5,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,10,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query,
                                       unique_kk_to_gl_summary_report,
                                       unique_cash_summary_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
        
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_crefn_query.loc[
                        (kk_exp_crefn_query['ACCOUNT.1'] == acct) &
                        (kk_exp_crefn_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            
           
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            
            
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]}')
            
        #print(each_payroll_totals.get(655120))
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'YTD Expense'].sum()
            ytd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'YTD Revenue'].sum()
            
            mtd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Expense'].sum()
            mtd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Revenue'].sum()
            
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            
            
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_report[acct] = ytd_summary_expense_report + ytd_summary_revenue_report
            each_mtd_summary_report[acct] = mtd_summary_revenue_report + mtd_summary_expense_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]}')
            
      
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        #Create new workbook and add results onto it
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\FUND_CASH_DSO_MFR.xlsx')
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\FUND_CASH_DSO_MFR.xlsx")
        worksheet = workbook.add_worksheet("FUND_CASH_DSO_RECON")
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:H3', 'Reports Data', merge_format)
        worksheet.merge_range('J3:L3', 'Query Data', merge_format)
        worksheet.merge_range('N3:S3', 'Variance Data', merge_format)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        
        
        #Query Header
        worksheet.write(3, 9,"YTD Summary",bold_larger_font)
        worksheet.write(3, 10,"MTD Summary",bold_larger_font)
        worksheet.write(3, 11,"GL/KK  Totals YTD",bold_larger_font)
        
        
        #Variance Header
        worksheet.write(3, 13,"YTD Summary",bold_larger_font)
        worksheet.write(3, 14,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 15,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 16,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 17,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 18,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            
            if (each_ytd_summary_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] != 0):
                       
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                     
                    #Write Query Results
                    worksheet.write(row, 9,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 10,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 11,each_glkk_totals_ytd[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 13,"=D"+str(row + 1)+"-J"+str(row + 1),number_format)
                    worksheet.write(row, 14,"=E"+str(row + 1)+"-K"+str(row + 1),number_format)
                    worksheet.write(row, 15,"=F"+str(row + 1)+"-K"+str(row + 1),number_format)
                    worksheet.write(row, 16,"=G"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 17,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=H"+str(row + 1)+"-D"+str(row + 1),number_format)
            
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')    
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format)
        
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 10,"=sum(K4:K"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format)
        
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 16,"=sum(Q4:Q"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format)
        
        
        
        worksheet.set_column('A:Z', 22)
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        
        
        workbook.close()
    
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")

#6
def fund_cash_uflor(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_CASH"
    excel_file_open_validation = False
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
    
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
    
        #Report Sheets - Appropriations uses a different header number
    #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FUND_CASH\Temporary xlsx'
    if excel_file_open_validation == False:
        temporary_file_path = newpath
        
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Open_Encumbrance_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
    
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_open_encum_report,
                                      unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund)&
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'YTD Expense'].sum()
            ytd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                            'YTD Revenue'].sum()
            mtd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Expense'].sum()
            mtd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Revenue'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Fund Code'] == fund),
                             'Open Amount'].sum()
            salary_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund),
                             'Salary'].sum()
            fringe_pool_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund),
                             'Fringe Pool Amount'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_report[acct] = ytd_summary_expense_report +  ytd_summary_revenue_report
            each_mtd_summary_report[acct] = mtd_summary_expense_report + mtd_summary_revenue_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = salary_totals_report + fringe_pool_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
            
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        
        #Create new workbook and add results onto it
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\FUND_CASH_UFLOR.xlsx')
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\FUND_CASH_MFR.xlsx")
        worksheet = workbook.add_worksheet("FUND_CASH_UFLOR_RECON")
        
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            if (each_ytd_summary_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] +
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] +
                each_payroll_totals[acct] != 0):
            
            
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=abs(D"+str(row + 1)+")-abs(L"+str(row + 1)+")",number_format)
                    worksheet.write(row, 18,"=abs(E"+str(row + 1)+")-abs(M"+str(row + 1)+")",number_format)
                    worksheet.write(row, 19,"=abs(F"+str(row + 1)+")-abs(M"+str(row + 1)+")",number_format)
                    worksheet.write(row, 20,"=abs(G"+str(row + 1)+")-abs(N"+str(row + 1)+")",number_format)
                    worksheet.write(row, 21,"=abs(G"+str(row + 1)+")-abs(H"+str(row + 1)+")",number_format)
                    worksheet.write(row, 22,"=abs(H"+str(row + 1)+")-abs(N"+str(row + 1)+")",number_format)
                    worksheet.write(row, 23,"=abs(I"+str(row + 1)+")-abs(O"+str(row + 1)+")",number_format)
                    worksheet.write(row, 24,"=abs(J"+str(row + 1)+")-abs(P"+str(row + 1)+")",number_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
            
            
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format_bold)
        
        worksheet.set_column('A:Z', 22)
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        workbook.close()
        
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")

#7
def residual(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\RESIDUAL"
    excel_file_open_validation = False 
    
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)
        
    #Add all new xlsx files to new Temporary Folder
    i = 0
    

    for items in file_list:
        temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
        print(temporary_directory)
        fname = file_list[i]
        #excel = win32.Dispatch()
        #ORIGINAL vv
        #excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            excel = win32.dynamic.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(fname)
            wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            excel_sheet_open_error = "Nothing"
        except AttributeError as error_as_string_two:
            #excel.Application.Quit()
            print("Error as String: " + str(error_as_string_two))
            tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            excel_file_open_validation = True
            print(str(excel_file_open_validation) + " is the Error Validtion")
            excel_sheet_open_error = "ExcelSheetOpenError"
            send2trash(base_dir + "\Temporary xlsx")
            break
            #print("FileNotFoundError")
            #return 'FileNotFoundError'
        i += 1
    
    #Report Sheets - Appropriations uses a different header number
    #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\RESIDUAL\Temporary xlsx'
    if excel_file_open_validation == False:
        temporary_file_path = newpath
        
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Open_Encumbrance_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Payroll_Reconciliation_Detail_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Projected_Payroll_Detail_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Transaction_Detail_Excel.xlsx")
        
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_open_encum_report,
                                      unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
    
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['PROJECT_ID'] == project) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[(ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['PROJECT_ID'] == project) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['DEPTID'] == department) &
                        (kk_exp_uflor_query['PROJECT_ID'] == project) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct) &
                        (kk_exp_uflor_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['DEPTID'] == department) &
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['DEPTID'] == department) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['DEPTID'] == department) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund) &
                            (cash_summary_report['Project Code'] == project),
                             'YTD Expense'].sum()
            ytd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund) &
                            (cash_summary_report['Project Code'] == project),
                             'YTD Revenue'].sum()
            mtd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund)&
                            (cash_summary_report['Project Code'] == project),
                             'MTD Expense'].sum()
            mtd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['DeptID'] == department) &
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Fund Code'] == fund)&
                            (cash_summary_report['Project Code'] == project),
                             'MTD Revenue'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['DeptID'] == department) &
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Fund Code'] == fund)&
                            (transaction_detail_report['Project Code'] == project),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund)&
                            (kk_to_gl_summary_report['Project Code'] == project),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['DeptID'] == department) &
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Fund Code'] == fund)&
                            (kk_to_gl_summary_report['Project Code'] == project),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['DeptID'] == department) &
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Fund Code'] == fund)&
                            (open_encum_report['Project Code'] == project),
                             'Open Amount'].sum()
            salary_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Project Code'] == project),
                             'Salary'].sum()
            fringe_pool_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Department Code'] == department) &
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Project Code'] == project),
                             'Fringe Pool Amount'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_report[acct] = ytd_summary_expense_report +  ytd_summary_revenue_report
            each_mtd_summary_report[acct] = mtd_summary_expense_report + mtd_summary_revenue_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = salary_totals_report + fringe_pool_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
            
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        #Create new workbook and add results onto it
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\RESIDUAL_MFR.xlsx')
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\RESIDUAL_MFR.xlsx")
        worksheet = workbook.add_worksheet("RESIDUAL_RECON")
    
    
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            if (each_ytd_summary_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] +
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] +
                each_payroll_totals[acct] != 0):
            
            
            
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                    worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
            
            
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format_bold)
        
        worksheet.set_column('A:Z', 22)
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        
        workbook.close()
    
    
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")
        

#8
def flex_fund(department,fund,sof,period,project,flex,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label):
    #Query Variables
    global payroll_query
    global kk_enc_query
    global kk_exp_crefn_query
    global kk_exp_uflor_query
    global budget_query
    global ledger_query
    global excel_sheet_open_error
    #Get all of the reports that end with .xls from the base directory and append them to a list
    #base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FLEX_FUND"
    excel_file_open_validation = False
    base_dir = str(""r""+es_reports_file_path)
    filename = r"*.xls"
    filename_as_list = []
    file_list = []
    i = 0
    flex_code = flex
    for files in os.listdir(es_reports_file_path):
        if files.find(".xls") > 0:
            if files.find(".xlsx") > 0:
                continue
            else:
                joined_directories = os.path.join(base_dir,filename)
                items_found = glob.glob(joined_directories)
                file_list.append(items_found[i])
                filename_as_list.append(items_found[i])
                print(re.split('-', filename_as_list[i]) [-1])
            i += 1
            
    #Create new Temporary Folder
    newpath = base_dir + "\Temporary xlsx"
    if not os.path.exists(newpath):
        os.makedirs(newpath)
        
    #Add all new xlsx files to new Temporary Folder
    i = 0
    
    for items in file_list:
            temporary_directory = base_dir + "\Temporary xlsx\\" + re.split('-', filename_as_list[i]) [-1]
            print(temporary_directory)
            fname = file_list[i]
            #excel = win32.Dispatch()
            #ORIGINAL vv
            #excel = win32.gencache.EnsureDispatch('Excel.Application')
            try:
                excel = win32.dynamic.Dispatch("Excel.Application")
                wb = excel.Workbooks.Open(fname)
                wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
                wb.Close()                               #FileFormat = 56 is for .xls extension
                excel.Application.Quit()
                excel_sheet_open_error = "Nothing"
            except AttributeError as error_as_string_two:
                #excel.Application.Quit()
                print("Error as String: " + str(error_as_string_two))
                tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
                excel_file_open_validation = True
                print(str(excel_file_open_validation) + " is the Error Validtion")
                excel_sheet_open_error = "ExcelSheetOpenError"
                send2trash(base_dir + "\Temporary xlsx")
                break
                #print("FileNotFoundError")
                #return 'FileNotFoundError'
            i += 1
    
    #Report Sheets - Appropriations uses a different header number
    #temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\FLEX_FUND\Temporary xlsx'
    if excel_file_open_validation == False:
        temporary_file_path = newpath
        
        cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted Cash_Summary_Excel.xlsx")
        kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,1,second_root,my_label,"Extracted KK_to_GL_Summary_Comparison_Excel.xlsx")
        open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
        step_progress_bar(my_progress,2,second_root,my_label,"Extracted Appropriations_Summary_Excel.xlsx")
        
        all_accounts = ()
        
        
        #Get all the unique variables in the query file
        unique_accounts_query = ledger_query['ACCOUNT'].dropna()
        
        #Get all the unique variables in the report files
        unique_cash_summary_report = cash_summary_report['Account Code'].dropna()
        len(unique_cash_summary_report)
        
        unique_kk_to_gl_summary_report = kk_to_gl_summary_report['Account Code'].dropna()
        len(unique_kk_to_gl_summary_report)
        
        unique_open_encum_report = open_encum_report['Account Code'].dropna()
        len(unique_open_encum_report)
        
        unique_payroll_recon_report = payroll_recon_report['Account Code'].dropna()
        len(unique_payroll_recon_report)
        
        unique_transaction_detail_report = transaction_detail_report['Account Code'].dropna()
        len(unique_transaction_detail_report)
        
        #all_accounts = unique_accounts_query + unique_accounts_appropriations_report
        all_accounts = np.concatenate((unique_accounts_query, 
                                      unique_kk_to_gl_summary_report,
                                      unique_open_encum_report,
                                      unique_payroll_recon_report,
                                      unique_transaction_detail_report,
                                      unique_transaction_detail_report))
        
        
        
        all_unique_accounts = np.sort(np.unique(all_accounts))
        all_unique_accounts
        
        step_progress_bar(my_progress,10,second_root,my_label,'Set All Unique Accounts To A Variable')
        
        #print("Total in Length: " + str(len(all_accounts)))
        #print("All unique accounts accross all raw data: \n" + all_unique_accounts)
    
        #1) Query Data Set into dictionaries format is -> account:amount
        #REMEMBER acct is a dynamic variable and must be maintained like this
        fringe = ''
        fringe_or_not = {}
        each_ytd_summary = {}
        each_mtd_summary = {}
        each_glkk_totals_ytd = {}
        each_open_enc = {}
        each_payroll_totals = {}
        
        for acct in all_unique_accounts:
            ytd_summary = ledger_query.loc[
                        (ledger_query['CHARTFIELD1'] == flex_code) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund),
                         'POSTED_TOTAL_AMT'].sum()
            mtd_summary = ledger_query.loc[
                        (ledger_query['CHARTFIELD1'] == flex_code) &
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['FUND_CODE'] == fund)&
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            glkk_totals_ytd = kk_exp_uflor_query.loc[
                        (kk_exp_uflor_query['FUND_CODE'] == fund) &
                        (kk_exp_uflor_query['CHARTFIELD1'] == flex_code) &
                        (kk_exp_uflor_query['ACCOUNT.1'] == acct),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            ytd_open_enc = kk_enc_query.loc[
                        (kk_enc_query['ACCOUNT.1'] == acct) &
                        (kk_enc_query['CHARTFIELD1'] == flex_code) &
                        (kk_enc_query['FUND_CODE'] == fund),
                        'SUM(A.MONETARY_AMOUNT)'].sum()
            
            
            if str(acct)[0] == '6' and str(acct)[4:6] == '20':
                fringe = 'Yes'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['CHARTFIELD1'] == flex_code) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = ledger_query.loc[
                        (ledger_query['ACCOUNT'] == acct) &
                        (ledger_query['CHARTFIELD1'] == flex_code) &
                        (ledger_query['FUND_CODE'] == fund) &
                        (ledger_query['ACCOUNTING_PERIOD'] == period),
                         'POSTED_TOTAL_AMT'].sum()
            else:
                fringe = 'No'
                payroll_totals_mtd_yes_1 = payroll_query.loc[
                        (payroll_query['ACCOUNT'] == acct) &
                        (payroll_query['CHARTFIELD1'] == flex_code) &
                        (payroll_query['FUND_CODE'] == fund),
                        'MONETARY_AMOUNT'].sum()
                payroll_totals_mtd_yes_2 = 0
                    
            fringe_or_not[acct] = fringe
            each_ytd_summary[acct] = ytd_summary
            each_mtd_summary[acct] = mtd_summary
            each_glkk_totals_ytd[acct] = glkk_totals_ytd
            each_open_enc[acct] = ytd_open_enc
            each_payroll_totals[acct] = payroll_totals_mtd_yes_1 + payroll_totals_mtd_yes_2
            print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
            
        #print(each_payroll_totals.get(655120))
        step_progress_bar(my_progress,10,second_root,my_label,"Set Query Data into various dictionaries")
        #2) Report Data into dictionaries format is -> account:amount
        fringe = ''
        fringe_or_not_report = {}
        each_ytd_summary_report = {}
        each_mtd_summary_report = {}
        each_tran_detail_report = {}
        each_kk_totals_ytd_report = {}
        each_gl_totals_ytd_report = {}
        each_open_encumbrance_report = {}
        each_payroll_totals_report = {}
        
        for acct in all_unique_accounts:
            
            ytd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Flex Code'] == flex_code) &
                            (cash_summary_report['Fund Code'] == fund),
                             'YTD Expense'].sum()
            ytd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Flex Code'] == flex_code) &
                            (cash_summary_report['Fund Code'] == fund),
                            'YTD Revenue'].sum()
            mtd_summary_expense_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Flex Code'] == flex_code) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Expense'].sum()
            mtd_summary_revenue_report = cash_summary_report.loc[
                            (cash_summary_report['Account Code'] == acct) &
                            (cash_summary_report['Flex Code'] == flex_code) &
                            (cash_summary_report['Fund Code'] == fund),
                             'MTD Revenue'].sum()
            tran_detail_report = transaction_detail_report.loc[
                            (transaction_detail_report['Account Code'] == acct) &
                            (transaction_detail_report['Flex Code'] == flex_code) &
                            (transaction_detail_report['Fund Code'] == fund),
                             'Amount'].sum()
            kk_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Flex Code'] == flex_code) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD KK Amount'].sum()
            gl_totals_ytd_report = kk_to_gl_summary_report.loc[
                            (kk_to_gl_summary_report['Account Code'] == acct) &
                            (kk_to_gl_summary_report['Flex Code'] == flex_code) &
                            (kk_to_gl_summary_report['Fund Code'] == fund),
                             'YTD GL Amount'].sum()
            open_encumbrance_report = open_encum_report.loc[
                            (open_encum_report['Account Code'] == acct) &
                            (open_encum_report['Flex'] == flex_code) &
                            (open_encum_report['Fund Code'] == fund),
                             'Open Amount'].sum()
            salary_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Department Flex Code'] == flex_code) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund),
                             'Salary'].sum()
            fringe_pool_totals_report = payroll_recon_report.loc[
                            (payroll_recon_report['Account Code'] == acct) &
                            (payroll_recon_report['Department Flex Code'] == flex_code) &
                            (payroll_recon_report['Fund Code'] == fund)&
                            (payroll_recon_report['Fund Code'] == fund),
                             'Fringe Pool Amount'].sum()
        
            fringe_or_not_report[acct] = fringe
            each_ytd_summary_report[acct] = ytd_summary_expense_report +  ytd_summary_revenue_report
            each_mtd_summary_report[acct] = mtd_summary_expense_report + mtd_summary_revenue_report
            each_tran_detail_report[acct] = tran_detail_report
            each_kk_totals_ytd_report[acct] = kk_totals_ytd_report
            each_gl_totals_ytd_report[acct] = gl_totals_ytd_report
            each_open_encumbrance_report[acct] = open_encumbrance_report
            each_payroll_totals_report[acct] = salary_totals_report + fringe_pool_totals_report
        
            print(f'{acct} : {fringe_or_not_report[acct]} : {each_ytd_summary_report[acct]} : {each_mtd_summary_report[acct]} : {each_tran_detail_report[acct] } :{each_kk_totals_ytd_report[acct]} : {each_gl_totals_ytd_report[acct]} : {each_open_encumbrance_report[acct]} : {each_payroll_totals_report[acct]}')
         
        step_progress_bar(my_progress,10,second_root,my_label,"Set ES Report Data to various dictionaries")
        #Create new workbook and add results onto it
        #workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\FLEX_FUND_MFR.xlsx')
        newpath = str(""r""+file_path_for_saving)
        workbook = xlsxwriter.Workbook(newpath + "\FLEX_FUND_MFR.xlsx")
        worksheet = workbook.add_worksheet("FLEX_FUND_RECON")
    
    
        
        row = 3
        col = 3
        
        #Formats
        number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','border': 1})
        number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)','bold': True,'border': 1})
        bold = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font = workbook.add_format({'bold': True,'border': 1})
        bold_larger_font.set_font_size(12)
        border =  workbook.add_format({'border': 1})
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BFBFBF'})
        merge_format.set_font_size(14)
        worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
        worksheet.merge_range('L3:P3', 'Query Data', merge_format)
        worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,"Created and saved a new workbook")
        
        #Accounts
        worksheet.write(3, 2,"Accounts",bold_larger_font)
        
        #Report Header
        worksheet.write(3, 3,"YTD Summary",bold_larger_font)
        worksheet.write(3, 4,"MTD Summary",bold_larger_font)
        worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
        worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
        worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
        
        #Query Header
        worksheet.write(3, 11,"YTD Summary",bold_larger_font)
        worksheet.write(3, 12,"MTD Summary",bold_larger_font)
        worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
        worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
        worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
        
        
        #Variance Header
        
        worksheet.write(3, 17,"YTD Summary",bold_larger_font)
        worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
        worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
        worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
        worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
        worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
        worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
        worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote headers to workbook')
        
        for acct in all_unique_accounts:
            
            if (each_ytd_summary_report[acct] + 
                each_mtd_summary_report[acct] + 
                each_tran_detail_report[acct] + 
                each_kk_totals_ytd_report[acct] + 
                each_gl_totals_ytd_report[acct] + 
                each_open_encumbrance_report[acct] +
                each_payroll_totals_report[acct] +
                each_ytd_summary[acct] + 
                each_mtd_summary[acct] + 
                each_glkk_totals_ytd[acct] +
                each_open_enc[acct] +
                each_payroll_totals[acct] != 0):
            
                    row += 1
                    #Write Accounts
                    worksheet.write(row, 2,acct,border)
                    
                    #Write Report Results
                    worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
                    worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
                    worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
                    worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
                    worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
                    worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
                    
                 
                    #Write Query Results
                    worksheet.write(row, 11,each_ytd_summary[acct],number_format)
                    worksheet.write(row, 12,each_mtd_summary[acct],number_format)
                    worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
                    worksheet.write(row, 14,each_open_enc[acct],number_format)
                    worksheet.write(row, 15,each_payroll_totals[acct],number_format)
                    #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
                    
                    #Variance Results
                    worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
                    worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
                    worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
                    worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
                    worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
                    worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
        
        
        step_progress_bar(my_progress,10,second_root,my_label,'Wrote Report, Query, and Variance Results to workook')
            
        #Total Amounts
        worksheet.write(row + 1, 2,"Totals: ",bold)
        worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format_bold)
        worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format_bold)
        
        worksheet.set_column('A:Z', 22)
        
        
        #Cost Center  department,fund,sof,period,es_reports_file_path,file_path_for_saving,second_root,my_progress,my_label
        worksheet.merge_range('AA3:AB3', 'Cost Center', merge_format)
        cost_center_format_bold = workbook.add_format({'bold': True,'border': 1})
        
        worksheet.write(3, 26,"Department: ",cost_center_format_bold)
        worksheet.write(4, 26,"Fund:",cost_center_format_bold)
        worksheet.write(5, 26,"SoF:",cost_center_format_bold)
        worksheet.write(6, 26,"Period:",cost_center_format_bold)
        worksheet.write(7, 26,"Project:",cost_center_format_bold)
        worksheet.write(8, 26,"Flex:",cost_center_format_bold)
        
        worksheet.write(3, 27,department,cost_center_format_bold)
        worksheet.write(4, 27,fund,cost_center_format_bold)
        worksheet.write(5, 27,sof,cost_center_format_bold)
        worksheet.write(6, 27,period,cost_center_format_bold)
        worksheet.write(7, 27,project,cost_center_format_bold)
        worksheet.write(8, 27,flex,cost_center_format_bold)
        
        workbook.close()
        
        #Send Temporary Folder To Trash
        send2trash(base_dir + "\Temporary xlsx")
        #Progress Bar
        step_progress_bar(my_progress,10,second_root,my_label,"Wrote total amounts to workbook")
        step_progress_bar(my_progress,10,second_root,my_label,"Script Complete!")
        

def create_workbook(all_unique_accounts,each_ytd_summary,each_mtd_summary,each_glkk_totals_ytd,each_open_enc,each_payroll_totals,each_ytd_summary_report,each_mtd_summary_report,each_tran_detail_report,each_kk_totals_ytd_report,each_gl_totals_ytd_report,each_open_encumbrance_report,each_payroll_totals_report):
    #Create new workbook and add results onto it
    workbook = xlsxwriter.Workbook(r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results\DEPT_CASH_MFR.xlsx')
    worksheet = workbook.add_worksheet("DEPT_CASH_RECON")
    
    row = 3
    col = 3
    
    #Formats
    number_format = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00),','border': 1})
    number_format_bold = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00),','bold': True,'border': 1})
    bold = workbook.add_format({'bold': True,'border': 1})
    bold_larger_font = workbook.add_format({'bold': True,'border': 1})
    bold_larger_font.set_font_size(12)
    border =  workbook.add_format({'border': 1})
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#BFBFBF'})
    merge_format.set_font_size(14)
    worksheet.merge_range('C3:J3', 'Reports Data', merge_format)
    worksheet.merge_range('L3:P3', 'Query Data', merge_format)
    worksheet.merge_range('R3:Y3', 'Variance Data', merge_format)
    
    
    
    
    #Accounts
    worksheet.write(3, 2,"Accounts",bold_larger_font)
    
    #Report Header
    worksheet.write(3, 3,"YTD Summary",bold_larger_font)
    worksheet.write(3, 4,"MTD Summary",bold_larger_font)
    worksheet.write(3, 5,"Tran Detail Report MTD",bold_larger_font)
    worksheet.write(3, 6,"KK Totals YTD",bold_larger_font)
    worksheet.write(3, 7,"GL Totals YTD",bold_larger_font)
    worksheet.write(3, 8,"Open Encumbrance",bold_larger_font)
    worksheet.write(3, 9,"Payroll Totals MTD",bold_larger_font)
    
    #Query Header
    worksheet.write(3, 11,"YTD Summary",bold_larger_font)
    worksheet.write(3, 12,"MTD Summary",bold_larger_font)
    worksheet.write(3, 13,"GL/KK  Totals YTD",bold_larger_font)
    worksheet.write(3, 14,"YTD Open Enc = Summary",bold_larger_font)
    worksheet.write(3, 15,"Payroll Totals MTD",bold_larger_font)
    
    
    #Variance Header
    
    worksheet.write(3, 17,"YTD Summary",bold_larger_font)
    worksheet.write(3, 18,"MTD Summary ",bold_larger_font)
    worksheet.write(3, 19,"Tran Detail Report MTD",bold_larger_font)
    worksheet.write(3, 20,"KK Totals YTD (Variance to Query)",bold_larger_font)
    worksheet.write(3, 21,"KK/GL Totals YTD (Variance KK/GL)",bold_larger_font)
    worksheet.write(3, 22,"GL Totals YTD (Variance to YTD Summary)",bold_larger_font)
    worksheet.write(3, 23,"Open Encumbrance",bold_larger_font)
    worksheet.write(3, 24,"Payroll Totals MTD",bold_larger_font)
    
    
    for acct in all_unique_accounts:
        row += 1
        #Write Accounts
        worksheet.write(row, 2,acct,border)
        
        #Write Report Results
        worksheet.write(row, 3,each_ytd_summary_report[acct],number_format)
        worksheet.write(row, 4,each_mtd_summary_report[acct],number_format)
        worksheet.write(row, 5,each_tran_detail_report[acct],number_format)
        worksheet.write(row, 6,each_kk_totals_ytd_report[acct],number_format)
        worksheet.write(row, 7,each_gl_totals_ytd_report[acct],number_format)
        worksheet.write(row, 8,each_open_encumbrance_report[acct],number_format)
        worksheet.write(row, 9,each_payroll_totals_report[acct],number_format)
        
     
        #Write Query Results
        worksheet.write(row, 11,each_ytd_summary[acct],number_format)
        worksheet.write(row, 12,each_mtd_summary[acct],number_format)
        worksheet.write(row, 13,each_glkk_totals_ytd[acct],number_format)
        worksheet.write(row, 14,each_open_enc[acct],number_format)
        worksheet.write(row, 15,each_payroll_totals[acct],number_format)
        #print(f'{acct} : {fringe_or_not[acct]} : {each_ytd_summary[acct]} : {each_mtd_summary[acct]} : {each_glkk_totals_ytd[acct]} : {each_open_enc[acct]} : {each_payroll_totals[acct]}')
        
        #Variance Results
        worksheet.write(row, 17,"=D"+str(row + 1)+"-L"+str(row + 1),number_format)
        worksheet.write(row, 18,"=E"+str(row + 1)+"-M"+str(row + 1),number_format)
        worksheet.write(row, 19,"=F"+str(row + 1)+"-M"+str(row + 1),number_format)
        worksheet.write(row, 20,"=G"+str(row + 1)+"-N"+str(row + 1),number_format)
        worksheet.write(row, 21,"=G"+str(row + 1)+"-H"+str(row + 1),number_format)
        worksheet.write(row, 22,"=H"+str(row + 1)+"-N"+str(row + 1),number_format)
        worksheet.write(row, 23,"=I"+str(row + 1)+"-O"+str(row + 1),number_format)
        worksheet.write(row, 24,"=J"+str(row + 1)+"-P"+str(row + 1),number_format)
    
        
        
    #Total Amounts
    worksheet.write(row + 1, 2,"Totals: ",bold)
    worksheet.write(row + 1, 3,"=sum(D4:D"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 4,"=sum(E4:E"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 5,"=sum(F4:F"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 6,"=sum(G4:G"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 7,"=sum(H4:H"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 8,"=sum(I4:I"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 9,"=sum(J4:J"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 11,"=sum(L4:L"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 12,"=sum(M4:M"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 13,"=sum(N4:N"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 14,"=sum(O4:O"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 15,"=sum(P4:P"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 17,"=sum(R4:R"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 18,"=sum(S4:S"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 19,"=sum(T4:T"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 20,"=sum(U4:U"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 21,"=sum(V4:V"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 22,"=sum(W4:W"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 23,"=sum(X4:X"+ str(row +1) + ")",number_format_bold)
    worksheet.write(row + 1, 24,"=sum(Y4:Y"+ str(row +1) + ")",number_format_bold)
    
    worksheet.set_column('A:Z', 22)
    workbook.close()    
    


if __name__ == '__main__':
    main()
