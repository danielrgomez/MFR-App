# -*- coding: utf-8 -*-
"""
Created on Tue Nov  3 17:30:55 2020

@author: dgomezpe
"""

import pandas as pd
import os
import glob
import win32com.client as win32
import re
import numpy as np
import xlsxwriter
from send2trash import send2trash

query_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\2021_PER_3_Queries_UFLOR.xlsx'
es_reports_file_path = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_APPROP"


#Query Sheets
payroll_query = pd.read_excel(query_file_path,'PAYROLL',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str})
kk_enc_query = pd.read_excel(query_file_path,'KK_ENC',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str})
kk_exp_crefn_query = pd.read_excel(query_file_path,'KK_EXP_CREFN',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str})
kk_exp_uflor_query = pd.read_excel(query_file_path,'KK_EXP_UFLOR',converters={'ACCOUNT.1':str,'DEPTID':str,'FUND_CODE':str})
budget_query = pd.read_excel(query_file_path,'BUDGET',converters={'ACCOUNT':str,'DEPTID':str,'FUND_CODE':str})
ledger_query = pd.read_excel(query_file_path,'LEDGER',converters={'ACCOUNT':str,'DEPTID':str,'ACCOUNTING_PERIOD':str,'FUND_CODE':str})

      
#Get all of the reports that end with .xls from the base directory and append them to a list
base_dir = r"C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_APPROP"
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
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(temporary_directory + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    i += 1
print("All are saved!")


#Report Sheets - Appropriations uses a different header number
temporary_file_path = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Reports From ES\DEPT_APPROP\Temporary xlsx'

appropriations_report = pd.read_excel(temporary_file_path+'\Appropriations_Summary_Excel.xlsx','Sheet1',header=1,converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
budget_transaction_report = pd.read_excel(temporary_file_path+'\Budget_Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
cash_summary_report = pd.read_excel(temporary_file_path+'\Cash_Summary_Excel.xlsx','Sheet1', converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
kk_to_gl_summary_report = pd.read_excel(temporary_file_path+'\KK_to_GL_Summary_Comparison_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
open_encum_report = pd.read_excel(temporary_file_path+'\Open_Encumbrance_Summary_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
payroll_recon_report = pd.read_excel(temporary_file_path+'\Payroll_Reconciliation_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
projected_payroll_report = pd.read_excel(temporary_file_path+'\Projected_Payroll_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})
transaction_detail_report = pd.read_excel(temporary_file_path+'\Transaction_Detail_Excel.xlsx','Sheet1',converters={'Project Code': str,'DeptID':str,'Account Code':str,'Fund Code':str})



#Hard coded variables. Will change to user input afterwards
department = '19050100'
fund = '101'
sof = ''
period = '3'
account = '611110'

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
    
#print(each_payroll_totals.get(655120))

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
newpath = r'C:\Miscellaneous\Macros\MFR Review\3. September 2020\Financial Review\Python MFR Results'
if not os.path.exists(newpath):
    os.makedirs(newpath)


#Create new workbook and add results onto it
workbook = xlsxwriter.Workbook(newpath + '\DEPT_APPROP_MFR.xlsx')
worksheet = workbook.add_worksheet("DEPT_APPROP_RECON")
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

workbook.close()

#Send Temporary Folder To Trash
send2trash(base_dir + "\Temporary xlsx")

















