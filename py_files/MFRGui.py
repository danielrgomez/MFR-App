# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 17:22:59 2020

@author: dgomezpe
"""

from tkinter import *
#import tkinter
import pandas as pd
import os
from tempfile import NamedTemporaryFile
import shutil
import csv
import MonthtlyFinancialReviewVersion1 as mfr
import tkinter.messagebox
from tkinter import ttk
import time



#global validation_to_assign_query
validation_to_assign_query = 1

def open_saved_file(clicked_variable):
    #print(savePathLocationEntry)
    own_working_directory = os.path.abspath(os.getcwd())
    #own_working_directory = own_working_directory = r"C:/Miscellaneous/Macros/MFR Review/Python Solution/"
    name_of_file = "userInput"
    complete_name = os.path.join(own_working_directory, name_of_file+".csv")    
    userInputFile = pd.read_csv(complete_name,converters={'Report':str,'Department':str,'Fund':str,'Sof':str,'Period':str,'Project':str,'Flex':str,'FilePath':str,'QueryFilePath':str,'SaveFilePath':str})
    save_path_as_text = (userInputFile.loc[userInputFile['Report'] == clicked.get()]['SaveFilePath'].values[0])
    file = save_path_as_text + '\\' + clicked_variable +'_MFR.xlsx'
    
    if clicked_variable == "All Reports":
        for reports in userInputFile['Report']:
            if reports != "All Reports":
                try:
                    file = save_path_as_text + '\\' + reports +'_MFR.xlsx'
                    os.startfile(file)
                except:
                    print("Could not find file!")
                    tkinter.messagebox.showinfo("File Not Found!",  clicked_variable + " File Not Found in File Path: \n\n" + save_path_as_text )
    else:
        try:
            os.startfile(file)
        except:
            print("Could not find file!")
            tkinter.messagebox.showinfo("File Not Found!",  clicked_variable + " File Not Found in File Path: \n\n" + save_path_as_text )

def set_validation():
    global validation_to_assign_query
    validation_to_assign_query = 0
    print("After setting the validation: " + str(validation_to_assign_query))



def showAllTextboxes(somestring):
    global departmentAsString
    global fundAsString
    global sofAsString
    global periodAsString
    global projectAsString
    global flexAsString
    global filePathAsString
    global queryPathAsString
    global savePathLocationAsString
    
    
    departmentAsString = StringVar()
    fundAsString = StringVar()
    sofAsString = StringVar()
    periodAsString = StringVar()
    projectAsString = StringVar()
    flexAsString = StringVar()
    filePathAsString = StringVar()
    queryPathAsString = StringVar()
    savePathLocationAsString = StringVar()
    
    #Read CSV and take user input and place in text fields
    #Gets own directory instead
    own_working_directory = os.path.abspath(os.getcwd())
    #own_working_directory = own_working_directory = r"C:/Miscellaneous/Macros/MFR Review/Python Solution/"
    name_of_file = "userInput"
    complete_name = os.path.join(own_working_directory, name_of_file+".csv")    
    
    userInputFile = pd.read_csv(complete_name,converters={'Report':str,'Department':str,'Fund':str,'Sof':str,'Period':str,'Project':str,'Flex':str,'FilePath':str,'QueryFilePath':str,'SaveFilePath':str})
        
    if somestring != "All Reports":
        departmentEntry = Entry(root,textvariable = departmentAsString)
        fundEntry = Entry(root,textvariable = fundAsString)
        sofEntry = Entry(root,textvariable = sofAsString)
        periodEntry = Entry(root,textvariable = periodAsString)
        projectEntry = Entry(root,textvariable = projectAsString)
        flexEntry = Entry(root,textvariable = flexAsString)
        filePathEntry = Entry(root,textvariable = filePathAsString)
        queryPathEntry = Entry(root,textvariable = queryPathAsString)
        savePathLocationEntry = Entry(root,textvariable = savePathLocationAsString)
        
        
        #Department
        departmentLabel.grid(row = 2, column = 0)
        departmentLabel.place(x=87, y=68, height=18, width=86)
        departmentEntry.grid(row = 2, column = 3)
        departmentEntry.delete(0,'end')
        departmentEntry.place(x=170, y=68, height=20, width=157)
        departmentEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Department'].values[0]))
    
        #Fund
        fundLabel.grid(row = 3, column = 0)
        fundLabel.place(x=68, y=90, height=18, width=86)
        fundEntry.grid(row = 3, column = 3)
        fundEntry.delete(0,'end')
        fundEntry.place(x=170, y=90, height=20, width=157)
        fundEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Fund'].values[0]))
    
        #Sof
        sofLabel.grid(row = 4, column = 0)
        sofLabel.place(x=64, y=112, height=18, width=86)
        sofEntry.grid(row = 4, column = 3)
        sofEntry.delete(0,'end')
        sofEntry.place(x=170, y=112, height=20, width=157)
        sofEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Sof'].values[0]))
    
        #Period
        periodLabel.grid(row = 5, column = 0)
        periodLabel.place(x=95, y=134, height=18, width=37)
        periodEntry.grid(row = 5, column = 3)
        periodEntry.delete(0,'end')
        periodEntry.place(x=170, y=134, height=20, width=157)
        periodEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Period'].values[0]))
    
        #Project
        projectLabel.grid(row = 6, column = 0)
        projectLabel.place(x=95, y=156, height=18, width=40)
        projectEntry.grid(row = 6, column = 3)
        projectEntry.delete(0,'end')
        projectEntry.place(x=170, y=156, height=20, width=157)
        projectEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Project'].values[0]))
    
        #Flex
        flexLabel.grid(row = 7, column = 0)
        flexLabel.place(x=87, y=178, height=18, width=40)
        flexEntry.grid(row = 7, column = 3)
        flexEntry.delete(0,'end')
        flexEntry.place(x=170, y=178, height=20, width=157)
        flexEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Flex'].values[0]))
    
        #Es Reports File Path
        es_reports_file_path.grid(row = 8, column = 0)
        es_reports_file_path.place(x=0, y=203, height=18, width=120)
        filePathEntry.delete(0,'end')
        filePathEntry.place(x=132, y=203, height=18, width=315)
        filePathEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['FilePath'].values[0]))
        
        #Query Report File Path
        query_reports_file_path.grid(row = 9, column = 0)
        query_reports_file_path.place(x=0, y=223, height=18, width=140)
        queryPathEntry.grid(row = 9, column = 3)
        queryPathEntry.delete(0,'end')
        queryPathEntry.place(x=132, y=223, height=18, width=315)
        queryPathEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['QueryFilePath'].values[0]))
        
        #Save location file path
        savel_location_file_path.grid(row = 10, column = 0)
        savel_location_file_path.place(x=0, y=244, height=18, width=140)
        savePathLocationEntry.grid(row = 10, column = 3)
        savePathLocationEntry.delete(0,'end')
        savePathLocationEntry.place(x=132, y=244, height=18, width=315)
        savePathLocationEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['SaveFilePath'].values[0]))
            
        
    elif somestring == "All Reports": 
        #Disable Department Entry
        departmentEntry = Entry(root,textvariable = departmentAsString)
        departmentEntry.insert(0,"Create All Reports")
        departmentEntry.configure(state="disabled")
        departmentEntry.update()
       
        #Disable Fund Entry
        fundEntry = Entry(root,textvariable = fundAsString)
        fundEntry.insert(0,"Create All Reports")
        fundEntry.configure(state="disabled")
        fundEntry.update()
        
        #Disable Sof Entry
        sofEntry = Entry(root,textvariable = sofAsString)
        sofEntry.insert(0,"Create All Reports")
        sofEntry.configure(state="disabled")
        sofEntry.update()
        
        #Disable Period Entry
        periodEntry = Entry(root,textvariable = periodAsString)
        periodEntry.insert(0,"Create All Reports")
        periodEntry.configure(state="disabled")
        periodEntry.update()
        
        #Disable Project Entry
        projectEntry = Entry(root,textvariable = projectAsString)
        projectEntry.insert(0,"Create All Reports")
        projectEntry.configure(state="disabled")
        projectEntry.update()
        
        #Disable Flex Entry
        flexEntry = Entry(root,textvariable = flexAsString)
        flexEntry.insert(0,"Create All Reports")
        flexEntry.configure(state="disabled")
        flexEntry.update()
            
        #Disable File Path Entry
        filePathEntry = Entry(root,textvariable = filePathAsString)
        #filePathEntry.delete(0,'end')
        filePathEntry.insert(0,"Create All Reports")
        filePathEntry.configure(state="disabled")
        filePathEntry.update()
                
        #Query Report File Path
        queryPathEntry = Entry(root,textvariable = queryPathAsString)
        #queryPathEntry.insert(0,"Create All Reports")
        queryPathEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['QueryFilePath'].values[0]))
        #queryPathEntry.update()
        
        #Save location file path
        savePathLocationEntry = Entry(root,textvariable = savePathLocationAsString)
        #savePathLocationEntry.insert(0,"Create All Reports")
        savePathLocationEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['SaveFilePath'].values[0]))
        #savePathLocationEntry.update()

    #Department
    departmentLabel.place(x=87, y=68, height=18, width=86)
    departmentEntry.delete(0,'end')
    departmentEntry.place(x=170, y=68, height=20, width=157)
    departmentEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Department'].values[0]))

    #Fund
    fundLabel.place(x=68, y=90, height=18, width=86)
    fundEntry.delete(0,'end')
    fundEntry.place(x=170, y=90, height=20, width=157)
    fundEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Fund'].values[0]))

    #Sof
    sofLabel.place(x=64, y=112, height=18, width=86)
    sofEntry.delete(0,'end')
    sofEntry.place(x=170, y=112, height=20, width=157)
    sofEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Sof'].values[0]))

    #Period
    periodLabel.place(x=95, y=134, height=18, width=37)
    periodEntry.delete(0,'end')
    periodEntry.place(x=170, y=134, height=20, width=157)
    periodEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Period'].values[0]))

    #Project
    projectLabel.place(x=95, y=156, height=18, width=40)
    projectEntry.delete(0,'end')
    projectEntry.place(x=170, y=156, height=20, width=157)
    projectEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Project'].values[0]))

    #Flex
    flexLabel.place(x=87, y=178, height=18, width=40)
    flexEntry.delete(0,'end')
    flexEntry.place(x=170, y=178, height=20, width=157)
    flexEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['Flex'].values[0]))

    #Es Reports File Path
    es_reports_file_path.place(x=0, y=203, height=18, width=120)
    filePathEntry.delete(0,'end')
    filePathEntry.place(x=132, y=203, height=18, width=315)
    filePathEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['FilePath'].values[0]))
    
    #Query Report File Path
    query_reports_file_path.place(x=0, y=223, height=18, width=140)
    queryPathEntry.delete(0,'end')
    queryPathEntry.place(x=132, y=223, height=18, width=315)
    queryPathEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['QueryFilePath'].values[0]))
    
    #Save location file path
    savel_location_file_path.place(x=0, y=244, height=18, width=140)
    savePathLocationEntry.delete(0,'end')
    savePathLocationEntry.place(x=132, y=244, height=18, width=315)
    savePathLocationEntry.insert(0,str(userInputFile.loc[userInputFile['Report'] == clicked.get()]['SaveFilePath'].values[0]))

def updateAllTextboxes(complete_name,clicked,root):
 global validation_to_assign_query
 textBoxesAsList = [clicked,
                    departmentAsString.get(),
                    fundAsString.get(),
                    sofAsString.get(),
                    periodAsString.get(),
                    projectAsString.get(),
                    flexAsString.get(),
                    filePathAsString.get(),
                    queryPathAsString.get(),
                    savePathLocationAsString.get()]

 print(textBoxesAsList)
 tkinter.messagebox.showinfo("Saved!", str(clicked) + " Saved")


 filename = complete_name
 tempfile = NamedTemporaryFile(mode='w', delete=False)

 fields = ['Report','Department', 'Fund', 'Sof', 'Period','Project','Flex','FilePath','QueryFilePath','SaveFilePath']

 with open(filename, 'r', encoding='ascii') as csvfile, tempfile:
     reader = csv.DictReader(csvfile, fieldnames=fields)
     writer = csv.DictWriter(tempfile, fieldnames=fields)
     for row in reader:
         if row['Report'] == str(textBoxesAsList[0]):
             print('updating row', row['Report'])
             row['Report'], row['Department'], row['Fund'], row['Sof'], row['Period'], row['Project'], row['Flex'], row['FilePath'],row['QueryFilePath'],row['SaveFilePath'] = textBoxesAsList[0], textBoxesAsList[1], textBoxesAsList[2], textBoxesAsList[3], textBoxesAsList[4], textBoxesAsList[5], textBoxesAsList[6], textBoxesAsList[7],textBoxesAsList[8],textBoxesAsList[9]
         row = {'Report':row['Report'],'Department': row['Department'], 'Fund': row['Fund'], 'Sof': row['Sof'], 'Period': row['Period'], 'Project': row['Project'], 'Period': row['Period'], 'Flex': row['Flex'], 'FilePath': row['FilePath'],'QueryFilePath':row['QueryFilePath'],'SaveFilePath':row['SaveFilePath']}
         writer.writerow(row)

 validation_to_assign_query = 1
 shutil.move(tempfile.name, filename)
 csvfile.close()

def gui_control(clicked_on_item):
    #Read CSV and take user input and place in text fields
    #Gets own directory instead
    own_working_directory = os.path.abspath(os.getcwd())
    #own_working_directory = own_working_directory = r"C:/Miscellaneous/Macros/MFR Review/Python Solution/"
    name_of_file = "userInput"
    complete_name = os.path.join(own_working_directory, name_of_file+".csv")    
    userInputFile = pd.read_csv(complete_name,converters={'Report':str,'Department':str,'Fund':str,'Sof':str,'Period':str,'Project':str,'Flex':str,'FilePath':str})
    user_input_as_list = []
    global validation_to_assign_query
    print('Before setting the validation: ' + str(validation_to_assign_query))
    if clicked_on_item != 'All Reports':
        
        
        
        #Second Root Is a Second Tkinter Window: Progress Bar
        #Within this Block It will begin to 
        second_root = Tk()
        second_root.title("Progress To Complete MFR Report")
        second_root.geometry("420x100")
        my_progress = ttk.Progressbar(second_root, orient = HORIZONTAL, length = 400, mode = 'determinate')
        my_progress.pack(pady = 20)
        running_as_text = "Running: "
        running_as_text = "{:>100}".format(running_as_text)
        my_label = Label(second_root,text = running_as_text)
        my_label.pack(pady = 1)
        #my_label.config(text = my_progress['value'])
               
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Report'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Department'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Fund'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Sof'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Period'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Project'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['Flex'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['FilePath'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['QueryFilePath'].values[0])
        user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['SaveFilePath'].values[0])
        validate_and_run = mfr.script_control(clicked.get(),user_input_as_list[1],user_input_as_list[2],user_input_as_list[3],user_input_as_list[4],user_input_as_list[5],user_input_as_list[6],user_input_as_list[7],user_input_as_list[8],user_input_as_list[9],validation_to_assign_query,second_root,my_progress,my_label)
        
        print(validate_and_run)
        #The no error validate and run validation is a way to ensure the no file found exception has not occurred
        if validate_and_run == "NoError":
            tkinter.messagebox.showinfo("Completed MFR Report","Completed MFR For Report: "+ str(clicked.get()))
            set_validation()
        elif validate_and_run == "ExcelSheetOpenError":
            #tkinter.messagebox.showinfo("Save and Close The Excel Application","Please make sure to save and close the Excel Application before creating any reports")
            set_validation()
        
        
        second_root.quit()
        second_root.destroy()
        #validation_to_assign_query = 0
        #tkinter.messagebox.showinfo("Completed MFR Report","Completed MFR For DEPT APPROP")
        #tkinter.messagebox.showinfo("Completed MFR Report","Completed MFR For Report: "+ str(clicked.get()))
        print("Before mainloop")
        second_root.mainloop()
        print("After mainloop")
                
        
        
    elif clicked_on_item == 'All Reports':
        
        
        
        for each_item in userInputFile['Report']:
            

            
            if each_item != 'All Reports' and each_item != '':
                
                second_root = Tk()
                second_root.title("Progress To Complete MFR Report")
                second_root.geometry("420x100")
                my_progress = ttk.Progressbar(second_root, orient = HORIZONTAL, length = 400, mode = 'determinate')
                my_progress.pack(pady = 20)
                running_as_text = "Running: "
                running_as_text = "{:>100}".format(running_as_text)
                my_label = Label(second_root,text = running_as_text)
                my_label.pack(pady = 1)
                #my_label.config(text = my_progress['value'])
                
                
                user_input_as_list = []
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Report'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Department'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Fund'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Sof'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Period'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Project'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['Flex'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == each_item]['FilePath'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['QueryFilePath'].values[0])
                user_input_as_list.append(userInputFile.loc[userInputFile['Report'] == clicked_on_item]['SaveFilePath'].values[0])
                mfr.script_control(each_item,user_input_as_list[1],user_input_as_list[2],user_input_as_list[3],user_input_as_list[4],user_input_as_list[5],user_input_as_list[6],user_input_as_list[7],user_input_as_list[8],user_input_as_list[9],validation_to_assign_query,second_root,my_progress,my_label)
                set_validation()
                second_root.quit()
                second_root.destroy()
        
        tkinter.messagebox.showinfo("Completed All MFR Reports","Completed All MFR Reports")
        validation_to_assign_query = 0
        #tkinter.messagebox.showinfo("Completed MFR Report","Completed MFR For DEPT APPROP")
        print("Before mainloop")
        second_root.mainloop()
        print("After mainloop")
                    
                
def dynamic_middlman(*args):
    print("the value changed...", clicked.get())
    showAllTextboxes(clicked.get())
    
#import MonthtlyFinancialReviewVersion1 as mfr
#Read CSV and take user input and place in text fields
#Gets own directory instead
own_working_directory = os.path.abspath(os.getcwd())
#own_working_directory = own_working_directory = r"C:/Miscellaneous/Macros/MFR Review/Python Solution/"
name_of_file = "userInput"
complete_name = os.path.join(own_working_directory, name_of_file+".csv")    
print(complete_name)

userInputFile = pd.read_csv(complete_name,converters={'Report':str,'Department':str,'Fund':str,'Sof':str,'Period':str,'Project':str,'Flex':str,'FilePath':str,'QueryFilePath':str,'SaveFilePath':str})

#Set Tkinter to variable root
root = Tk()
root.geometry("460x320")
root.title("Monthly Financial Review ")
root.resizable(0, 0) 
root.iconbitmap(own_working_directory + "\logo.ico")



#Labels
titleLabel = Label(root, text = 'Monthly Financial Review')
departmentLabel = Label(root, text = 'Department: ')
fundLabel = Label(root, text = 'Fund: ')
sofLabel = Label(root, text = 'SoF: ')
periodLabel = Label(root, text = 'Period: ')
projectLabel = Label(root, text = 'Project: ')
flexLabel = Label(root, text = 'Flex: ')
es_reports_file_path = Label(root, text = 'ES Reports File Path: ')
query_reports_file_path = Label(root, text = 'Query Reports File Path: ')
savel_location_file_path = Label(root, text = 'Save Location File Path: ')




#Drop Down Menu
options = [
    "All Reports",
    "DEPT_APPROP",
    "DEPT_CASH",
    "DEPT_STUGOV",
    "FUND_HOUSE",
    "FUND_CASH_DSO",
    "FUND_CASH",
    "RESIDUAL",
    "FLEX_FUND"
]

clicked = StringVar()
#clicked.trace is a way to trace the clicked variable so whenever it changes it will run a function. In this case whenever clicked changes dynamic_middlman is run
clicked.trace("w", dynamic_middlman)
#sets default options when the applciatin is first opened. 
clicked.set(options[1])
drop_down_menu_item = 'DEPT_CASH'
drop_down_menu = OptionMenu(root,clicked,*options)
clickedItem = clicked.get()


showAllTextboxes(clickedItem)


#Buttons
#createReportsButton = Button(root,text='Create Reports',padx = 50,pady = 5)
#searchUserInputButton = Button(root,text='Lookup',command= lambda: showAllTextboxes(clicked.get()))
updateUserInputButton = Button(root,text='Save',command= lambda: updateAllTextboxes(complete_name,clicked.get(),root))
#searchUserInputButton = Button(root,text='Lookup',command= showAllTextboxes(clickedItem))
createReportsButton = Button(root,text='Create Reports',padx = 50,pady = 5,command = lambda: gui_control(clicked.get()))
openButton = Button(root,text="Open Report",padx = 50, pady=50,command = lambda: open_saved_file(clicked.get()))


#Title location
titleLabel.place(x=169, y=5)

#Drop Down Menu Location
drop_down_menu.grid(row = 1, column = 2,padx=15)
drop_down_menu.place(x=175, y=30, height=30, width=130)

#Lookup Button Location
#searchUserInputButton.grid(row = 1,column = 3)

#Update User Input Button Location (Save)
#updateUserInputButton.grid(row = 11,column = 3)
updateUserInputButton.place(x=365, y=115, height=20, width=50)

#Create Reports Button
#createReportsButton.grid(row = 11, column = 2)
createReportsButton.place(x=120, y=270, height=40, width=100)

#Open Buttom
openButton.place(x = 240, y = 270, height = 40, width = 100)


root.mainloop() 

