#==========================================================================
# Program: FindLogging.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Feb-17-2018
# Date Last Modified: Feb-18-2018
# Summary: This is Class to for the Circuits
#==========================================================================

#***************************************************************
# ==================== Libraries Required <BEGIN> =============*
#***************************************************************

#************************ For PING ************************
import re  # Required for teh Class
import subprocess
from time import time, sleep
try:
    import socket
    import threading
#    fromthreading import *
except:
    print ("NO Sockets is available")
#************************ For PING ************************

import os
import sys
import math
import datetime
import time
import random
import tkinter
import tkinter.messagebox
import tkinter.filedialog
from tkinter import *           # Importing the Tkinter (tool box) library
from tkinter import ttk
if sys.version_info < (3,0): 
    import Tkinter as tkinter 
    import tkMessageBox as mbox 
    import Tkinter.font as tkfont
else: 
    import tkinter 
    import tkinter.messagebox as mbox 
    import tkinter.font as tkfont
#import PyPDF2


try:
    from odbc_connector import *
    Is_ODBC_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO ODBC Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_ODBC_Available = False
    sys.exit(1)

try:
    from SaveAs import *
    Is_SaveAs_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO SaveAs Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False


#*******************************************************
#============= READING VARIABLES TABLE ================*
#*******************************************************
try:
    from Utils import *
    Utils = Class_Utils()
    Utils.Get_Values()
    #------- DNS NAME ---------
    ODBC_DSN_name = Utils.Get_ODBC_Name()
    Windows_Scaling = Utils.Get_Windows_Scaling()
    #--------------------------
except:
    #------- DNS NAME ---------
    ODBC_DSN_name = "BV"
    Windows_Scaling = 1.0
    #--------------------------

#print (Windows_Scaling)

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Find Logging','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False


#*********************************************************************************************************************************************
#                                   Logging Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************
class Class_FindLogging:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.FindLoggingWindowExist = False
        self.Username = os.getlogin()
        self.date = ""
        self.Windows_Scaling = Windows_Scaling

    def treeview_sort_column(self,tv, col, reverse):
        #print('sorting %s!' % col)
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            #print('Moving Index:%r, Value:%r, k:%r' % (index, val, k))
            tv.move(k, '', index)

        # reverse sort next time
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def on_FindLoggingWindow_quit(self):
        if (self.FindLoggingWindowExist):
            self.FindLoggingWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Logging','CLOSED Window']    
                Logging.Log(Parameter)
            self.FindLoggingWindow.destroy()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.LoggingTreeview.focus()
        dic = self.LoggingTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.LoggingTreeview.next(curItem)
        self.LoggingTreeview.selection_set((NextItem, NextItem))
        self.LoggingTreeview.focus(NextItem)

        NextItem = self.LoggingTreeview.focus()
        dic = self.LoggingTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
            
    def on_Logging_Tree_select_click(self,event): # Not needed here for left for future use :)
            curItem = self.LoggingTreeview.focus()  
            dic = self.LoggingTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1


    def Call_ButtonFindLogging(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Logging','Find Button']    
            Logging.Log(Parameter)
        #self.buttonLogging['state'] = DISABLED
        #self.buttonLoggingFacility['state'] = DISABLED
        value = self.FindLoggingFrameEntry.get() # Value to querry
        index = (self.ComboBoxLoggingTableFields.current())
        field = (self.LoggingTableFieldsSQL[index])
        self.LoggingSites = "NO"
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == "*"):
                        sql = """SELECT * FROM Logging 
                             ORDER BY Log_Year desc, Log_Month desc, Log_Day desc, \
                             Log_Hour desc, Log_Min desc, Log_sec desc, User_ID asc \
                          """
                    else:
                        sql = """SELECT * FROM Logging 
                            WHERE %s LIKE '%s' ORDER BY Log_Year desc, Log_Month desc, Log_Day desc, \
                            Log_Hour desc, Log_Min desc, Log_sec desc \
                          """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Tree View <BEGIN> --------
                        x = self.LoggingTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.LoggingTreeview.delete(child)
                        #------------------------------- Deleting Tree View <END> --------
                        #-------------- Initializing Arrays <BEGIN> ----------------------
                        self.LoggingTableUserIDArray =  [] 
                        self.LoggingTableLogDateArray =  []
                        self.LoggingTableLogDayArray =  []
                        self.LoggingTableLogMonthArray =  []
                        self.LoggingTableLogYearArray =  []
                        self.LoggingTableLogTimeArray =  []
                        self.LoggingTableLogHourArray =  []
                        self.LoggingTableLogMinArray =  []
                        self.LoggingTableLogSecArray =  []
                        self.LoggingTableLogModuleArray =  []
                        self.LoggingTableLogFunctionArray =  []
                        self.LoggingTableExecutedByUserArray = []
                        self.LoggingTableArrayColumns = []
                        self.LoggingTableArrayTemp = []
                        self.results = []

                        self.LoggingTableArrayColumns = (
                            'User ID',
                            'Log Date',
                            'Log Day',
                            'Log Month',
                            'Log Year',
                            'Log Time',
                            'Log Hour',
                            'Log Min',
                            'Log Sec',
                            'Log Module',
                            'Log Function',
                            'Executed_by_UserID'
                            )
                        
                        #-------------- Initializing Arrays <END> ----------------------
                        self.data_ready = True
                        i = 0
                        while (i < len(self.db.results)):
                            self.LoggingTableUserIDArray.append(self.db.results[i][0].strip())
                            self.LoggingTableLogDateArray.append(self.db.results[i][1].strip())
                            self.LoggingTableLogDayArray.append(self.db.results[i][2])
                            self.LoggingTableLogMonthArray.append(self.db.results[i][3])
                            self.LoggingTableLogYearArray.append(self.db.results[i][4])
                            self.LoggingTableLogTimeArray.append(self.db.results[i][5].strip())
                            self.LoggingTableLogHourArray.append(self.db.results[i][6])
                            self.LoggingTableLogMinArray.append(self.db.results[i][7])
                            self.LoggingTableLogSecArray.append(self.db.results[i][8])
                            self.LoggingTableLogModuleArray.append(self.db.results[i][9].strip())
                            self.LoggingTableLogFunctionArray.append(self.db.results[i][10].strip())
                            self.LoggingTableExecutedByUserArray.append(self.db.results[i][11].strip())
                            i = i + 1
                        i = 0
                        while (i < len(self.LoggingTableUserIDArray)):
                            num = i + 1
                            tags = self.LoggingTableLogModuleArray[i]
                            item = [
                                    self.LoggingTableUserIDArray[i],
                                    self.LoggingTableLogDateArray[i],
                                    self.LoggingTableLogDayArray[i],
                                    self.LoggingTableLogMonthArray[i],
                                    self.LoggingTableLogYearArray[i],
                                    self.LoggingTableLogTimeArray[i],
                                    self.LoggingTableLogHourArray[i],
                                    self.LoggingTableLogMinArray[i],
                                    self.LoggingTableLogSecArray[i],
                                    self.LoggingTableLogModuleArray[i],
                                    self.LoggingTableLogFunctionArray[i],
                                    self.LoggingTableExecutedByUserArray[i]
                                    ]
                            self.results.append(item)
                            self.LoggingTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.FindLoggingFrame,title='Find User LOGGING Activity',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindLoggingFrame,title='Find User LOGGING Activity',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindLoggingFrame,title='Find User LOGGING Activity',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindLogging_From_Entry(self,event):
        self.Call_ButtonFindLogging()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Find User LOGGING Activity"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.LoggingTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.FindLoggingFrame,title='Find User LOGGING Activity List',
                            message = '!!! The File was saved !!!')
                    self.FindLoggingWindow.title("Find User LOGGING Activity                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Find User LOGGING Activity','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindLoggingFrame,title='Find User LOGGING Activity List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Find_Logging_Window(self):
        if not self.FindLoggingWindowExist:
            # Set up the Window
            self.FindLoggingWindowExist = True
            self.FindLoggingWindow = Tk()
            self.FindLoggingWindow.geometry('1350x700+350+70')
            self.FindLoggingWindow.title("Find User LOGGING Activity")
            self.FindLoggingWindow.protocol("WM_DELETE_WINDOW", self.on_FindLoggingWindow_quit)
            self.FindLoggingWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindLoggingLabel = Label(self.FindLoggingWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindLoggingLabel["text"] = "Find User LOGGING Activity"
            self.FindLoggingLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindLoggingWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindLoggingWindow_quit)
            self.FindLoggingWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.FindLoggingFrame = Frame(self.FindLoggingWindow)
            self.FindLoggingFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindLoggingFrame.rowconfigure(0, weight=1)
            self.FindLoggingFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindLoggingFrameLabel = Label(self.FindLoggingFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindLoggingFrameLabel["text"] = "Find:"
            self.FindLoggingFrameLabel.place(x=10, y=10)

            '''
                    User_ID              CHAR(20) NOT NULL,
                    Log_Date             CHAR(20) NOT NULL,
                    Log_Day              INT,
                    Log_Month            INT,
                    Log_Year             INT,
                    Log_Time             CHAR(20) NOT NULL,
                    Log_Hour             INT,
                    Log_Min              INT,
                    Log_Sec              INT,
                    Log_Module               CHAR(50) NOT NULL,
                    Log_Function             CHAR(50) NOT NULL,
                    Executed_by_UserID   CHAR(20))"""

            '''

            # Setup ComboBox
            self.LoggingTableFields = [
                                'User ID',
                                'Date',
                                'Day',
                                'Month',
                                'Year',
                                'Time',
                                'Hour',
                                'Min',
                                'Sec',
                                'Module',
                                'Function',
                                'Executed by UserID'
                                ]

            self.LoggingTableFieldsSQL = [
                                'User_ID',
                                'Log_Date',
                                'Log_Day',
                                'Log_Month',
                                'Log_Year',
                                'Log_Time',
                                'Log_Hour',
                                'Log_Min',
                                'Log_Sec',
                                'Log_Module',
                                'Log_Function',
                                'Executed_by_UserID'
                                ]

            self.ComboBoxLoggingTableFields = ttk.Combobox(self.FindLoggingFrame, state='readonly', width = 25)
            self.ComboBoxLoggingTableFields['values'] = self.LoggingTableFields
            self.ComboBoxLoggingTableFields.current(0)
            self.ComboBoxLoggingTableFields.place(x = 50, y = 10)

            # Setup Entry
            self.FindLoggingFrameEntry = Entry(self.FindLoggingFrame)
            self.FindLoggingFrameEntry['width']=50
            self.FindLoggingFrameEntry.place(x=235, y=10)
            self.FindLoggingFrameEntry.bind('<Return>', self.Call_ButtonFindLogging_From_Entry)
            #FindLoggingFrameEntry['state'] = DISABLED

            if self.db.Connect():                
                # Setup Button
                self.buttonFindLoggingExecute = Button(self.FindLoggingFrame, text = 'Execute', command = self.Call_ButtonFindLogging, state=ACTIVE)
                self.buttonFindLoggingExecute.place(x = 550, y = 7, width=75, height=25)

                #self.buttonLoggingFacility = Button(self.FindLoggingFrame, text = 'Facility', command = self.Call_Logging_ButtonFacility, state=DISABLED)
                #self.buttonLoggingFacility.place(x = 850, y = 7, width=75, height=25)

                #self.buttonLogging = Button(self.FindLoggingFrame, text = 'Point of Contacts', command = self.Call_Logging_ButtonLogging, state=DISABLED)
                #self.buttonLogging.place(x = 950, y = 7, width=140, height=25)

                    
                #------------------ TREE VIEW For Logging Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                '''
                                'User ID',
                                'Date',
                                'Day',
                                'Month',
                                'Year',
                                'Time',
                                'Hour',
                                'Min',
                                'Sec',
                                'Module',
                                'Function',
                                'Executed by UserID'
                '''
                self.LoggingTreeviewDataColumns = ('User ID','Date','Day','Month','Year','Time','Hour','Min','Sec','Module','Function','Executed by')
                    
                self.LoggingTreeview = ttk.Treeview(self.FindLoggingFrame,columns=self.LoggingTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.LoggingTreeviewysb = Scrollbar(self.FindLoggingFrame,orient=VERTICAL, command=self.LoggingTreeview.yview)           # <--- Make sure the frame is correct !!
                self.LoggingTreeviewxsb = Scrollbar(self.FindLoggingFrame,orient=HORIZONTAL, command=self.LoggingTreeview.xview)         # <--- Make sure the frame is correct !!
                self.LoggingTreeview['yscroll'] = self.LoggingTreeviewysb.set
                self.LoggingTreeview['xscroll'] = self.LoggingTreeviewxsb.set

                # setup headings and column
                self.LoggingTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.LoggingTreeview.heading('#1', text='User ID',    anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'User ID', False)) # E for East and W for West and CENTER
                self.LoggingTreeview.heading('#2', text='Date',   anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Date', False)) # E for East and W for West and CENTER
                self.LoggingTreeview.heading('#3', text='Day',    anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Day', False)) # E for East and W for West
                self.LoggingTreeview.heading('#4', text='Month',     anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Month', False)) # E for East and W for West
                self.LoggingTreeview.heading('#5', text='Year',     anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Year', False)) 
                self.LoggingTreeview.heading('#6', text='Time', anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Time', False)) 
                self.LoggingTreeview.heading('#7', text='Hour', anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Hour', False)) # E for East and W for West
                self.LoggingTreeview.heading('#8', text='Min',     anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Min', False)) # E for East and W for West
                self.LoggingTreeview.heading('#9', text='Sec',     anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Sec', False)) 
                self.LoggingTreeview.heading('#10', text='Module', anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Module', False)) 
                self.LoggingTreeview.heading('#11', text='Function',      anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Function', False))
                self.LoggingTreeview.heading('#12', text='Executed by',      anchor=W,command=lambda: self.treeview_sort_column(self.LoggingTreeview, 'Executed by', False))
                                   
                self.LoggingTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.LoggingTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.LoggingTreeview.column('#2', stretch=1, width=10)
                self.LoggingTreeview.column('#3', stretch=1, width=10)
                self.LoggingTreeview.column('#4', stretch=1, width=10)
                self.LoggingTreeview.column('#5', stretch=1, width=10)
                self.LoggingTreeview.column('#6', stretch=1, width=10)
                self.LoggingTreeview.column('#7', stretch=1, width=10)
                self.LoggingTreeview.column('#8', stretch=1, width=10)
                self.LoggingTreeview.column('#9', stretch=1, width=10)
                self.LoggingTreeview.column('#10', stretch=1, width=10)
                self.LoggingTreeview.column('#11', stretch=1, width=10)
                self.LoggingTreeview.column('#12', stretch=1, width=10)


                # add tree and scrollbars to frame
                self.LoggingTreeview.grid(row=1, column=0, sticky=NSEW)
                self.LoggingTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.LoggingTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.

                self.LoggingTreeview.tag_configure('Main', font=('Helvetica', 8), background='green2')
                self.LoggingTreeview.tag_configure('Find', font=('Helvetica', 8), background='orange')
                self.LoggingTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.LoggingTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.LoggingTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.LoggingTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                self.LoggingTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.LoggingTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.LoggingTreeview.tag_configure('Closed', font=('Helvetica', 8), background='red3') 
                self.LoggingTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                # Bind the double Click
                #self.LoggingTreeview.bind('<ButtonRelease-1>', self.on_Logging_Tree_select_click) # When Select the Tree
                #self.LoggingTreeview.bind("<Double-1>", self.on_Tree_Logging_double_click)         
                
                #------------------ TREE VIEW For Logging Database <END> -----------------------------------                
                self.FindLoggingWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find User Logging Activity',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   Logging Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Circuits Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindLogging = Class_FindLogging("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindLogging.Find_Logging_Window()


if __name__ == '__main__':
    Main()
