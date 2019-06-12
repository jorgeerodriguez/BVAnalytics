#==========================================================================
# Program: FindCircuitsEquipmentFacilities.py ---- CLASS ---
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
    from Calendar_Jorge import *
    Is_Calendar_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO Calendar Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Calendar_Available = False  

try:
    from Country import *
    Is_Country_Available = True
    Country = Class_Country(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Country Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Country_Available = False  

try:
    from Region import *
    Is_Region_Available = True
    Region = Class_Region(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Region Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Region_Available = False  

try:
    from Facility import *
    Is_Facility_Available = True
    Location = []
    Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO LocalPointOfContacts Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Facility_Available = False  

try:
    from LocalPointOfContacts import *
    Is_LocalPointOfContacts_Available = True
    Location = []
    LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Local Point Of Contacts Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    LocalPointOfContacts = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Find Local Point Of Contatcs','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False


#*********************************************************************************************************************************************
#                                   LocalPointOfContacts Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************
class Class_FindLocalPointOfContacts:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPSummaryWindowExist = False
        self.FindCircuitWindowExist = False
        self.FindDeviceWindowExist = False
        self.CircuitsWindowExist = False
        self.FindLocalPointOfContactsWindowExist = False
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


    def on_FindLocalPointOfContactsWindow_quit(self):
        if (self.FindLocalPointOfContactsWindowExist):
            self.FindLocalPointOfContactsWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Local Point Of Contatcs','CLOSED Window']    
                Logging.Log(Parameter)
            self.FindLocalPointOfContactsWindow.destroy()

    def Call_LocalPointOfContacts_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Local Point Of Contatcs','Facility Button']    
            Logging.Log(Parameter)

        if (Is_LocalPointOfContacts_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_LocalPointOfContacts_ButtonLocalPointOfContacts(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Local Point Of Contatcs','Local Point Of Contatcs Button']    
            Logging.Log(Parameter)

        if (Is_LocalPointOfContacts_Available):
            LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,self.location)
            LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.LocalPointOfContactsTreeview.focus()
        dic = self.LocalPointOfContactsTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.LocalPointOfContactsTreeview.next(curItem)
        self.LocalPointOfContactsTreeview.selection_set((NextItem, NextItem))
        self.LocalPointOfContactsTreeview.focus(NextItem)

        NextItem = self.LocalPointOfContactsTreeview.focus()
        dic = self.LocalPointOfContactsTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
            
    def on_LocalPointOfContacts_Tree_select_click(self,event):
            curItem = self.LocalPointOfContactsTreeview.focus()  
            dic = self.LocalPointOfContactsTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []              
                self.location = [self.LocalPointOfContactsTableCountryIDArray[curItem],self.LocalPointOfContactsTableRegionIDArray[curItem],
                            self.LocalPointOfContactsTableFacilityIDArray[curItem],self.LocalPointOfContactsSites]
                if (Is_LocalPointOfContacts_Available):
                    self.buttonLocalPointOfContacts['state'] = ACTIVE
                if (Is_Facility_Available):
                    self.buttonLocalPointOfContactsFacility['state'] = ACTIVE

    def on_Tree_LocalPointOfContacts_double_click(self,event):
            curItem = self.LocalPointOfContactsTreeview.focus()  
            dic = self.LocalPointOfContactsTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []
                self.location = [self.LocalPointOfContactsTableCountryIDArray[curItem],self.LocalPointOfContactsTableRegionIDArray[curItem],
                            self.LocalPointOfContactsTableFacilityIDArray[curItem],self.LocalPointOfContactsSites]
                if (Is_LocalPointOfContacts_Available):
                    LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,self.location)
                    LocalPointOfContacts.Display_LocalPointOfContacts_Window()


    def Call_ButtonFindLocalPointOfContacts(self):

        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Local Point Of Contatcs','Find Button']    
            Logging.Log(Parameter)

        self.buttonLocalPointOfContacts['state'] = DISABLED
        self.buttonLocalPointOfContactsFacility['state'] = DISABLED
        value = self.FindLocalPointOfContactsFrameEntry.get() # Value to querry
        index = (self.ComboBoxLocalPointOfContactsTableFields.current())
        field = (self.LocalPointOfContactsTableFieldsSQL[index])
        self.LocalPointOfContactsSites = "NO"
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == '*'):
                        sql = """SELECT * FROM Local_Point_Of_Contacts 
                                 ORDER BY Region_ID,Country_ID, Facility_ID
                               """
                    else:
                        sql = """SELECT * FROM Local_Point_Of_Contacts 
                                WHERE %s LIKE '%s' ORDER BY Region_ID,Country_ID, Facility_ID
                              """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Tree View <BEGIN> --------
                        x = self.LocalPointOfContactsTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.LocalPointOfContactsTreeview.delete(child)
                        #------------------------------- Deleting Tree View <END> --------
                        #-------------- Initializing Arrays <BEGIN> ----------------------
                        self.LocalPointOfContactsTablePrimaryKeyArray =  []
                        self.LocalPointOfContactsTableCountryIDArray =  [] 
                        self.LocalPointOfContactsTableRegionIDArray =  []
                        self.LocalPointOfContactsTableFacilityIDArray =  []
                        self.LocalPointOfContactsTableSeqNoArray =  []
                        self.LocalPointOfContactsTablePrimaryPOCNameArray =  []
                        self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray =  []
                        self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray =  []
                        self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray =  []
                        self.LocalPointOfContactsTableSecondaryPOCNameArray =  []
                        self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray =  []
                        self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray =  []
                        self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray =  []
                        self.LocalPointOfContactsTableLastTimeItWasUpdatedArray = []
                        self.LocalPointOfContactsTableExecutedByUserArray = []
                        self.LocalPointOfContactsTableExecutedByUserArray = []
                        self.LocalPointOfContactsTableArrayColumns = []
                        self.LocalPointOfContactsTableArrayTemp = []
                        self.results = []

                        self.LocalPointOfContactsTableArrayColumns = (
                                        'Country ID',
                                        'Region ID',
                                        'Facility ID',
                                        'SEQ No',
                                        'Primary POC Name',
                                        'Primary POC Office No',
                                        'Primary POC Alternate No',
                                        'Primary POC 2nd Alternate No',
                                        'Secondary POC Name',
                                        'Secondary POC Office No',
                                        'Secondary POC Alternate No',
                                        'Secondary POC 2nd Alternate No',
                                        'Last Time it was Updated',
                                        'Executed by UserID'
                                        )
                        
                        #-------------- Initializing Arrays <END> ----------------------
                        self.data_ready = True
                        i = 0
                        while (i < len(self.db.results)):
                            self.LocalPointOfContactsTablePrimaryKeyArray.append(self.db.results[i][0].strip())
                            self.LocalPointOfContactsTableCountryIDArray.append(self.db.results[i][1].strip())
                            self.LocalPointOfContactsTableRegionIDArray.append(self.db.results[i][2].strip())
                            self.LocalPointOfContactsTableFacilityIDArray.append(self.db.results[i][3].strip())
                            self.LocalPointOfContactsTableSeqNoArray.append(self.db.results[i][4])
                            if (self.db.results[i][5] == None):
                                self.LocalPointOfContactsTablePrimaryPOCNameArray.append("")
                            else:
                                self.LocalPointOfContactsTablePrimaryPOCNameArray.append(self.db.results[i][5].strip())
                            if (self.db.results[i][6] == None):
                                self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray.append("")
                            else:
                                self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray.append(self.db.results[i][6].strip())
                            if (self.db.results[i][7] == None):
                                self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray.append("")
                            else: 
                                self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray.append(self.db.results[i][7].strip())
                            if (self.db.results[i][8] == None):
                                self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray.append("")
                            else:
                                self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray.append(self.db.results[i][8].strip())
                            if (self.db.results[i][9] == None):
                                self.LocalPointOfContactsTableSecondaryPOCNameArray.append("")
                            else:
                                self.LocalPointOfContactsTableSecondaryPOCNameArray.append(self.db.results[i][9].strip())
                            if (self.db.results[i][10] == None):
                                self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray.append("")
                            else:
                                self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray.append(self.db.results[i][10].strip())
                            if (self.db.results[i][11] == None):
                                self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray.append("")
                            else:
                                self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray.append(self.db.results[i][11].strip())
                            if (self.db.results[i][12] == None):
                                self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray.append("")
                            else:
                                self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray.append(self.db.results[i][12].strip())
                            self.LocalPointOfContactsTableLastTimeItWasUpdatedArray.append(self.db.results[i][13].strip())
                            self.LocalPointOfContactsTableExecutedByUserArray.append(self.db.results[i][14].strip())
                            i = i + 1
                        i = 0
                        while (i < len(self.LocalPointOfContactsTablePrimaryKeyArray)):
                            num = i + 1
                            tags = "-" # To use in the futire
                            item = [
                                    self.LocalPointOfContactsTableCountryIDArray[i],
                                    self.LocalPointOfContactsTableFacilityIDArray[i],
                                    
                                    self.LocalPointOfContactsTablePrimaryPOCNameArray[i],
                                    self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray[i],
                                    self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray[i],
                                    self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray[i],
                                    self.LocalPointOfContactsTableSecondaryPOCNameArray[i],
                                    self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray[i],
                                    self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray[i],
                                    self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray[i],
                                    self.LocalPointOfContactsTableExecutedByUserArray[i],
                                    self.LocalPointOfContactsTableLastTimeItWasUpdatedArray[i],
                                    ]
                            self.LocalPointOfContactsTableArrayTemp = [
                                self.LocalPointOfContactsTableCountryIDArray[i],
                                self.LocalPointOfContactsTableRegionIDArray[i],
                                self.LocalPointOfContactsTableFacilityIDArray[i],
                                self.LocalPointOfContactsTableSeqNoArray[i],
                                self.LocalPointOfContactsTablePrimaryPOCNameArray[i],
                                self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray[i],
                                self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray[i],
                                self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray[i],
                                self.LocalPointOfContactsTableSecondaryPOCNameArray[i],
                                self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray[i],
                                self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray[i],
                                self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray[i],
                                self.LocalPointOfContactsTableLastTimeItWasUpdatedArray[i],
                                self.LocalPointOfContactsTableExecutedByUserArray[i]
                                ]
                            self.results.append(self.LocalPointOfContactsTableArrayTemp)
                            self.LocalPointOfContactsTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.FindLocalPointOfContactsFrame,title='Find Local Point Of Contacts',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindLocalPointOfContactsFrame,title='Find Local Point Of Contacts',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindLocalPointOfContactsFrame,title='Find Local Point Of Contacts',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindLocalPointOfContacts_From_Entry(self,event):
        self.Call_ButtonFindLocalPointOfContacts()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Find Local Point Of Contacts"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.LocalPointOfContactsTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.FindLocalPointOfContactsFrame,title='Find Local Point Of Contacts List',
                            message = '!!! The File was saved !!!')
                    self.FindLocalPointOfContactsWindow.title("Find Local Point Of Contacts List                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Find Local Point Of Contacts','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindLocalPointOfContactsFrame,title='Find Local Point Of Contacts List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Find_LocalPointOfContacts_Window(self):
        if not self.FindLocalPointOfContactsWindowExist:
            # Set up the Window
            self.FindLocalPointOfContactsWindowExist = True
            self.FindLocalPointOfContactsWindow = Tk()
            self.FindLocalPointOfContactsWindow.geometry('1350x700+350+70')
            self.FindLocalPointOfContactsWindow.title("Find Local Point Of Contacts")
            self.FindLocalPointOfContactsWindow.protocol("WM_DELETE_WINDOW", self.on_FindLocalPointOfContactsWindow_quit)
            self.FindLocalPointOfContactsWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindLocalPointOfContactsLabel = Label(self.FindLocalPointOfContactsWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindLocalPointOfContactsLabel["text"] = "Find Local Point Of Contacts"
            self.FindLocalPointOfContactsLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindLocalPointOfContactsWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindLocalPointOfContactsWindow_quit)
            self.FindLocalPointOfContactsWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.FindLocalPointOfContactsFrame = Frame(self.FindLocalPointOfContactsWindow)
            self.FindLocalPointOfContactsFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindLocalPointOfContactsFrame.rowconfigure(0, weight=1)
            self.FindLocalPointOfContactsFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindLocalPointOfContactsFrameLabel = Label(self.FindLocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindLocalPointOfContactsFrameLabel["text"] = "Find:"
            self.FindLocalPointOfContactsFrameLabel.place(x=10, y=10)

            '''                  
                                Country_ID
                                Region_ID
                                Facility_ID
                                SEQ_No
                                Primary_POC_Name
                                Primary_POC_Office_No
                                Primary_POC_Alternate_No
                                Primary_POC_2nd_Alternate_No
                                Secondary_POC_Name
                                Secondary_POC_Office_No
                                Secondary_POC_Alternate_No
                                Secondary_POC_2nd_Alternate_No
            '''

            # Setup ComboBox
            self.LocalPointOfContactsTableFields = [
                                'Country ID',
                                'Region ID',
                                'Facility ID',
                                #'SEQ_No',
                                'Primary POC Name',
                                'Primary POC Office No',
                                'Primary POC Alternate No',
                                'Primary POC 2nd Alternate No',
                                'Secondary POC Name',
                                'Secondary POC Office No',
                                'Secondary POC Alternate No',
                                'Secondary POC 2nd_Alternate No',
                                'Last Time It Was Updated',
                                ]

            self.LocalPointOfContactsTableFieldsSQL = [
                                'Country_ID',
                                'Region_ID',
                                'Facility_ID',
                                #'SEQ_No',
                                'Primary_POC_Name',
                                'Primary_POC_Office_No',
                                'Primary_POC_Alternate_No',
                                'Primary_POC_2nd_Alternate_No',
                                'Secondary_POC_Name',
                                'Secondary_POC_Office_No',
                                'Secondary_POC_Alternate_No',
                                'Secondary_POC_2nd_Alternate_No',
                                'LastTimeItWasUpdated',                
                                ]

            self.ComboBoxLocalPointOfContactsTableFields = ttk.Combobox(self.FindLocalPointOfContactsFrame, state='readonly', width = 25)
            self.ComboBoxLocalPointOfContactsTableFields['values'] = self.LocalPointOfContactsTableFields
            self.ComboBoxLocalPointOfContactsTableFields.current(0)
            self.ComboBoxLocalPointOfContactsTableFields.place(x = 50, y = 10)

            # Setup Entry
            self.FindLocalPointOfContactsFrameEntry = Entry(self.FindLocalPointOfContactsFrame)
            self.FindLocalPointOfContactsFrameEntry['width']=50
            self.FindLocalPointOfContactsFrameEntry.place(x=235, y=10)
            self.FindLocalPointOfContactsFrameEntry.bind('<Return>', self.Call_ButtonFindLocalPointOfContacts_From_Entry)
            #FindLocalPointOfContactsFrameEntry['state'] = DISABLED

            if self.db.Connect():                
                # Setup Button
                self.buttonFindLocalPointOfContactsExecute = Button(self.FindLocalPointOfContactsFrame, text = 'Execute', command = self.Call_ButtonFindLocalPointOfContacts, state=ACTIVE)
                self.buttonFindLocalPointOfContactsExecute.place(x = 550, y = 7, width=75, height=25)

                self.buttonLocalPointOfContactsFacility = Button(self.FindLocalPointOfContactsFrame, text = 'Facility', command = self.Call_LocalPointOfContacts_ButtonFacility, state=DISABLED)
                self.buttonLocalPointOfContactsFacility.place(x = 850, y = 7, width=75, height=25)

                self.buttonLocalPointOfContacts = Button(self.FindLocalPointOfContactsFrame, text = 'Point of Contacts', command = self.Call_LocalPointOfContacts_ButtonLocalPointOfContacts, state=DISABLED)
                self.buttonLocalPointOfContacts.place(x = 950, y = 7, width=140, height=25)

                    
                #------------------ TREE VIEW For LocalPointOfContacts Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.LocalPointOfContactsTreeviewDataColumns = ('Country ID','Facility ID','Primay POC','Office No','Alternate','2nd Alternate',
                                                                    'Secondary POC','Office No','Alternate','2nd Alternate','Added By','Date & Time')
                    
                self.LocalPointOfContactsTreeview = ttk.Treeview(self.FindLocalPointOfContactsFrame,columns=self.LocalPointOfContactsTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.LocalPointOfContactsTreeviewysb = Scrollbar(self.FindLocalPointOfContactsFrame,orient=VERTICAL, command=self.LocalPointOfContactsTreeview.yview)           # <--- Make sure the frame is correct !!
                self.LocalPointOfContactsTreeviewxsb = Scrollbar(self.FindLocalPointOfContactsFrame,orient=HORIZONTAL, command=self.LocalPointOfContactsTreeview.xview)         # <--- Make sure the frame is correct !!
                self.LocalPointOfContactsTreeview['yscroll'] = self.LocalPointOfContactsTreeviewysb.set
                self.LocalPointOfContactsTreeview['xscroll'] = self.LocalPointOfContactsTreeviewxsb.set

                # setup headings and column
                self.LocalPointOfContactsTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.LocalPointOfContactsTreeview.heading('#1', text='Country ID',    anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Country ID', False)) # E for East and W for West and CENTER
                self.LocalPointOfContactsTreeview.heading('#2', text='Facility ID',   anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Facility ID', False)) # E for East and W for West and CENTER
                self.LocalPointOfContactsTreeview.heading('#3', text='Primay POC',    anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Primay POC', False)) # E for East and W for West
                self.LocalPointOfContactsTreeview.heading('#4', text='Office No',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Office No', False)) # E for East and W for West
                self.LocalPointOfContactsTreeview.heading('#5', text='Alternate',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Alternate', False)) 
                self.LocalPointOfContactsTreeview.heading('#6', text='2nd Alternate', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, '2nd Alternate', False)) 
                self.LocalPointOfContactsTreeview.heading('#7', text='Secondary POC', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Secondary POC', False)) # E for East and W for West
                self.LocalPointOfContactsTreeview.heading('#8', text='Office No',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Office No', False)) # E for East and W for West
                self.LocalPointOfContactsTreeview.heading('#9', text='Alternate',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Alternate', False)) 
                self.LocalPointOfContactsTreeview.heading('#10', text='2nd Alternate', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, '2nd Alternate', False)) 
                self.LocalPointOfContactsTreeview.heading('#11', text='Added By',      anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Added By', False))
                self.LocalPointOfContactsTreeview.heading('#12', text='Date & Time',      anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Date & Time', False))
                                   
                self.LocalPointOfContactsTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.LocalPointOfContactsTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.LocalPointOfContactsTreeview.column('#2', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#3', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#4', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#5', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#6', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#7', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#8', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#9', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#10', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#11', stretch=1, width=10)
                self.LocalPointOfContactsTreeview.column('#12', stretch=1, width=10)


                # add tree and scrollbars to frame
                self.LocalPointOfContactsTreeview.grid(row=1, column=0, sticky=NSEW)
                self.LocalPointOfContactsTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.LocalPointOfContactsTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.

                self.LocalPointOfContactsTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.LocalPointOfContactsTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                self.LocalPointOfContactsTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.LocalPointOfContactsTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.LocalPointOfContactsTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.LocalPointOfContactsTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                self.LocalPointOfContactsTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.LocalPointOfContactsTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.LocalPointOfContactsTreeview.tag_configure('Closed', font=('Helvetica', 8), background='red3') 
                self.LocalPointOfContactsTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                # Bind the double Click
                self.LocalPointOfContactsTreeview.bind('<ButtonRelease-1>', self.on_LocalPointOfContacts_Tree_select_click) # When Select the Tree
                self.LocalPointOfContactsTreeview.bind("<Double-1>", self.on_Tree_LocalPointOfContacts_double_click)         
                
                #------------------ TREE VIEW For LocalPointOfContacts Database <END> -----------------------------------                
                self.FindLocalPointOfContactsWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find Local Point Of Contacts',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   LocalPointOfContacts Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Circuits Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindLPOC = Class_FindLocalPointOfContacts("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindLPOC.Find_LocalPointOfContacts_Window()


if __name__ == '__main__':
    Main()
