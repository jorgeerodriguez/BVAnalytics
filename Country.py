#==========================================================================
# Program: Country.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-22-2017
# Date Last Modified: Dec-23-2017
# Summary: This is Class to for the Country
#==========================================================================

#***************************************************************
# ==================== Libraries Required <BEGIN> =============*
#***************************************************************

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

#-------------------------- Validating the Variables Values <BEGIN> ----------------
#print (ODBC_DSN_name)
#print (DatabaseVersion)
#print (BVAnalyticsDatabaseVersion)
#print (Windows_Scaling)
#-------------------------- Validating the Variables Values <END> ----------------

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Country','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False


#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_Country:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.CountryWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.version = self.db.Get_Version()

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


    def on_CountryWindow_quit(self):
        if (self.CountryWindowExist):
            self.CountryWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Country','CLOSE Window']    
                Logging.Log(Parameter) 
            self.CountryWindow.destroy()

    def on_Country_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.CountryTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.CountryTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

            # SQL Querry to the COUNTRY Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.results = []
                i = 0
                self.progress['maximum'] = len(self.db.results)
                while (i < len(self.db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [self.db.results[i][0],
                            self.db.results[i][1],
                            self.db.results[i][2]]
                    self.results.append(item)
                    self.CountryTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                    self.progress['value'] = i
            else:
                mbox.showerror(master=self.CountryFrame,title='Country',
                message = 'No Records found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.CountryFrame,title='Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_Country_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.CountryIDFrameEntry['state'] = 'normal'
        self.CountryIDFrameEntry.delete(0,END)
        self.CountryIDFrameEntry['state'] = 'readonly'
        self.CountryIDFrameEntry['state'] = 'normal'
            
        self.CountryNameFrameEntry['state'] = 'normal'
        self.CountryNameFrameEntry.delete(0,END)
        self.CountryNameFrameEntry['state'] = 'readonly'
        self.CountryNameFrameEntry['state'] = 'normal'

        CountryID = self.CountryIDFrameEntry.get()
        CountryName = self.CountryNameFrameEntry.get()

        self.ButtonCountryAdd['state']    = DISABLED
        self.ButtonCountryEdit['state']   = DISABLED
        self.ButtonCountryRemove['state'] = DISABLED
        self.ButtonCountryOK['state']     = ACTIVE
        self.ButtonCountryCancel['state'] = ACTIVE



    def Call_Button_Country_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.CountryIDFrameEntry['state'] = 'readonly'
        self.CountryNameFrameEntry['state'] = 'normal'

        self.ButtonCountryAdd['state']    = DISABLED
        self.ButtonCountryEdit['state']   = DISABLED
        self.ButtonCountryRemove['state'] = DISABLED
        self.ButtonCountryOK['state']     = ACTIVE
        self.ButtonCountryCancel['state'] = ACTIVE


    def Call_Button_Country_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.CountryFrame,title='Country',message = 'Are you Sure you want to Remove it?')):
                self.ButtonCountryAdd['state']    = ACTIVE
                self.ButtonCountryEdit['state']   = DISABLED
                self.ButtonCountryRemove['state'] = DISABLED
                self.ButtonCountryOK['state']     = DISABLED
                self.ButtonCountryCancel['state'] = DISABLED                
                CountryID = self.CountryIDFrameEntry.get()
                CountryName = self.CountryNameFrameEntry.get()
                if (CountryID == "UNKNOWN"):
                    mbox.showerror(master=self.CountryFrame,title='Country',
                        message = '*** The Country ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM REGION
                            WHERE Country_ID = '%s'
                          """ % (CountryID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CountryFrame,title='Country',
                                    message = '*** The Country ID you entered CANNOT be Removed, \r\n Becuase is used in a Region ***')
                            self.on_Country_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM COUNTRY
                                WHERE Country_ID = '%s'
                              """ % (CountryID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM COUNTRY WHERE Country_ID = '%s'" % (CountryID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.CountryFrame,title='Country',
                                        message = '*** The Country ID you entered was Removed ***')
                                self.on_Country_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.CountryFrame,title='Country',
                                        message = '*** The Country ID you entered was NOT Removed ***')
                                self.on_Country_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.CountryFrame,title='Country',
                            message = '*** The Country ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CountryFrame,title='Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Country_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonCountryAdd['state']    = DISABLED
            self.ButtonCountryEdit['state']   = DISABLED
            self.ButtonCountryRemove['state'] = DISABLED
            self.ButtonCountryOK['state']     = DISABLED #
            self.ButtonCountryCancel['state'] = DISABLED #

            CountryID = self.CountryIDFrameEntry.get()
            CountryName = self.CountryNameFrameEntry.get()
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(CountryID) > 0) and (len(CountryName) > 0)):
                    sql = """
                            SELECT * FROM COUNTRY
                            WHERE Country_ID = '%s'
                          """ % (CountryID)
                    if (self.db.Execute(sql)):
                        self.ButtonCountryOK['state']     = ACTIVE #
                        self.ButtonCountryCancel['state'] = ACTIVE #                        
                        mbox.showerror(master=self.CountryFrame,title='Country',
                        message = '*** The Country ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO COUNTRY(Country_ID, \
                                   Country_Name, Executed_by_UserID) \
                                   VALUES ('%s', '%s', '%s')" % \
                                   (CountryID, CountryName, self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.CountryFrame,title='Country',
                                message = '*** The Country ID you entered was Added ***')
                            self.on_Country_Table_Refresh()
                            self.CountryIDFrameEntry['state'] = 'readonly'
                            self.CountryNameFrameEntry['state'] = 'readonly'                            
                        else:
                            #print ("Error adding the record, posible dupliated it")
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CountryFrame,title='Country',
                                message = '*** The Country ID you entered was NOT Added ***')
                            self.on_Country_Table_Refresh()
                            self.CountryIDFrameEntry['state'] = 'readonly'
                            self.CountryNameFrameEntry['state'] = 'readonly'
                else:
                    self.ButtonCountryOK['state']     = ACTIVE #
                    self.ButtonCountryCancel['state'] = ACTIVE #
                    mbox.showerror(master=self.CountryFrame,title='Country',
                        message = '*** The Country ID and Country Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM COUNTRY
                        WHERE Country_ID = '%s'
                      """ % (CountryID)
                if (len(CountryName) == 0):
                     mbox.showerror(master=self.CountryFrame,title='Country',
                        message = '*** The Country Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):

                        sql = "UPDATE COUNTRY SET Country_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID = '%s'" % (CountryName,self.Username,CountryID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.CountryFrame,title='Country',
                                message = '*** The Country ID you entered was Updated ***')
                            self.on_Country_Table_Refresh()
                            self.CountryNameFrameEntry['state'] = 'readonly'
                            
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CountryFrame,title='Country',
                                message = '*** The Country ID you entered was NOT Upadted ***')
                            self.on_Country_Table_Refresh()               
                            self.CountryNameFrameEntry['state'] = 'readonly'
                    else:
                        mbox.showerror(master=self.CountryFrame,title='Country',
                        message = '*** The Country ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CountryFrame,title='Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Country_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonCountryAdd['state']    = ACTIVE
        self.ButtonCountryEdit['state']   = DISABLED
        self.ButtonCountryRemove['state'] = DISABLED
        self.ButtonCountryOK['state']     = DISABLED
        self.ButtonCountryCancel['state'] = DISABLED

        self.CountryIDFrameEntry['state'] = 'normal'
        self.CountryIDFrameEntry.delete(0,END)
        self.CountryIDFrameEntry['state'] = 'readonly'
            
        self.CountryNameFrameEntry['state'] = 'normal'
        self.CountryNameFrameEntry.delete(0,END)
        self.CountryNameFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_Country_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.CountryTreeview.focus()
        dic = self.CountryTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            #print (values)
            #print (values[0])
            #print (values[1])

            self.ButtonCountryAdd['state']    = ACTIVE
            self.ButtonCountryEdit['state']   = ACTIVE
            self.ButtonCountryRemove['state'] = ACTIVE
            self.ButtonCountryOK['state']     = DISABLED
            self.ButtonCountryCancel['state'] = DISABLED

            self.CountryIDFrameEntry['state'] = 'normal'
            self.CountryIDFrameEntry.delete(0,END)
            self.CountryIDFrameEntry.insert(0,str(values[0]).strip())
            self.CountryIDFrameEntry['state'] = 'readonly'
            
            self.CountryNameFrameEntry['state'] = 'normal'
            self.CountryNameFrameEntry.delete(0,END)
            self.CountryNameFrameEntry.insert(0,str(values[1]).strip())
            self.CountryNameFrameEntry['state'] = 'readonly'

    def On_Country_Tree_Refresh(self,event):
        #--- Double Click --
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Country','REFRESH Button']    
            Logging.Log(Parameter)
            
        self.on_Country_Table_Refresh()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Country List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.CountryTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = self.results[i]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    #Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.CountryFrame,title='Country',
                            message = '!!! The File was saved !!!')
                    self.CountryWindow.title("Country                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Country','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.CountryFrame,title='Country',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Display_Country_Window(self):
        global CountryLabel
        if not self.CountryWindowExist:
            self.CountryWindowExist = True
            # Set up the Window
            self.CountryWindow = Tk()
            self.CountryWindow.geometry('1000x400+350+70')
            self.CountryWindow.title("Country")
            self.CountryWindow.protocol("WM_DELETE_WINDOW", self.on_CountryWindow_quit)
            self.CountryWindow.call('tk', 'scaling', self.Windows_Scaling)
            CountryLabel = Label(self.CountryWindow,text="Helvetica", font=("Helvetica", 19))
            CountryLabel["text"] = "Country"
            CountryLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.CountryWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_CountryWindow_quit)
            self.CountryWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.CountryFrame = Frame(self.CountryWindow)
            self.CountryFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.CountryFrame.rowconfigure(0, weight=1)
            self.CountryFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    CountryIDFrameLabel = Label(self.CountryFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country ID:"
                    CountryIDFrameLabel.place(x=10, y=10)
                    
                    CountryNameFrameLabel = Label(self.CountryFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryNameFrameLabel["text"] = "Country Name:"
                    CountryNameFrameLabel.place(x=10, y=35)

                   # Setup Entry
                    self.CountryIDFrameEntry = Entry(self.CountryFrame)
                    self.CountryIDFrameEntry['width']=50
                    self.CountryIDFrameEntry.place(x=110, y=10)
                    self.CountryIDFrameEntry['state'] = DISABLED

                    self.CountryNameFrameEntry = Entry(self.CountryFrame)
                    self.CountryNameFrameEntry['width']=50
                    self.CountryNameFrameEntry.place(x=110, y=35)
                    self.CountryNameFrameEntry['state'] = DISABLED



                    '''
                    # Setup ComboBox
                    DeviceTableFields = ['Device Name',
                                    'Description',
                                    'Country',
                                    'Region',
                                    'Facility',
                                    'Site',
                                    'Device Type',
                                    'IP Address',
                                    'Status']
                    DeviceTableFieldsSQL = ['Device_ID',
                                    'Device_Description',
                                    'Country_ID',
                                    'Region_ID',
                                    'Facility_ID',
                                    'Site_ID',
                                    'Device_Type_ID',
                                    'IP4_Address',
                                    'Status']

                    ComboBoxICMPTableFields = ttk.Combobox(ICMPFrame, state='readonly', width = 25)
                    ComboBoxICMPTableFields['values'] = DeviceTableFields
                    ComboBoxICMPTableFields.current(0)
                    ComboBoxICMPTableFields.place(x = 50, y = 10)
                    #ComboBoxButtonWEBOSName['state'] = DISABLED
                    '''

                    
                    
                    
                    # Setup Buttons
                    self.ButtonCountryAdd = Button(self.CountryFrame, text = 'Add', command = self.Call_Button_Country_Add, state=DISABLED)
                    self.ButtonCountryAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonCountryAdd['state'] = ACTIVE

                    self.ButtonCountryEdit = Button(self.CountryFrame, text = 'Edit', command = self.Call_Button_Country_Edit, state=DISABLED)
                    self.ButtonCountryEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonCountryRemove = Button(self.CountryFrame, text = 'Remove', command = self.Call_Button_Country_Remove, state=DISABLED)
                    self.ButtonCountryRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonCountryOK = Button(self.CountryFrame, text = 'OK / UPDATE', command = self.Call_Button_Country_OK, state=DISABLED)
                    self.ButtonCountryOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonCountryCancel = Button(self.CountryFrame, text = 'Cancel', command = self.Call_Button_Country_Cancel, state=DISABLED)
                    self.ButtonCountryCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.CountryFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=35)


                    #------------------ TREE VIEW For COUNTRY Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.CountryTreeviewDataColumns = ('Country ID','Country Name','Added By')
                    
                    self.CountryTreeview = ttk.Treeview(self.CountryFrame,columns=self.CountryTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.CountryTreeviewysb = Scrollbar(self.CountryFrame,orient=VERTICAL, command=self.CountryTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.CountryTreeviewxsb = Scrollbar(self.CountryFrame,orient=HORIZONTAL, command=self.CountryTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.CountryTreeview['yscroll'] = self.CountryTreeviewysb.set
                    self.CountryTreeview['xscroll'] = self.CountryTreeviewxsb.set

                    # setup headings and column 
                    self.CountryTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.CountryTreeview.heading('#1', text='Country ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CountryTreeview, 'Country ID', False)) # E for East and W for West
                    self.CountryTreeview.heading('#2', text='Country Name', anchor=W,command=lambda: self.treeview_sort_column(self.CountryTreeview, 'Country Name', False)) 
                    self.CountryTreeview.heading('#3', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.CountryTreeview, 'Added By', False)) 
                    
                    self.CountryTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.CountryTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.CountryTreeview.column('#2', stretch=1, width=10)
                    self.CountryTreeview.column('#3', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.CountryTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.CountryTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.CountryTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.CountryTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.CountryTreeview.bind('<ButtonRelease-1>', self.on_Country_Tree_select_click) # When Select the Tree
                    self.CountryTreeview.bind("<Double-1>", self.On_Country_Tree_Refresh)         

                    # SQL Querry to the COUNTRY Table
                    sql = """
                                SELECT * FROM COUNTRY
                                ORDER BY Country_Name ASC
                              """
                    if (self.db.Execute(sql)):
                        i = 0
                        self.data_ready = True
                        self.results = []
                        while (i < len(self.db.results)):
                            num = i + 1
                            tags = "-" # To use in the futire                          
                            item = [self.db.results[i][0],
                                    self.db.results[i][1],
                                    self.db.results[i][2]]
                            self.results.append(item)
                            self.CountryTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.CountryFrame,title='Country',
                        message = 'No Records found')
                    #------------------ TREE VIEW For COUNTRY Database <END> -----------------------------------
                    #self.db.Disconnect()
                    self.CountryWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.CountryFrame,title='Country',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Country Class....:")
    Country = Class_Country("BV",Windows_Scaling)
    Country.Display_Country_Window()


if __name__ == '__main__':
    Main()

