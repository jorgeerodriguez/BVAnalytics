#==========================================================================
# Program: DeviceModelCountry.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-23-2017
# Date Last Modified: Dec-29-2017
# Summary: This is Class to for the DeviceModelCountry
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
    from Country import *
    Is_Country_Available = True
    Country = Class_Country(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Country Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Country_Available = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Device Model Country','OPEN Window']    
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

class Class_DeviceModelCountry:

    def __init__(self,DSN_Name,Windows_Scaling,Parameters,Treeview):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.DeviceModelCountryWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.DeviceModelCountryCostTreeview = Treeview
        if (len(Parameters) > 0 ):
            #Parameters = ["AVPN","AVPN",10.1,"USA","EDIT",""]
            self.Device_Model_ID_Parameter = Parameters[0]
            self.Device_Type_Name_Parameter = Parameters[1]
            self.Device_Type_Outsource_Parameter = Parameters[2]
            self.Device_Type_Country_ID_Parameter = Parameters[3]
            self.Device_Type_Action_Parameter = Parameters[4] # Add or Edit
            self.Device_Type_Button_Parameter = Parameters[5] # OK or Cancel
            
            
        
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

#---------------------------- Country SELECTION SECTION <BEGIN> ------------------------
        
    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model Country','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_Country_combo_changed(self,event):
        self.DeviceModelCountryIDFrameEntry['state'] = 'normal'
        self.DeviceModelCountryIDFrameEntry.delete(0,END)
        self.DeviceModelCountryIDFrameEntry.insert(0,self.Device_Model_ID_Parameter)
        self.DeviceModelCountryIDFrameEntry['state'] = 'readonly'
            
        self.DeviceModelCountryNameFrameEntry['state'] = 'normal'
        self.DeviceModelCountryNameFrameEntry.delete(0,END)
        self.DeviceModelCountryNameFrameEntry.insert(0,self.Device_Type_Name_Parameter)
        self.DeviceModelCountryNameFrameEntry['state'] = 'readonly'

        self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
        self.DeviceModelCountryCostFrameEntry.delete(0,END)
        self.DeviceModelCountryCostFrameEntry.insert(0,str(self.Device_Type_Outsource_Parameter))
        self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
            
        if self.db.Connect():
            # SQL Querry to the DeviceModelCountry Table  
            
            Primarykey = self.Device_Model_ID_Parameter + "-" + self.CountryIDComboBox[self.ComboBoxCountryID.current()]
            #print (Primarykey)

            sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID_Country_ID = '%s'
                  """ % (Primarykey)
            if (self.db.Execute(sql)):
                self.Device_Type_Outsource_Parameter = self.db.results[0][3]
            else:
                self.Device_Type_Outsource_Parameter = 0.0

            self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
            self.DeviceModelCountryCostFrameEntry.delete(0,END)
            self.DeviceModelCountryCostFrameEntry.insert(0,str(self.Device_Type_Outsource_Parameter))
            self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
                
            sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID = '%s'
                    ORDER BY Country_ID
                  """ % (self.Device_Model_ID_Parameter)
            
            if (self.db.Execute(sql)):
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.DeviceModelCountryCostTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.DeviceModelCountryCostTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                self.DeviceModelCountryTablePKArray = []
                self.DeviceModelCountryTableDeviceModelIDArray= []
                self.DeviceModelCountryTableCountryIDArray = []
                self.DeviceModelCountryTableOutsourceCostArray = []
                self.DeviceModelCountryTableExecutedbyArray = []
                        
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceModelCountryTablePKArray.append(self.db.results[i][0].strip())
                    self.DeviceModelCountryTableDeviceModelIDArray.append(self.db.results[i][1].strip())
                    self.DeviceModelCountryTableCountryIDArray.append(self.db.results[i][2].strip())
                    self.DeviceModelCountryTableOutsourceCostArray.append(self.db.results[i][3])
                    self.DeviceModelCountryTableExecutedbyArray.append(self.db.results[i][4].strip())
                    i = i + 1
                   
                i = 0
                while (i < len(self.DeviceModelCountryTablePKArray)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.DeviceModelCountryTableCountryIDArray[i],
                            self.DeviceModelCountryTableDeviceModelIDArray[i],
                            self.DeviceModelCountryTableOutsourceCostArray[i],
                            self.DeviceModelCountryTableExecutedbyArray[i]
                            ]
                    self.DeviceModelCountryCostTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1

                self.ButtonDeviceModelCountryRemove['state'] = ACTIVE
                self.ButtonDeviceModelCountryOK['state']     = ACTIVE
                self.ButtonDeviceModelCountryCancel['state'] = ACTIVE
            else:
                mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                message = 'No Records found')

                self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
                self.DeviceModelCountryCostFrameEntry.delete(0,END)
                self.DeviceModelCountryCostFrameEntry.insert(0,str(0.0))
                self.DeviceModelCountryCostFrameEntry['state'] = 'normal'

                self.ButtonDeviceModelCountryRemove['state'] = DISABLED
                self.ButtonDeviceModelCountryOK['state']     = ACTIVE
                self.ButtonDeviceModelCountryCancel['state'] = ACTIVE
                
        else:
            mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model Country','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDComboBox = []
            self.CountryNameComboBox = []

            # SQL Querry to the DeviceModelCountry Table
            sql = """ SELECT * FROM Country ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CountryIDComboBox.append(self.db.results[i][0].strip())
                    self.CountryNameComboBox.append(self.db.results[i][1].strip())
                    i = i + 1
                self.ComboBoxCountryID['values'] = self.CountryNameComboBox
                if (len(self.CountryNameComboBox)== 0):
                    self.ComboBoxCountryID['state'] = DISABLED
                else:
                    if (len(self.Device_Type_Country_ID_Parameter) == 0):
                        self.ComboBoxCountryID.current(0)
                    else:
                        i = 0
                        while (i < len (self.CountryIDComboBox)):
                            if (self.CountryIDComboBox[i] == self.Device_Type_Country_ID_Parameter):
                                self.ComboBoxCountryID.current(i)
                                i = i + len (self.CountryIDComboBox)
                            i = i + 1
                    self.on_DeviceModelCountry_Table_Refresh()
            else:
                mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                message = 'No Device Type Records found')
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        

#---------------------------- Country SELECTION SECTION <END> ------------------------

        
    def on_DeviceModelCountryWindow_quit(self):
        if (self.DeviceModelCountryWindowExist):
            self.DeviceModelCountryWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Device Model Country','CLOSE Window']    
                Logging.Log(Parameter)
            self.DeviceModelCountryWindow.destroy()


    def on_DeviceModelCountry_Table_Refresh(self):
        self.on_Country_combo_changed("event")
       
    def Call_Button_DeviceModelCountry_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model Country','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():            
            if (mbox.askyesnocancel(master=self.DeviceModelCountryFrame,title='Device Model Country',message = 'Are you Sure you want to Remove '+
                                    self.CountryNameComboBox[self.ComboBoxCountryID.current()] +'?')):

                Primarykey = self.Device_Model_ID_Parameter + "-" + self.CountryIDComboBox[self.ComboBoxCountryID.current()]
                sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID_Country_ID = '%s'
                  """ % (Primarykey)
                if (self.db.Execute(sql)):
                    sql = "DELETE FROM Device_Model_Country WHERE Device_Model_ID_Country_ID = '%s'" % (Primarykey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        mbox.showwarning(master=self.DeviceModelCountryFrame,title='Device Model Country',
                                        message = '*** The Cost per Country of the Device Model ID you entered was Removed ***')
                    else:
                        #self.db.Disconnect()
                        mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                                        message = '*** The Cost per Country of the Device Model ID you entered was NOT Removed ***')
                    self.on_DeviceModelCountry_Table_Refresh()               
                else:
                    mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                    message = '*** The Cost per Country of the Device Model ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceModelCountryFrame,title='DeviceModelCountry',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceModelCountry_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model Country','ADD/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ComboBoxCountryID['state'] = 'readonly'

            self.ButtonDeviceModelCountryRemove['state'] = DISABLED
            self.ButtonDeviceModelCountryOK['state']     = DISABLED 
            self.ButtonDeviceModelCountryCancel['state'] = DISABLED

            Primarykey = self.Device_Model_ID_Parameter + "-" + self.CountryIDComboBox[self.ComboBoxCountryID.current()]
            #print (Primarykey)
    
            sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID_Country_ID = '%s'
                  """ % (Primarykey)
            
            DeviceModelCountryCountryID = self.CountryIDComboBox[self.ComboBoxCountryID.current()]
            DeviceModelCountryDeviceModelID = self.Device_Model_ID_Parameter
            DeviceModelCountryDeviceCost = self.DeviceModelCountryCostFrameEntry.get()
            if (len(DeviceModelCountryDeviceCost) > 0):
                DeviceModelCountryDeviceCost = float(self.DeviceModelCountryCostFrameEntry.get())
            else:
                DeviceModelCountryDeviceCost = 0.0
            '''
                Device_Model_ID_Country_ID      CHAR(60) NOT NULL PRIMARY KEY,
                Device_Model_ID                 CHAR(30) NOT NULL,
                Country_ID                      CHAR(20) NOT NULL,
                Outsource_Cost                  FLOAT,
                Executed_by_UserID              CHAR(20))"""
            '''
            if (self.db.Execute(sql)):
                #-------------- EDIT ----------------------
                sql = "UPDATE Device_Model_Country SET Outsource_Cost = '%f', \
                               Executed_by_UserID = '%s' WHERE Device_Model_ID_Country_ID = '%s'" % (DeviceModelCountryDeviceCost,
                                                                                                    self.Username,Primarykey)
                if (self.db.Add_Move_Change_Data(sql)):
                    mbox.showwarning(master=self.DeviceModelCountryFrame,title='Device Model Country',
                        message = '*** The Device Model Outsource Cost per Country you entered was Updated ***')
                else:
                    mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                        message = '*** The Device Model Outsource Cost per Country you entered was NOT Upadted ***')
            else:
                #-------------- ADD ----------------------
                sql = "INSERT INTO Device_Model_Country(Device_Model_ID_Country_ID, Device_Model_ID, Country_ID, \
                                   Outsource_Cost, Executed_by_UserID) \
                                   VALUES ('%s','%s','%s', '%f','%s')" % (Primarykey, self.Device_Model_ID_Parameter,
                                                                          DeviceModelCountryCountryID, DeviceModelCountryDeviceCost, self.Username)
                if (self.db.Add_Move_Change_Data(sql)):
                    mbox.showwarning(master=self.DeviceModelCountryFrame,title='Device Model Country',
                        message = '*** The Device Model Outsource Cost per Country you entered was Added ***')
                else:
                    #print ("Error adding the record, posible dupliated it")
                    #self.db.Disconnect()
                    mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                        message = '*** The Device Model ID you entered was NOT Added ***')                    
        else:
            mbox.showerror(master=self.DeviceModelCountryFrame,title='Device Model Country',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        self.DeviceModelCountryIDFrameEntry['state'] = 'readonly'                           
        self.DeviceModelCountryNameFrameEntry['state'] = 'readonly'
        self.DeviceModelCountryCostFrameEntry['state'] = 'normal'
        self.on_DeviceModelCountry_Table_Refresh()


    def Display_DeviceModelCountry_Window(self):
        #global DeviceModelCountryLabel
 
        if not self.DeviceModelCountryWindowExist:
            # Set up the Window
            self.DeviceModelCountryWindowExist = True
            self.DeviceModelCountryWindow = Tk()
            self.DeviceModelCountryWindow.geometry('750x200+350+70')
            self.DeviceModelCountryWindow.title("Device Model Country")
            self.DeviceModelCountryWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceModelCountryWindow_quit)
            self.DeviceModelCountryWindow.call('tk', 'scaling', self.Windows_Scaling)
            DeviceModelCountryLabel = Label(self.DeviceModelCountryWindow,text="Helvetica", font=("Helvetica", 19))
            DeviceModelCountryLabel["text"] = "Device Model Outsource Cost Per Country"
            DeviceModelCountryLabel.pack()
                    
            # Setup Frame
            self.DeviceModelCountryFrame = Frame(self.DeviceModelCountryWindow)
            self.DeviceModelCountryFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.DeviceModelCountryFrame.rowconfigure(0, weight=1)
            self.DeviceModelCountryFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.DeviceModelCountryFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 420, y = 7, width=100, height=25)

                        self.ButtonCountryRefresh = Button(self.DeviceModelCountryFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 535, y = 7, width=75, height=25)

                    self.ButtonDeviceModelCountryOK = Button(self.DeviceModelCountryFrame, text = 'ADD / UPDATE', command = self.Call_Button_DeviceModelCountry_OK, state=DISABLED)
                    self.ButtonDeviceModelCountryOK.place(x = 420, y = 37, width=100, height=25)

                    self.ButtonDeviceModelCountryRemove = Button(self.DeviceModelCountryFrame, text = 'Remove', command = self.Call_Button_DeviceModelCountry_Remove, state=DISABLED)
                    self.ButtonDeviceModelCountryRemove.place(x = 535, y = 37, width=75, height=25)

                    self.ButtonDeviceModelCountryCancel = Button(self.DeviceModelCountryFrame, text = 'CLOSE', command = self.on_DeviceModelCountryWindow_quit, state=DISABLED)
                    self.ButtonDeviceModelCountryCancel.place(x = 625, y = 37, width=75, height=25)


                    # Setup Labels
                    CountryIDFrameLabel = Label(self.DeviceModelCountryFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country:"
                    CountryIDFrameLabel.place(x=10, y=10)
                    
                    DeviceModelCountryIDFrameLabel = Label(self.DeviceModelCountryFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelCountryIDFrameLabel["text"] = "Device Model ID:"
                    DeviceModelCountryIDFrameLabel.place(x=10, y=40)
                    
                    DeviceModelCountryNameFrameLabel = Label(self.DeviceModelCountryFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelCountryNameFrameLabel["text"] = "Device Model Name:"
                    DeviceModelCountryNameFrameLabel.place(x=10, y=70)

                    DeviceModelCountryNameFrameLabel = Label(self.DeviceModelCountryFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelCountryNameFrameLabel["text"] = "Monthly Outsource Cost:"
                    DeviceModelCountryNameFrameLabel.place(x=10, y=100)

                   # Setup Entry
                    self.DeviceModelCountryIDFrameEntry = Entry(self.DeviceModelCountryFrame)
                    self.DeviceModelCountryIDFrameEntry['width']=40
                    self.DeviceModelCountryIDFrameEntry.place(x=165, y=40)
                    self.DeviceModelCountryIDFrameEntry['state'] = 'readonly'

                    self.DeviceModelCountryNameFrameEntry = Entry(self.DeviceModelCountryFrame)
                    self.DeviceModelCountryNameFrameEntry['width']=40
                    self.DeviceModelCountryNameFrameEntry.place(x=165, y=70)
                    self.DeviceModelCountryNameFrameEntry['state'] = 'readonly'              

                    self.DeviceModelCountryCostFrameEntry = Entry(self.DeviceModelCountryFrame)
                    self.DeviceModelCountryCostFrameEntry['width']=20
                    self.DeviceModelCountryCostFrameEntry.place(x=165, y=100)
                    self.DeviceModelCountryCostFrameEntry['state'] = 'normal'              


                    # Setup ComboBox
                    self.ComboBoxCountryID = ttk.Combobox(self.DeviceModelCountryFrame, state='readonly', width = 37)
                    self.ComboBoxCountryID.bind("<<ComboboxSelected>>", self.on_Country_combo_changed)
                    self.ComboBoxCountryID.place(x = 165, y = 10)
                    self.on_Country_Table_Refresh()           
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.DeviceModelCountryWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.DeviceModelCountryFrame,title='DeviceModelCountry',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")


################

def on_DeviceModelCountryWindow_quit():
    global DeviceModelCountry
    global DeviceModelWindow

    DeviceModelCountry.on_DeviceModelCountryWindow_quit()
    DeviceModelWindow.destroy()
    

                
def Display_DeviceModel_Window():
    #global DeviceModelCountryLabel
    global DeviceModelCountryCostTreeview
    global DeviceModelWindowExist
    global db
    global DeviceModelCountry
    global DeviceModelWindow
 
    if not DeviceModelWindowExist:
        # Set up the Window
        DeviceModelWindowExist = True
        DeviceModelWindow = Tk()
        DeviceModelWindow.geometry('1000x300+350+70')
        DeviceModelWindow.title("DeviceModelCountry")
        DeviceModelWindow.protocol("WM_DELETE_WINDOW", on_DeviceModelCountryWindow_quit)
        DeviceModelWindow.call('tk', 'scaling', "1.3")
        DeviceModelLabel = Label(DeviceModelWindow,text="Helvetica", font=("Helvetica", 19))
        DeviceModelLabel["text"] = "Device Model"
        DeviceModelLabel.pack()
                    
        # Setup Frame
        DeviceModelFrame = Frame(DeviceModelWindow)
        DeviceModelFrame.pack(side=TOP, fill=BOTH, expand=Y)

        # set frame resizing priorities
        DeviceModelFrame.rowconfigure(0, weight=1)
        DeviceModelFrame.columnconfigure(0, weight=1)

        if db.Connect():

            #------------------ TREE VIEW For Device Model Country Cost Database <BEGIN> -----------------------------------
            # Create Tree and Scrollbars
            DeviceModelCountryCostTreeviewDataColumns = ('Country ID','Device Model ID','Outsource Cost','Added By')
                    
            DeviceModelCountryCostTreeview = ttk.Treeview(DeviceModelFrame,columns=DeviceModelCountryCostTreeviewDataColumns, height=10) # <--- Make sure the frame is correct !!
            DeviceModelCountryCostTreeviewysb = Scrollbar(DeviceModelFrame,orient=VERTICAL, command=DeviceModelCountryCostTreeview.yview)           # <--- Make sure the frame is correct !!
            DeviceModelCountryCostTreeviewxsb = Scrollbar(DeviceModelFrame,orient=HORIZONTAL, command=DeviceModelCountryCostTreeview.xview)         # <--- Make sure the frame is correct !!
            DeviceModelCountryCostTreeview['yscroll'] = DeviceModelCountryCostTreeviewysb.set
            DeviceModelCountryCostTreeview['xscroll'] = DeviceModelCountryCostTreeviewxsb.set

            # setup headings and column 
            DeviceModelCountryCostTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
            DeviceModelCountryCostTreeview.heading('#1', text='Country ID',   anchor=W) # E for East and W for West
            DeviceModelCountryCostTreeview.heading('#2', text='Device Model ID',    anchor=W) # E for East and W for West
            DeviceModelCountryCostTreeview.heading('#3', text='Outsource Cost',  anchor=W) 
            DeviceModelCountryCostTreeview.heading('#4', text='Added By',     anchor=W) 
                    
            DeviceModelCountryCostTreeview.column('#0', stretch=1, width=3 , anchor=W)
            DeviceModelCountryCostTreeview.column('#1', stretch=1, width=10, anchor=W)
            DeviceModelCountryCostTreeview.column('#2', stretch=1, width=10)
            DeviceModelCountryCostTreeview.column('#3', stretch=1, width=10)
            DeviceModelCountryCostTreeview.column('#4', stretch=1, width=10)

            DeviceModelCountryCostTreeview.grid(row=1, column=0, sticky=NSEW)
            DeviceModelCountryCostTreeviewysb.grid(row=1, column=1, sticky=NS)
            DeviceModelCountryCostTreeviewxsb.grid(row=2, column=0, sticky=EW)

            # create fonts and tags
            DeviceModelCountryCostTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
            # Bind the double Click
            #DeviceModelCountryCostTreeview.bind('<ButtonRelease-1>', self.on_DeviceModelCountry_Tree_select_click) # When Select the Tree
            #self.DeviceModelCountryTreeview.bind("<Double-1>", self.On_DeviceModelCountry_Tree_Refresh)         

            #------------------ TREE VIEW For Device Model Country Cost Database <END> -----------------------------------
            Parameters = []
            Parameters = ["AVPN","AVPN",10.1,"USA","EDIT",""]
            Parameters = ["AVPN","AVPN",0,"","ADD",""]
            Treeview = DeviceModelCountryCostTreeview
            DeviceModelCountry = Class_DeviceModelCountry("BV",Windows_Scaling,Parameters,Treeview)
            DeviceModelCountry.Display_DeviceModelCountry_Window()   
            DeviceModelWindow.mainloop()

        else:
            mbox.showerror(master=DeviceModelFrame,title='Device Model',
                            message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                    + 'Please make sure the ODBC DSN Name mathes: ['
                                    + self.ODBC_name + "]")

                
################


def Main():
    global DeviceModelWindowExist
    global db

    db = ODBC("BV")
    print ("Testing the DeviceModelCountry Class....:")
    DeviceModelWindowExist = False
    Display_DeviceModel_Window()
    #Parameters = []
    #Parameters = ["AVPN","AVPN","USA","EDIT",""]
    #Parameters = ["AVPN","AVPN","","ADD",""]
    #Treeview = DeviceModelCountryCostTreeview
    #DeviceModelCountry = Class_DeviceModelCountry("BV",Windows_Scaling,Parameters,Treeview)
    #DeviceModelCountry.Display_DeviceModelCountry_Window()


if __name__ == '__main__':
    Main()

