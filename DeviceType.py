#==========================================================================
# Program: DeviceType.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-22-2017
# Date Last Modified: Dec-23-2017
# Summary: This is Class to for the DeviceType
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
    Parameter = ['Device Type','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

try:
    from Vendor import *
    Is_Vendor_Available = True
    Vendor = Class_Vendor(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Vendor Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Vendor_Available = False

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_DeviceType:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.DeviceTypeWindowExist = False
        self.Username = os.getlogin()
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


    def on_DeviceTypeWindow_quit(self):
        if (self.DeviceTypeWindowExist):
            self.DeviceTypeWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Device Type','CLOSE Window']    
                Logging.Log(Parameter) 
            self.DeviceTypeWindow.destroy()

    def on_DeviceType_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.DeviceTypeTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.DeviceTypeTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                    Device_Type_ID          CHAR(30) NOT NULL PRIMARY KEY,
                    Device_Type_Name        CHAR(50),
                    Vendor_ID               CHAR(20),
                    Executed_by_UserID      CHAR(20))"""
            '''
            self.Get_Vendor()
            #self.ComboBoxVendorID['state'] = 'readonly'
            #self.ComboBoxVendorID.current(0)

            # SQL Querry to the DeviceType Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """
                        SELECT * FROM Device_Type
                        ORDER BY Device_Type_Name ASC
                  """
            if (self.db.Execute(sql)):                        
                self.DeviceTypeIDArray = []
                self.DeviceTypeNameArray = []
                self.VendorIDArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceTypeIDArray.append(self.db.results[i][0].strip())
                    self.DeviceTypeNameArray.append(self.db.results[i][1].strip())
                    if (self.db.results[i][2] == None):
                        self.VendorIDArray.append("UNKNOWN")
                    else:
                        self.VendorIDArray.append(self.db.results[i][2].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][3].strip())
                    i = i + 1
                i = 0
                while (i < len(self.DeviceTypeIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.DeviceTypeIDArray[i],
                            self.DeviceTypeNameArray[i],
                            self.VendorIDArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    self.DeviceTypeTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                    self.progress['value'] = i
            else:
                mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                message = 'No Records found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")



    def Call_Button_Vendor_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','Vendor Button']    
            Logging.Log(Parameter)
        Vendor.Display_Vendor_Window()

    def Call_Button_Vendor_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','Vendor Refresh Button']    
            Logging.Log(Parameter)
        self.Get_Vendor()
        self.ComboBoxVendorID['state'] = 'readonly'
        self.ComboBoxVendorID.current(0)

        
    def Call_Button_DeviceType_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.DeviceTypeIDFrameEntry['state'] = 'normal'
        self.DeviceTypeIDFrameEntry.delete(0,END)
        self.DeviceTypeIDFrameEntry['state'] = 'readonly'
        self.DeviceTypeIDFrameEntry['state'] = 'normal'
            
        self.DeviceTypeNameFrameEntry['state'] = 'normal'
        self.DeviceTypeNameFrameEntry.delete(0,END)
        self.DeviceTypeNameFrameEntry['state'] = 'readonly'
        self.DeviceTypeNameFrameEntry['state'] = 'normal'

        self.ComboBoxVendorID['state'] = 'readonly'
        self.ComboBoxVendorID.current(0)

        #DeviceTypeID = self.DeviceTypeIDFrameEntry.get()
        #DeviceTypeName = self.DeviceTypeNameFrameEntry.get()

        self.ButtonDeviceTypeAdd['state']    = DISABLED
        self.ButtonDeviceTypeEdit['state']   = DISABLED
        self.ButtonDeviceTypeRemove['state'] = DISABLED
        self.ButtonDeviceTypeOK['state']     = ACTIVE
        self.ButtonDeviceTypeCancel['state'] = ACTIVE

        self.ButtonVendorTypeAdd['state'] = ACTIVE
        self.ButtonVendorTypeRefresh['state'] = ACTIVE



    def Call_Button_DeviceType_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.DeviceTypeIDFrameEntry['state'] = 'readonly'
        self.DeviceTypeNameFrameEntry['state'] = 'normal'
        self.ComboBoxVendorID['state'] = 'readonly'

        self.ButtonDeviceTypeAdd['state']    = DISABLED
        self.ButtonDeviceTypeEdit['state']   = DISABLED
        self.ButtonDeviceTypeRemove['state'] = DISABLED
        self.ButtonDeviceTypeOK['state']     = ACTIVE
        self.ButtonDeviceTypeCancel['state'] = ACTIVE
        self.ButtonVendorTypeAdd['state'] = ACTIVE
        self.ButtonVendorTypeRefresh['state'] = ACTIVE


    def Call_Button_DeviceType_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.DeviceTypeFrame,title='Device Type',message = 'Are you Sure you want to Remove it?')):
                self.ButtonDeviceTypeAdd['state']    = ACTIVE
                self.ButtonDeviceTypeEdit['state']   = DISABLED
                self.ButtonDeviceTypeRemove['state'] = DISABLED
                self.ButtonDeviceTypeOK['state']     = DISABLED
                self.ButtonDeviceTypeCancel['state'] = DISABLED                
                self.ButtonVendorTypeAdd['state'] = DISABLED
                self.ButtonVendorTypeRefresh['state'] = DISABLED
                DeviceTypeID = self.DeviceTypeIDFrameEntry.get()
                DeviceTypeName = self.DeviceTypeNameFrameEntry.get()
                DeviceTypeVendorID = self.VendorTableVendorIDArray[self.ComboBoxVendorID.current()]
                #print (DeviceTypeVendorID)
                if (DeviceTypeID == "UNKNOWN"):
                    mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = '*** The Device Type ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM Device_Model
                            WHERE Device_Type_ID = '%s'
                          """ % (DeviceTypeID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                                    message = '*** The Device Type ID you entered CANNOT be Removed, \r\n Becuase is used in a Device Model ***')
                            self.on_DeviceType_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Device_Type
                                WHERE Device_Type_ID = '%s'
                              """ % (DeviceTypeID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Device_Type WHERE Device_Type_ID = '%s'" % (DeviceTypeID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.DeviceTypeFrame,title='Device Type',
                                        message = '*** The Device Type ID you entered was Removed ***')
                                self.on_DeviceType_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                                        message = '*** The Device Type ID you entered was NOT Removed ***')
                                self.on_DeviceType_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                            message = '*** The Device Type ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceType_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonDeviceTypeAdd['state']    = DISABLED
            self.ButtonDeviceTypeEdit['state']   = DISABLED
            self.ButtonDeviceTypeRemove['state'] = DISABLED
            self.ButtonDeviceTypeOK['state']     = DISABLED 
            self.ButtonDeviceTypeCancel['state'] = DISABLED 
            self.ButtonVendorTypeAdd['state'] = DISABLED
            self.ButtonVendorTypeRefresh['state'] = DISABLED

            DeviceTypeID = self.DeviceTypeIDFrameEntry.get()
            DeviceTypeName = self.DeviceTypeNameFrameEntry.get()
            DeviceTypeVendorID = self.VendorTableVendorIDArray[self.ComboBoxVendorID.current()]
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(DeviceTypeID) > 0) and (len(DeviceTypeName) > 0)):
                    sql = """
                            SELECT * FROM Device_Type
                            WHERE Device_Type_ID = '%s'
                          """ % (DeviceTypeID)
                    if (self.db.Execute(sql)):
                        self.ButtonDeviceTypeOK['state']     = ACTIVE 
                        self.ButtonDeviceTypeCancel['state'] = ACTIVE
                        self.ButtonVendorTypeAdd['state'] = ACTIVE
                        self.ButtonVendorTypeRefresh['state'] = ACTIVE

                        mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = '*** The Device Type ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Device_Type(Device_Type_ID, \
                                   Device_Type_Name, Vendor_ID,Executed_by_UserID) \
                                   VALUES ('%s', '%s', '%s', '%s')" % \
                                   (DeviceTypeID, DeviceTypeName, DeviceTypeVendorID, self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.DeviceTypeFrame,title='Device Type',
                                message = '*** The Device Type ID you entered was Added ***')
                            self.on_DeviceType_Table_Refresh()
                            self.DeviceTypeIDFrameEntry['state'] = 'readonly'
                            self.DeviceTypeNameFrameEntry['state'] = 'readonly'
                            self.ComboBoxVendorID['state'] = DISABLED
                        else:
                            #print ("Error adding the record, posible dupliated it")
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                                message = '*** The Device Type ID you entered was NOT Added ***')
                            self.on_DeviceType_Table_Refresh()
                            self.DeviceTypeIDFrameEntry['state'] = 'readonly'
                            self.DeviceTypeNameFrameEntry['state'] = 'readonly'
                            self.ComboBoxVendorID['state'] = DISABLED
                else:
                    self.ButtonDeviceTypeOK['state']     = ACTIVE 
                    self.ButtonDeviceTypeCancel['state'] = ACTIVE
                    self.ButtonVendorTypeAdd['state'] = ACTIVE
                    self.ButtonVendorTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = '*** The Device Type ID and Device Type Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Device_Type
                        WHERE Device_Type_ID = '%s'
                      """ % (DeviceTypeID)
                if (len(DeviceTypeName) == 0):
                    self.ButtonDeviceTypeOK['state']     = ACTIVE 
                    self.ButtonDeviceTypeCancel['state'] = ACTIVE
                    self.ButtonVendorTypeAdd['state'] = ACTIVE
                    self.ButtonVendorTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = '*** The DeviceType Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):

                        sql = "UPDATE Device_Type SET Device_Type_Name = '%s', Vendor_ID = '%s', Executed_by_UserID = '%s' WHERE Device_Type_ID = '%s'" % (DeviceTypeName,
                                                                                DeviceTypeVendorID, self.Username,DeviceTypeID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.DeviceTypeFrame,title='Device Type',
                                message = '*** The Device Type ID you entered was Updated ***')
                            self.on_DeviceType_Table_Refresh()
                            self.DeviceTypeNameFrameEntry['state'] = 'readonly'
                            
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                                message = '*** The Device Type ID you entered was NOT Upadted ***')
                            self.on_DeviceType_Table_Refresh()               
                            self.DeviceTypeNameFrameEntry['state'] = 'readonly'
                    else:
                        mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = '*** The DeviceType ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceType_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonDeviceTypeAdd['state']    = ACTIVE
        self.ButtonDeviceTypeEdit['state']   = DISABLED
        self.ButtonDeviceTypeRemove['state'] = DISABLED
        self.ButtonDeviceTypeOK['state']     = DISABLED
        self.ButtonDeviceTypeCancel['state'] = DISABLED

        self.ButtonVendorTypeAdd['state'] = ACTIVE
        self.ButtonVendorTypeRefresh['state'] = DISABLED


        self.DeviceTypeIDFrameEntry['state'] = 'normal'
        self.DeviceTypeIDFrameEntry.delete(0,END)
        self.DeviceTypeIDFrameEntry['state'] = 'readonly'
            
        self.DeviceTypeNameFrameEntry['state'] = 'normal'
        self.DeviceTypeNameFrameEntry.delete(0,END)
        self.DeviceTypeNameFrameEntry['state'] = 'readonly'

        self.ComboBoxVendorID['state'] = DISABLED
        self.ComboBoxVendorID.current(0)

        self.Selection = 'cancel'

    def on_DeviceType_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.DeviceTypeTreeview.focus()
        dic = self.DeviceTypeTreeview.item(curItem)
        curItem = int(dic.get('text')) - 1
        values = dic.get('values')        
        if (len(values) > 0):
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonDeviceTypeAdd['state']    = ACTIVE
            self.ButtonDeviceTypeEdit['state']   = ACTIVE
            self.ButtonDeviceTypeRemove['state'] = ACTIVE
            self.ButtonDeviceTypeOK['state']     = DISABLED
            self.ButtonDeviceTypeCancel['state'] = DISABLED
            self.ButtonVendorTypeAdd['state'] = ACTIVE
            self.ButtonVendorTypeRefresh['state'] = DISABLED

            self.DeviceTypeIDFrameEntry['state'] = 'normal'
            self.DeviceTypeIDFrameEntry.delete(0,END)
            self.DeviceTypeIDFrameEntry.insert(0,str(values[0]).strip())
            self.DeviceTypeIDFrameEntry['state'] = 'readonly'
            
            self.DeviceTypeNameFrameEntry['state'] = 'normal'
            self.DeviceTypeNameFrameEntry.delete(0,END)
            self.DeviceTypeNameFrameEntry.insert(0,str(values[1]).strip())
            self.DeviceTypeNameFrameEntry['state'] = 'readonly'

            i = 0
            while (i < len(self.VendorTableVendorNameArray)):
                if (self.VendorTableVendorIDArray[i] == self.VendorIDArray[curItem]):
                    self.ComboBoxVendorID.current(i)
                    i = i + len(self.VendorTableVendorNameArray)
                else:
                    i = i + 1
            self.ComboBoxVendorID['state'] = DISABLED


    def On_DeviceType_Tree_Refresh(self,event):
        #--- Double Click --
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Type','REFRESH Button']    
            Logging.Log(Parameter)            
        self.on_DeviceType_Table_Refresh()

    
    def Get_Vendor(self):
        self.Is_Get_Vendor = False
        if self.db.Connect():
            # SQL Querry to the Device Type
            sql = """
                    SELECT * FROM Vendor
                    ORDER BY Vendor_Name ASC
                  """
            #print (sql)
            self.VendorTableVendorIDArray = []
            self.VendorTableVendorNameArray = []
            self.ComboBoxVendorID['values'] = "Empty"
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.VendorTableVendorIDArray = []
                self.VendorTableVendorNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.VendorTableVendorIDArray.append(self.db.results[i][0].strip())
                    self.VendorTableVendorNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxVendorID['values'] = self.VendorTableVendorNameArray
                if (len(self.VendorTableVendorNameArray)== 0):
                    self.ComboBoxVendorID['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.ComboBoxVendorID['state'] = DISABLED
                    self.ComboBoxVendorID.current(0)
                    #self.DeviceComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                message = 'No Records found for Vendors')

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Device Type List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.DeviceTypeTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.DeviceTypeFrame,title='Device Type List',
                            message = '!!! The File was saved !!!')
                    self.DeviceTypeWindow.title("Device Type                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Device Type','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.DeviceTypeFrame,title='Device Type List',
                            message = '*** The File was not saved, Perhaps It is already open ***')



    def Display_DeviceType_Window(self):
        global DeviceTypeLabel
        if not self.DeviceTypeWindowExist:
            self.DeviceTypeWindowExist = True
            # Set up the Window
            self.DeviceTypeWindow = Tk()
            self.DeviceTypeWindow.geometry('1000x650+350+70')
            self.DeviceTypeWindow.title("Device Type")
            self.DeviceTypeWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceTypeWindow_quit)
            self.DeviceTypeWindow.call('tk', 'scaling', self.Windows_Scaling)
            DeviceTypeLabel = Label(self.DeviceTypeWindow,text="Helvetica", font=("Helvetica", 19))
            DeviceTypeLabel["text"] = "Device Type"
            DeviceTypeLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.DeviceTypeWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_DeviceTypeWindow_quit)
            self.DeviceTypeWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.DeviceTypeFrame = Frame(self.DeviceTypeWindow)
            self.DeviceTypeFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.DeviceTypeFrame.rowconfigure(0, weight=1)
            self.DeviceTypeFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    DeviceTypeIDFrameLabel = Label(self.DeviceTypeFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceTypeIDFrameLabel["text"] = "Device Type ID:"
                    DeviceTypeIDFrameLabel.place(x=10, y=10)
                    self.DeviceTypeIDFrameEntry = Entry(self.DeviceTypeFrame)
                    self.DeviceTypeIDFrameEntry['width']=50
                    self.DeviceTypeIDFrameEntry.place(x=140, y=10)
                    self.DeviceTypeIDFrameEntry['state'] = 'readonly'
                    
                    DeviceTypeNameFrameLabel = Label(self.DeviceTypeFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceTypeNameFrameLabel["text"] = "Device Type Name:"
                    DeviceTypeNameFrameLabel.place(x=10, y=40)
                    self.DeviceTypeNameFrameEntry = Entry(self.DeviceTypeFrame)
                    self.DeviceTypeNameFrameEntry['width']=50
                    self.DeviceTypeNameFrameEntry.place(x=140, y=40)
                    self.DeviceTypeNameFrameEntry['state'] = 'readonly'                    

                    DeviceTypeNameFrameLabel = Label(self.DeviceTypeFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceTypeNameFrameLabel["text"] = "Vendor Name:"
                    DeviceTypeNameFrameLabel.place(x=10, y=70)
                    self.ComboBoxVendorID = ttk.Combobox(self.DeviceTypeFrame, state='disabled', width = 47)
                    #self.ComboBoxVendorID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxVendorID.place(x = 140, y = 70)

                    self.ButtonVendorTypeAdd = Button(self.DeviceTypeFrame, text = '+ Vendor', command = self.Call_Button_Vendor_Add, state=DISABLED)
                    self.ButtonVendorTypeAdd.place(x = 450, y = 67, width=75, height=25)
                    self.ButtonVendorTypeAdd['state'] = ACTIVE

                    self.ButtonVendorTypeRefresh = Button(self.DeviceTypeFrame, text = 'Refresh', command = self.Call_Button_Vendor_Refresh, state=DISABLED)
                    self.ButtonVendorTypeRefresh.place(x = 550, y = 67, width=75, height=25)
                    self.ButtonVendorTypeRefresh['state'] = DISABLED

                   
                    # Setup Buttons
                    self.ButtonDeviceTypeAdd = Button(self.DeviceTypeFrame, text = 'Add', command = self.Call_Button_DeviceType_Add, state=DISABLED)
                    self.ButtonDeviceTypeAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonDeviceTypeAdd['state'] = ACTIVE

                    self.ButtonDeviceTypeEdit = Button(self.DeviceTypeFrame, text = 'Edit', command = self.Call_Button_DeviceType_Edit, state=DISABLED)
                    self.ButtonDeviceTypeEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonDeviceTypeRemove = Button(self.DeviceTypeFrame, text = 'Remove', command = self.Call_Button_DeviceType_Remove, state=DISABLED)
                    self.ButtonDeviceTypeRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonDeviceTypeOK = Button(self.DeviceTypeFrame, text = 'OK / UPDATE', command = self.Call_Button_DeviceType_OK, state=DISABLED)
                    self.ButtonDeviceTypeOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonDeviceTypeCancel = Button(self.DeviceTypeFrame, text = 'Cancel', command = self.Call_Button_DeviceType_Cancel, state=DISABLED)
                    self.ButtonDeviceTypeCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.DeviceTypeFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For DeviceType Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.DeviceTypeTreeviewDataColumns = ('Device Type ID','Device Type Name','Vendor ID','Added By')
                    
                    self.DeviceTypeTreeview = ttk.Treeview(self.DeviceTypeFrame,columns=self.DeviceTypeTreeviewDataColumns, height=23) # <--- Make sure the frame is correct !!
                    self.DeviceTypeTreeviewysb = Scrollbar(self.DeviceTypeFrame,orient=VERTICAL, command=self.DeviceTypeTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.DeviceTypeTreeviewxsb = Scrollbar(self.DeviceTypeFrame,orient=HORIZONTAL, command=self.DeviceTypeTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.DeviceTypeTreeview['yscroll'] = self.DeviceTypeTreeviewysb.set
                    self.DeviceTypeTreeview['xscroll'] = self.DeviceTypeTreeviewxsb.set

                    # setup headings and column 
                    self.DeviceTypeTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.DeviceTypeTreeview.heading('#1', text='Device Type ID',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTypeTreeview, 'Device Type ID', False)) # E for East and W for West
                    self.DeviceTypeTreeview.heading('#2', text='Device Type Name', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTypeTreeview, 'Device Type Name', False)) 
                    self.DeviceTypeTreeview.heading('#3', text='Vendor ID',     anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTypeTreeview, 'Vendor ID', False)) 
                    self.DeviceTypeTreeview.heading('#4', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTypeTreeview, 'Added By', False)) 
                    
                    self.DeviceTypeTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.DeviceTypeTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.DeviceTypeTreeview.column('#2', stretch=1, width=10)
                    self.DeviceTypeTreeview.column('#3', stretch=1, width=10)
                    self.DeviceTypeTreeview.column('#4', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.DeviceTypeTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.DeviceTypeTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.DeviceTypeTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.DeviceTypeTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.DeviceTypeTreeview.bind('<ButtonRelease-1>', self.on_DeviceType_Tree_select_click) # When Select the Tree
                    self.DeviceTypeTreeview.bind("<Double-1>", self.On_DeviceType_Tree_Refresh)         

                    '''
                    Device_Type_ID          CHAR(30) NOT NULL PRIMARY KEY,
                    Device_Type_Name        CHAR(50),
                    Vendor_ID               CHAR(20),
                    Executed_by_UserID      CHAR(20))"""
                    '''
                    self.Get_Vendor()
                    # SQL Querry to the DeviceType Table
                    sql = """
                                SELECT * FROM Device_Type
                                ORDER BY Device_Type_Name ASC
                              """
                    if (self.db.Execute(sql)):                        
                        self.DeviceTypeIDArray = []
                        self.DeviceTypeNameArray = []
                        self.VendorIDArray = []
                        self.ExecutedbyUserIDArray = []
                        self.results = []
                        i = 0
                        self.data_ready = True
                        while (i < len(self.db.results)):
                            num = i + 1
                            self.DeviceTypeIDArray.append(self.db.results[i][0].strip())
                            self.DeviceTypeNameArray.append(self.db.results[i][1].strip())
                            if (self.db.results[i][2] == None):
                                self.VendorIDArray.append("UNKNOWN")
                            else:
                                self.VendorIDArray.append(self.db.results[i][2].strip())
                            self.ExecutedbyUserIDArray.append(self.db.results[i][3].strip())
                            i = i + 1
                        i = 0
                        while (i < len(self.DeviceTypeIDArray)):
                            num = i + 1
                            tags = "-" # To use in the futire                          
                            item = [
                                    self.DeviceTypeIDArray[i],
                                    self.DeviceTypeNameArray[i],
                                    self.VendorIDArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results.append(item)
                            self.DeviceTypeTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                        message = 'No Records found')
                        self.data_ready = False
                    #------------------ TREE VIEW For DeviceType Database <END> -----------------------------------
                    #self.db.Disconnect()
                    self.DeviceTypeWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.DeviceTypeFrame,title='Device Type',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the DeviceType Class....:")
    DeviceType = Class_DeviceType("BV",Windows_Scaling)
    DeviceType.Display_DeviceType_Window()


if __name__ == '__main__':
    Main()

