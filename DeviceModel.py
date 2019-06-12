#==========================================================================
# Program: DeviceModel.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-23-2017
# Date Last Modified: Dec-29-2017
# Summary: This is Class to for the DeviceModel
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
    from DeviceType import *
    Is_DeviceType_Available = True
    DeviceType = Class_DeviceType(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO DeviceType Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_DeviceType_Available = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Device Model','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

try:
    from DeviceModelCountry import *
    Is_DeviceModelCountry_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO DeviceModelCountry Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_DeviceModelCountry_Available = False  



#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_DeviceModel:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.DeviceModelWindowExist = False
        self.DeviceModelCountryWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.DeviceModelCountryID = "Country"
        
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

#---------------------------- DeviceType SELECTION SECTION <BEGIN> ------------------------
        
    def Display_DeviceType_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','DeviceType Window']    
            Logging.Log(Parameter)
        DeviceType.Display_DeviceType_Window()

    def on_DeviceType_combo_changed(self,event):
        #print (self.ComboBoxDeviceTypeID.current())
        #print (self.DeviceTypeIDComboBox[self.ComboBoxDeviceTypeID.current()])
        #print (self.DeviceTypeNameComboBox[self.ComboBoxDeviceTypeID.current()])

        self.DeviceModelIDFrameEntry['state'] = 'normal'
        self.DeviceModelIDFrameEntry.delete(0,END)
        self.DeviceModelIDFrameEntry['state'] = 'readonly'
            
        self.DeviceModelNameFrameEntry['state'] = 'normal'
        self.DeviceModelNameFrameEntry.delete(0,END)
        self.DeviceModelNameFrameEntry['state'] = 'readonly'

        self.DeviceModelCostFrameEntry['state'] = 'normal'
        self.DeviceModelCostFrameEntry.delete(0,END)
        self.DeviceModelCostFrameEntry['state'] = 'readonly'
         
        if self.db.Connect():
            # SQL Querry to the DeviceModel Table
            sql = """
                    SELECT * FROM Device_Model
                    WHERE Device_Type_ID = '%s'
                  """ % (self.DeviceTypeIDComboBox[self.ComboBoxDeviceTypeID.current()])

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.DeviceModelCountryCostTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.DeviceModelCountryCostTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.DeviceModelTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.DeviceModelTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

            if (Is_DeviceModelCountry_Available and self.DeviceModelCountryWindowExist):
                self.DeviceModelCountry.on_DeviceModelCountryWindow_quit()
                self.DeviceModelCountryWindowExist = False
            
            if (self.db.Execute(sql)):
                self.DeviceModelIDArray = []
                self.DeviceModelNameArray = []
                self.VendorIDArray = []
                self.OutsourceCostArray = []
                self.ExecutedbyUserIDArray = []
                self.DeviceTypeIDArray = []
                self.results = []
                self.results2 = []

                self.data_ready = True   
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceModelIDArray.append(self.db.results[i][0].strip())
                    self.DeviceModelNameArray.append(self.db.results[i][1].strip())
                    self.VendorIDArray.append(self.db.results[i][2].strip())
                    self.OutsourceCostArray.append(self.db.results[i][3])
                    self.ExecutedbyUserIDArray.append(self.db.results[i][4].strip())
                    self.DeviceTypeIDArray.append(self.db.results[i][5].strip())
                    i = i + 1
                   
                i = 0
                while (i < len(self.DeviceModelIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.DeviceModelIDArray[i],
                            self.DeviceModelNameArray[i],
                            #self.VendorIDArray[i],
                            self.OutsourceCostArray[i],
                            self.ExecutedbyUserIDArray[i],
                            #self.DeviceTypeIDArray[i]
                            ]
                    self.results.append(item)
                    self.DeviceModelTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonDeviceModelAdd['state']    = ACTIVE
                self.ButtonDeviceModelEdit['state']   = DISABLED
                self.ButtonDeviceModelRemove['state'] = DISABLED
                self.ButtonDeviceModelOK['state']     = DISABLED
                self.ButtonDeviceModelCancel['state'] = DISABLED
                self.ButtonDeviceModelAddCountry['state']   = DISABLED
                self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
                self.ButtonDeviceModelUpdateDevices['state'] = DISABLED

            else:
                mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                message = 'No Records found')
                self.data_ready = False
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.DeviceModelTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.DeviceModelTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                self.ButtonDeviceModelAdd['state']    = ACTIVE
                self.ButtonDeviceModelEdit['state']   = DISABLED
                self.ButtonDeviceModelRemove['state'] = DISABLED
                self.ButtonDeviceModelOK['state']     = DISABLED
                self.ButtonDeviceModelCancel['state'] = DISABLED
                self.ButtonDeviceModelAddCountry['state']   = DISABLED
                self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
                self.ButtonDeviceModelUpdateDevices['state'] = DISABLED

                self.DeviceModelIDFrameEntry['state'] = 'normal'
                self.DeviceModelIDFrameEntry.delete(0,END)
                self.DeviceModelIDFrameEntry['state'] = 'readonly'
                
                self.DeviceModelNameFrameEntry['state'] = 'normal'
                self.DeviceModelNameFrameEntry.delete(0,END)
                self.DeviceModelNameFrameEntry['state'] = 'readonly'

                self.DeviceModelCostFrameEntry['state'] = 'normal'
                self.DeviceModelCostFrameEntry.delete(0,END)
                self.DeviceModelCostFrameEntry['state'] = 'readonly'
                
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='DeviceModel',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_DeviceType_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','DeviceType Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.DeviceTypeIDComboBox = []
            self.DeviceTypeNameComboBox = []
            self.DeviceTypeVendorIDComboBox = []

            # SQL Querry to the DeviceModel Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """ SELECT * FROM Device_Type ORDER BY Device_Type_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceTypeIDComboBox.append(self.db.results[i][0].strip())
                    self.DeviceTypeNameComboBox.append(self.db.results[i][1].strip())
                    self.DeviceTypeVendorIDComboBox.append(self.db.results[i][2].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxDeviceTypeID['values'] = self.DeviceTypeNameComboBox
                if (len(self.DeviceTypeNameComboBox)== 0):
                    self.ComboBoxDeviceTypeID['state'] = DISABLED
                else:
                    self.ComboBoxDeviceTypeID.current(0)
                    self.on_DeviceModel_Table_Refresh()
            else:
                mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                message = 'No Device Type Records found')
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        

#---------------------------- DeviceType SELECTION SECTION <END> ------------------------

        
    def on_DeviceModelWindow_quit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','CLOSE Window']    
            Logging.Log(Parameter)

        if (self.DeviceModelWindowExist):
            self.DeviceModelWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Device Model','CLOSE Window']    
                Logging.Log(Parameter)
            if (Is_DeviceModelCountry_Available and self.DeviceModelCountryWindowExist):
                self.DeviceModelCountry.on_DeviceModelCountryWindow_quit()
            self.DeviceModelWindow.destroy()


    def on_DeviceModel_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','Device Model Refresh Button']    
            Logging.Log(Parameter)
            
        self.on_DeviceType_combo_changed("event")


    def Call_Button_DeviceModelCountry_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','Cost per Country Button']    
            Logging.Log(Parameter)

        Parameters = []
        DeviceModelID = self.DeviceModelIDFrameEntry.get()
        DeviceModelName = self.DeviceModelNameFrameEntry.get()
        Parameters = [DeviceModelID,DeviceModelName,0.0,self.DeviceModelCountryID,"ADD/EDIT","Status"]
        Treeview = self.DeviceModelCountryCostTreeview
        if Is_DeviceModelCountry_Available:
            if not self.DeviceModelCountryWindowExist:
                self.DeviceModelCountryWindowExist = True
                self.DeviceModelCountry = Class_DeviceModelCountry(ODBC_DSN_name,Windows_Scaling,Parameters,Treeview)
                self.DeviceModelCountry.Display_DeviceModelCountry_Window()


    def Call_Button_DeviceModelCountry_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','REMOVE Country Cost Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():            
            if (mbox.askyesnocancel(master=self.DeviceModelFrame,title='Device Model',message = 'Are you Sure you want to Remove '+
                                self.DeviceModelCountryID +'?')):

                DeviceModelID = self.DeviceModelIDFrameEntry.get()
                DeviceModelName = self.DeviceModelNameFrameEntry.get()
                Primarykey = DeviceModelID + "-" + self.DeviceModelCountryID
                sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID_Country_ID = '%s'
                  """ % (Primarykey)
                if (self.db.Execute(sql)):
                    sql = "DELETE FROM Device_Model_Country WHERE Device_Model_ID_Country_ID = '%s'" % (Primarykey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        mbox.showwarning(master=self.DeviceModelFrame,title='Device Model',
                                        message = '*** The Cost per Country of the Device Model ID you entered was Removed ***')
                    else:
                        #self.db.Disconnect()
                        mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                        message = '*** The Cost per Country of the Device Model ID you entered was NOT Removed ***')               
                else:
                    mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                    message = '*** The Cost per Country of the Device Model ID you try to Remove Does not exist Anymore ***')
                self.on_DeviceModelCountryCost_search()
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceModelCountry_UpdateDevices(self):
        #print ("Update Devices")
        #self.DeviceModelCountryID = values[0]
        #self.DeviceModelPerCountryCost = float(values[2])
        DeviceModelID = self.DeviceModelIDFrameEntry.get()
        DeviceModelName = self.DeviceModelNameFrameEntry.get()
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','Update Devices Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            sql = """
                    SELECT * FROM Devices
                    WHERE Device_Model_ID = '%s' AND Country_ID = '%s'
                  """ % (DeviceModelID,self.DeviceModelCountryID)
            if (self.db.Execute(sql)): 
                sql = "UPDATE Devices SET Outsource_Cost = '%f', \
                        Executed_by_UserID = '%s' WHERE (Device_Model_ID = '%s' AND Country_ID = '%s')" % (self.DeviceModelPerCountryCost,
                                                                                        self.Username,DeviceModelID,self.DeviceModelCountryID)
                if (self.db.Add_Move_Change_Data(sql)):
                    #self.db.Disconnect()
                    mbox.showwarning(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Cost Per Country of the Device Model ID you selected was Updated ***')
                    self.on_DeviceModel_Table_Refresh()
                    self.DeviceModelNameFrameEntry['state'] = 'readonly'
                else:
                    #self.db.Disconnect()
                    mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Cost Per Country of the Device Model ID you selected was NOT Upadted ***')
                    self.on_DeviceModel_Table_Refresh()
                    self.DeviceModelNameFrameEntry['state'] = 'readonly'
            else:
                mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                message = '*** The Device Model ID in the Country you try to Update Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

       
    def Call_Button_DeviceModel_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','ADD Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.ComboBoxDeviceTypeID['state'] = DISABLED
        
        self.DeviceModelIDFrameEntry['state'] = 'normal'
        self.DeviceModelIDFrameEntry.delete(0,END)

        self.DeviceModelNameFrameEntry['state'] = 'normal'
        self.DeviceModelNameFrameEntry.delete(0,END)

        self.DeviceModelCostFrameEntry['state'] = 'normal'
        self.DeviceModelCostFrameEntry.delete(0,END)


        self.ButtonDeviceModelAdd['state']    = DISABLED
        self.ButtonDeviceModelEdit['state']   = DISABLED
        self.ButtonDeviceModelRemove['state'] = DISABLED
        self.ButtonDeviceModelOK['state']     = ACTIVE
        self.ButtonDeviceModelCancel['state'] = ACTIVE
        self.ButtonDeviceModelAddCountry['state']   = DISABLED
        self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
        self.ButtonDeviceModelUpdateDevices['state'] = DISABLED


    def Call_Button_DeviceModel_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'

        self.ComboBoxDeviceTypeID['state'] = DISABLED
        
        self.DeviceModelIDFrameEntry['state'] = 'readonly'
        self.DeviceModelNameFrameEntry['state'] = 'normal'
        self.DeviceModelCostFrameEntry['state'] = 'normal'

        self.ButtonDeviceModelAdd['state']    = DISABLED
        self.ButtonDeviceModelEdit['state']   = DISABLED
        self.ButtonDeviceModelRemove['state'] = DISABLED
        self.ButtonDeviceModelOK['state']     = ACTIVE
        self.ButtonDeviceModelCancel['state'] = ACTIVE
        self.ButtonDeviceModelAddCountry['state']   = DISABLED
        self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
        self.ButtonDeviceModelUpdateDevices['state'] = DISABLED


    def Call_Button_DeviceModel_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove' 
            if (mbox.askyesnocancel(master=self.DeviceModelFrame,title='Device Model',message = 'Are you Sure you want to Remove it?')):
                DeviceModelID = self.DeviceModelIDFrameEntry.get()
                DeviceModelName = self.DeviceModelNameFrameEntry.get()
                DeviceTypeID = self.DeviceTypeIDComboBox[self.ComboBoxDeviceTypeID.current()]
                PrimaryKey = DeviceModelID
                if (DeviceModelID == 'UNKNOWN'):
                    mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Device Model ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Facility --------
                    sql = """
                            SELECT * FROM DEVICES
                            WHERE Device_Model_ID = '%s'
                          """ % (DeviceModelID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                    message = '*** The Device Model ID you entered CANNOT be Removed Becuase is used in Devices ***')
                            self.on_DeviceModel_Table_Refresh()
                    else:
                        DeviceModelID = self.DeviceModelIDFrameEntry.get()
                        DeviceModelName = self.DeviceModelNameFrameEntry.get()
                        sql = """
                                SELECT * FROM Device_Model_Country
                                WHERE Device_Model_ID = '%s'
                                ORDER BY Country_ID
                              """ % (DeviceModelID)
                        if (self.db.Execute(sql)):
                            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                    message = '*** The Device Model ID you entered CANNOT be Removed Becuase it has Cost per Countries ***')
                            self.on_DeviceModel_Table_Refresh()
                        else:
                            sql = """
                                    SELECT * FROM Device_Model
                                    WHERE Device_Model_ID = '%s'
                                  """ % (PrimaryKey)
                            if (self.db.Execute(sql)):
                                sql = "DELETE FROM Device_Model WHERE Device_Model_ID = '%s'" % (PrimaryKey)
                                if (self.db.Add_Move_Change_Data(sql)):
                                    #self.db.Disconnect()
                                    mbox.showwarning(master=self.DeviceModelFrame,title='Device Model',
                                            message = '*** The Device Model ID you entered was Removed ***')
                                    self.on_DeviceModel_Table_Refresh()
                                else:
                                    #self.db.Disconnect()
                                    mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                            message = '*** The Device Model ID you entered was NOT Removed ***')
                                    self.on_DeviceModel_Table_Refresh()               
                            else:
                                mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                message = '*** The Device Model ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='DeviceModel',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceModel_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ComboBoxDeviceTypeID['state'] = 'readonly'

            self.ButtonDeviceModelAdd['state']    = DISABLED
            self.ButtonDeviceModelEdit['state']   = DISABLED
            self.ButtonDeviceModelRemove['state'] = DISABLED
            self.ButtonDeviceModelOK['state']     = DISABLED 
            self.ButtonDeviceModelCancel['state'] = DISABLED
            self.ButtonDeviceModelAddCountry['state']   = DISABLED
            self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
            self.ButtonDeviceModelUpdateDevices['state'] = DISABLED

            DeviceModelID = self.DeviceModelIDFrameEntry.get()
            DeviceModelName = self.DeviceModelNameFrameEntry.get()
            DeviceModelCost = self.DeviceModelCostFrameEntry.get()
            DeviceTypeID = self.DeviceTypeIDComboBox[self.ComboBoxDeviceTypeID.current()]
            DeviceTypeVendorID = self.DeviceTypeVendorIDComboBox[self.ComboBoxDeviceTypeID.current()]
            PrimaryKey = DeviceModelID
            if (len(self.DeviceModelCostFrameEntry.get()) > 0):
                DeviceModelCost = float(self.DeviceModelCostFrameEntry.get())
            else:
                DeviceModelCost = 0
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(DeviceModelID) > 0) and (len(DeviceModelName) > 0)):
                    sql = """
                            SELECT * FROM Device_Model
                            WHERE Device_Model_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Device Model ID you entered already exist ***')
                    else:
                        '''
                        Device_Model_ID         CHAR(30) NOT NULL PRIMARY KEY,
                        Device_Model_Name       CHAR(50),
                        Vendor_ID               CHAR(20),
                        Outsource_Cost          FLOAT,
                        Executed_by_UserID      CHAR(20),
                        Device_Type_ID          CHAR(30))"""
                        '''
                        sql = "INSERT INTO Device_Model(Device_Model_ID, Device_Model_Name, Vendor_ID, \
                                   Outsource_Cost, Executed_by_UserID, Device_Type_ID) \
                                   VALUES ('%s','%s','%s', '%f','%s','%s')" % \
                                   (PrimaryKey,DeviceModelName,DeviceTypeVendorID, DeviceModelCost, self.Username,DeviceTypeID )
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.DeviceModelFrame,title='Device Model',
                                message = '*** The Device Model ID you entered was Added ***')
                            self.on_DeviceModel_Table_Refresh()
                            self.DeviceModelIDFrameEntry['state'] = 'readonly'                           
                            self.DeviceModelNameFrameEntry['state'] = 'readonly'
                            self.DeviceModelCostFrameEntry['state'] = 'readonly'

                        else:
                            #print ("Error adding the record, posible dupliated it")
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                message = '*** The Device Model ID you entered was NOT Added ***')
                            self.on_DeviceModel_Table_Refresh()
                else:
                     mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Device Model ID and DeviceModel Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                if (len(DeviceModelName) == 0):
                    mbox.showerror(master=self.DeviceModelFrame,title='DeviceModel',
                    message = '*** The DeviceModel Name Cannot be BLANK ***')
                else:
                    sql = """
                            SELECT * FROM Device_Model
                            WHERE Device_Model_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)): 
                        sql = "UPDATE Device_Model SET Device_Model_Name = '%s', Vendor_ID = '%s', Outsource_Cost = '%f', \
                               Executed_by_UserID = '%s' WHERE Device_Model_ID = '%s'" % (DeviceModelName,DeviceTypeVendorID, DeviceModelCost, self.Username,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.DeviceModelFrame,title='Device Model',
                                message = '*** The Device Model ID you entered was Updated ***')
                            self.on_DeviceModel_Table_Refresh()
                            self.DeviceModelNameFrameEntry['state'] = 'readonly'
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                                message = '*** The Device Model ID you entered was NOT Upadted ***')
                            self.on_DeviceModel_Table_Refresh()
                            self.DeviceModelNameFrameEntry['state'] = 'readonly'
                    else:
                        mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                        message = '*** The Device Model ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_DeviceModel_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device Model','CANCEL Button']    
            Logging.Log(Parameter)
        self.ButtonDeviceModelAdd['state']    = ACTIVE
        self.ButtonDeviceModelEdit['state']   = DISABLED
        self.ButtonDeviceModelRemove['state'] = DISABLED
        self.ButtonDeviceModelOK['state']     = DISABLED
        self.ButtonDeviceModelCancel['state'] = DISABLED
        self.ButtonDeviceModelAddCountry['state']   = DISABLED
        self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
        self.ButtonDeviceModelUpdateDevices['state'] = DISABLED

        self.ComboBoxDeviceTypeID['state'] = 'readonly'

        self.DeviceModelIDFrameEntry['state'] = 'normal'
        self.DeviceModelIDFrameEntry.delete(0,END)
        self.DeviceModelIDFrameEntry['state'] = 'readonly'
            
        self.DeviceModelNameFrameEntry['state'] = 'normal'
        self.DeviceModelNameFrameEntry.delete(0,END)
        self.DeviceModelNameFrameEntry['state'] = 'readonly'

        self.DeviceModelCostFrameEntry['state'] = 'normal'
        self.DeviceModelCostFrameEntry.delete(0,END)
        self.DeviceModelCostFrameEntry['state'] = 'readonly'              

        self.Selection = 'cancel'


    def on_DeviceModelCountryCost_Tree_select_click(self,event):
        #print ("CLICK")
        curItem = self.DeviceModelCountryCostTreeview.focus()
        dic = self.DeviceModelCountryCostTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.DeviceModelCountryID = values[0]
            self.DeviceModelPerCountryCost = float(values[2])
            self.ButtonDeviceModelRemoveCountry['state'] = ACTIVE
            self.ButtonDeviceModelUpdateDevices['state'] = ACTIVE
                    
    def on_DeviceModelCountryCost_search(self):
        self.DeviceModelCountryID = 'Country'
        #self.ButtonDeviceModelAddCountry['state']   = DISABLED
        self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
        self.ButtonDeviceModelUpdateDevices['state'] = DISABLED
        if self.db.Connect():
            # SQL Querry to the DeviceModelCountry Table  
            '''

                self.DeviceModelIDArray = []
                self.DeviceModelNameArray = []
                self.VendorIDArray = []
                self.OutsourceCostArray = []
                self.ExecutedbyUserIDArray = []
                self.DeviceTypeIDArray = []
            '''
            DeviceModelID = self.DeviceModelIDFrameEntry.get()
            DeviceModelName = self.DeviceModelNameFrameEntry.get()
            sql = """
                    SELECT * FROM Device_Model_Country
                    WHERE Device_Model_ID = '%s'
                    ORDER BY Country_ID
                  """ % (DeviceModelID)

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.DeviceModelCountryCostTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.DeviceModelCountryCostTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            
            if (self.db.Execute(sql)):
                self.DeviceModelCountryTablePKArray = []
                self.DeviceModelCountryTableDeviceModelIDArray= []
                self.DeviceModelCountryTableCountryIDArray = []
                self.DeviceModelCountryTableOutsourceCostArray = []
                self.DeviceModelCountryTableExecutedbyArray = []
                self.results2 = []
                
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
                    self.results2.append(item)
                    self.DeviceModelCountryCostTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                message = 'No Records found')       
        else:
            mbox.showerror(master=self.DeviceModelFrame,title='Device Model',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def on_DeviceModel_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.DeviceModelTreeview.focus()
        dic = self.DeviceModelTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])

            #self.DeviceModelIDArray[curItem],
            #self.DeviceModelNameArray[curItem],
            #self.VendorIDArray[curItem],
            #self.OutsourceCostArray[curItem],
            #self.ExecutedbyUserIDArray[curItem],
            #self.DeviceTypeIDArray[curItem]

            self.ButtonDeviceModelAdd['state']    = ACTIVE
            self.ButtonDeviceModelEdit['state']   = ACTIVE
            self.ButtonDeviceModelRemove['state'] = ACTIVE
            self.ButtonDeviceModelOK['state']     = DISABLED
            self.ButtonDeviceModelCancel['state'] = DISABLED
            self.ButtonDeviceModelAddCountry['state']   = ACTIVE
            self.ButtonDeviceModelRemoveCountry['state'] = DISABLED
            self.ButtonDeviceModelUpdateDevices['state'] = DISABLED

            self.ComboBoxDeviceTypeID['state'] = 'readonly'

            self.DeviceModelIDFrameEntry['state'] = 'normal'
            self.DeviceModelIDFrameEntry.delete(0,END)
            self.DeviceModelIDFrameEntry.insert(0,self.DeviceModelIDArray[curItem])
            self.DeviceModelIDFrameEntry['state'] = 'readonly'
            
            self.DeviceModelNameFrameEntry['state'] = 'normal'
            self.DeviceModelNameFrameEntry.delete(0,END)
            self.DeviceModelNameFrameEntry.insert(0,self.DeviceModelNameArray[curItem])
            self.DeviceModelNameFrameEntry['state'] = 'readonly'

            self.DeviceModelCostFrameEntry['state'] = 'normal'
            self.DeviceModelCostFrameEntry.delete(0,END)
            self.DeviceModelCostFrameEntry.insert(0,str(self.OutsourceCostArray[curItem]))
            self.DeviceModelCostFrameEntry['state'] = 'readonly'

            if (Is_DeviceModelCountry_Available and self.DeviceModelCountryWindowExist):
                self.DeviceModelCountry.on_DeviceModelCountryWindow_quit()
                self.DeviceModelCountryWindowExist = False

            self.on_DeviceModelCountryCost_search()


    def On_DeviceModel_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_DeviceModel_Table_Refresh()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Device Model Cost"
                Tab1 = "Device Model Cost Per Country"
                Tabs = [Tab0,Tab1]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.DeviceModelTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                ExcellFile.Add_DataToWorksheet(self.DeviceModelCountryCostTreeviewDataColumns,Row,Column,Tab1,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                Row = 3
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab1,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.DeviceModelFrame,title='Device Model List',
                            message = '!!! The File was saved !!!')
                    self.DeviceModelWindow.title("Device Model                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Device Model','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.DeviceModelFrame,title='Device Model List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_DeviceModel_Window(self):
        #global DeviceModelLabel
 
        if not self.DeviceModelWindowExist:
            # Set up the Window
            self.DeviceModelWindowExist = True
            self.DeviceModelWindow = Tk()
            self.DeviceModelWindow.geometry('1000x700+350+70')
            self.DeviceModelWindow.title("Device Model")
            self.DeviceModelWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceModelWindow_quit)
            self.DeviceModelWindow.call('tk', 'scaling', self.Windows_Scaling)
            DeviceModelLabel = Label(self.DeviceModelWindow,text="Helvetica", font=("Helvetica", 19))
            DeviceModelLabel["text"] = "Device Model"
            DeviceModelLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.DeviceModelWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_DeviceModelWindow_quit)
            self.DeviceModelWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.DeviceModelFrame = Frame(self.DeviceModelWindow)
            self.DeviceModelFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.DeviceModelFrame.rowconfigure(0, weight=1)
            self.DeviceModelFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_DeviceType_Available):
                        self.ButtonDeviceTypeAdd = Button(self.DeviceModelFrame, text = '+ Device Type', command = self.Display_DeviceType_Window, state=ACTIVE)
                        self.ButtonDeviceTypeAdd.place(x = 450, y = 7, width=85, height=25)

                        self.ButtonDeviceTypeRefresh = Button(self.DeviceModelFrame, text = 'Refresh', command = self.on_DeviceType_Table_Refresh, state=ACTIVE)
                        self.ButtonDeviceTypeRefresh.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonDeviceModelAdd = Button(self.DeviceModelFrame, text = 'Add', command = self.Call_Button_DeviceModel_Add, state=DISABLED)
                    self.ButtonDeviceModelAdd.place(x = 450, y = 37, width=85, height=25)

                    self.ButtonDeviceModelEdit = Button(self.DeviceModelFrame, text = 'Edit', command = self.Call_Button_DeviceModel_Edit, state=DISABLED)
                    self.ButtonDeviceModelEdit.place(x = 550, y = 37, width=75, height=25)

                    self.ButtonDeviceModelRemove = Button(self.DeviceModelFrame, text = 'Remove', command = self.Call_Button_DeviceModel_Remove, state=DISABLED)
                    self.ButtonDeviceModelRemove.place(x = 650, y = 37, width=75, height=25)

                    self.ButtonDeviceModelOK = Button(self.DeviceModelFrame, text = 'OK / UPDATE', command = self.Call_Button_DeviceModel_OK, state=DISABLED)
                    self.ButtonDeviceModelOK.place(x = 750, y = 37, width=100, height=25)

                    self.ButtonDeviceModelCancel = Button(self.DeviceModelFrame, text = 'Cancel', command = self.Call_Button_DeviceModel_Cancel, state=DISABLED)
                    self.ButtonDeviceModelCancel.place(x = 875, y = 37, width=75, height=25)

                    self.ButtonDeviceModelAddCountry = Button(self.DeviceModelFrame, text = 'Cost Per Country', command = self.Call_Button_DeviceModelCountry_Add, state=DISABLED)
                    self.ButtonDeviceModelAddCountry.place(x = 450, y = 67, width=100, height=25)

                    self.ButtonDeviceModelRemoveCountry = Button(self.DeviceModelFrame, text = 'Remove Country', command = self.Call_Button_DeviceModelCountry_Remove, state=DISABLED)
                    self.ButtonDeviceModelRemoveCountry.place(x = 575, y = 67, width=100, height=25)

                    self.ButtonDeviceModelUpdateDevices = Button(self.DeviceModelFrame, text = 'Update Devices', command = self.Call_Button_DeviceModelCountry_UpdateDevices, state=DISABLED)
                    self.ButtonDeviceModelUpdateDevices.place(x = 700, y = 67, width=100, height=25)


                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.DeviceModelFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=97)

                    # Setup Labels
                    DeviceTypeIDFrameLabel = Label(self.DeviceModelFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceTypeIDFrameLabel["text"] = "Device Type:"
                    DeviceTypeIDFrameLabel.place(x=10, y=10)
                    
                    DeviceModelIDFrameLabel = Label(self.DeviceModelFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelIDFrameLabel["text"] = "Device Model ID:"
                    DeviceModelIDFrameLabel.place(x=10, y=40)
                    
                    DeviceModelNameFrameLabel = Label(self.DeviceModelFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelNameFrameLabel["text"] = "Device Model Name:"
                    DeviceModelNameFrameLabel.place(x=10, y=70)

                    DeviceModelNameFrameLabel = Label(self.DeviceModelFrame,text="Helvetica", font=("Helvetica", 10))
                    DeviceModelNameFrameLabel["text"] = "Outsource Cost:"
                    DeviceModelNameFrameLabel.place(x=10, y=100)

                   # Setup Entry
                    self.DeviceModelIDFrameEntry = Entry(self.DeviceModelFrame)
                    self.DeviceModelIDFrameEntry['width']=40
                    self.DeviceModelIDFrameEntry.place(x=140, y=40)
                    self.DeviceModelIDFrameEntry['state'] = DISABLED

                    self.DeviceModelNameFrameEntry = Entry(self.DeviceModelFrame)
                    self.DeviceModelNameFrameEntry['width']=40
                    self.DeviceModelNameFrameEntry.place(x=140, y=70)
                    self.DeviceModelNameFrameEntry['state'] = DISABLED               

                    self.DeviceModelCostFrameEntry = Entry(self.DeviceModelFrame)
                    self.DeviceModelCostFrameEntry['width']=20
                    self.DeviceModelCostFrameEntry.place(x=140, y=100)
                    self.DeviceModelCostFrameEntry['state'] = DISABLED               

                    #------------------ TREE VIEW For Device Model Country Cost Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.DeviceModelCountryCostTreeviewDataColumns = ('Country ID','Device Model ID','Outsource Cost','Added By')
                    
                    self.DeviceModelCountryCostTreeview = ttk.Treeview(self.DeviceModelFrame,columns=self.DeviceModelCountryCostTreeviewDataColumns, height=10) # <--- Make sure the frame is correct !!
                    self.DeviceModelCountryCostTreeviewysb = Scrollbar(self.DeviceModelFrame,orient=VERTICAL, command=self.DeviceModelCountryCostTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.DeviceModelCountryCostTreeviewxsb = Scrollbar(self.DeviceModelFrame,orient=HORIZONTAL, command=self.DeviceModelCountryCostTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.DeviceModelCountryCostTreeview['yscroll'] = self.DeviceModelCountryCostTreeviewysb.set
                    self.DeviceModelCountryCostTreeview['xscroll'] = self.DeviceModelCountryCostTreeviewxsb.set

                    # setup headings and column 
                    self.DeviceModelCountryCostTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.DeviceModelCountryCostTreeview.heading('#1', text='Country ID',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelCountryCostTreeview, 'Country ID', False)) # E for East and W for West
                    self.DeviceModelCountryCostTreeview.heading('#2', text='Device Model ID',    anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelCountryCostTreeview, 'Device Model ID', False)) # E for East and W for West
                    self.DeviceModelCountryCostTreeview.heading('#3', text='Outsource Cost',  anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelCountryCostTreeview, 'Outsource Cost', False)) 
                    self.DeviceModelCountryCostTreeview.heading('#4', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelCountryCostTreeview, 'Added By', False)) 
                    
                    self.DeviceModelCountryCostTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.DeviceModelCountryCostTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.DeviceModelCountryCostTreeview.column('#2', stretch=1, width=10)
                    self.DeviceModelCountryCostTreeview.column('#3', stretch=1, width=10)
                    self.DeviceModelCountryCostTreeview.column('#4', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    #self.DeviceModelCountryCostTreeview.grid(row=11, column=0, sticky=NSEW)
                    #self.DeviceModelCountryCostTreeviewysb.grid(row=11, column=1, sticky=NS)
                    #self.DeviceModelCountryCostTreeviewxsb.grid(row=12, column=0, sticky=EW)

                    self.DeviceModelCountryCostTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.DeviceModelCountryCostTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.DeviceModelCountryCostTreeviewxsb.grid(row=2, column=0, sticky=EW)

                    # create fonts and tags
                    self.DeviceModelCountryCostTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.DeviceModelCountryCostTreeview.bind('<ButtonRelease-1>', self.on_DeviceModelCountryCost_Tree_select_click) # When Select the Tree
                    #self.DeviceModelTreeview.bind("<Double-1>", self.On_DeviceModel_Tree_Refresh)         

                    #------------------ TREE VIEW For Device Model Country Cost Database <END> -----------------------------------


                    
                    #------------------ TREE VIEW For Device Model Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.DeviceModelTreeviewDataColumns = ('Device Model ID','Device Model Name','Default Outsource Cost','Added By')
                    
                    self.DeviceModelTreeview = ttk.Treeview(self.DeviceModelFrame,columns=self.DeviceModelTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.DeviceModelTreeviewysb = Scrollbar(self.DeviceModelFrame,orient=VERTICAL, command=self.DeviceModelTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.DeviceModelTreeviewxsb = Scrollbar(self.DeviceModelFrame,orient=HORIZONTAL, command=self.DeviceModelTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.DeviceModelTreeview['yscroll'] = self.DeviceModelTreeviewysb.set
                    self.DeviceModelTreeview['xscroll'] = self.DeviceModelTreeviewxsb.set

                    # setup headings and column 
                    self.DeviceModelTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.DeviceModelTreeview.heading('#1', text='Device Model ID',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelTreeview, 'Device Model ID', False)) # E for East and W for West
                    self.DeviceModelTreeview.heading('#2', text='Device Model Name',    anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelTreeview, 'Device Model Name', False)) # E for East and W for West
                    self.DeviceModelTreeview.heading('#3', text='Default Outsource Cost',  anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelTreeview, 'Default Outsource Cost', False)) 
                    self.DeviceModelTreeview.heading('#4', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.DeviceModelTreeview, 'Added By', False)) 
                    
                    self.DeviceModelTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.DeviceModelTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.DeviceModelTreeview.column('#2', stretch=1, width=10)
                    self.DeviceModelTreeview.column('#3', stretch=1, width=10)
                    self.DeviceModelTreeview.column('#4', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.DeviceModelTreeview.grid(row=11, column=0, sticky=NSEW)
                    self.DeviceModelTreeviewysb.grid(row=11, column=1, sticky=NS)
                    self.DeviceModelTreeviewxsb.grid(row=12, column=0, sticky=EW)

                    #self.DeviceModelTreeview.grid(row=1, column=0, sticky=NSEW)
                    #self.DeviceModelTreeviewysb.grid(row=1, column=1, sticky=NS)
                    #self.DeviceModelTreeviewxsb.grid(row=2, column=0, sticky=EW)

                    # create fonts and tags
                    self.DeviceModelTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.DeviceModelTreeview.bind('<ButtonRelease-1>', self.on_DeviceModel_Tree_select_click) # When Select the Tree
                    #self.DeviceModelTreeview.bind("<Double-1>", self.On_DeviceModel_Tree_Refresh)         

                    #------------------ TREE VIEW For Device Model Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxDeviceTypeID = ttk.Combobox(self.DeviceModelFrame, state='readonly', width = 37)
                    self.ComboBoxDeviceTypeID.bind("<<ComboboxSelected>>", self.on_DeviceType_combo_changed)
                    self.ComboBoxDeviceTypeID.place(x = 140, y = 10)
                    self.on_DeviceType_Table_Refresh()           
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.DeviceModelWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.DeviceModelFrame,title='DeviceModel',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the DeviceModel Class....:")
    DeviceModel = Class_DeviceModel("BV",Windows_Scaling)
    DeviceModel.Display_DeviceModel_Window()


if __name__ == '__main__':
    Main()

