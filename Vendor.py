#==========================================================================
# Program: Vendor.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-22-2017
# Date Last Modified: Dec-23-2017
# Summary: This is Class to for the Vendor
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
'''
try:
    from SaveAs import *
    Is_SaveAs_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO SaveAs Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False
'''

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
    Parameter = ['Vendor','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

try:
    from SaveAs import *
    Is_SaveAs_Available = True
    #ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,FileName,"Ver 4.0")
except:
    print ("********************************************************************************** \n")
    print ("*** NO SaveAs Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_Vendor:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.VendorWindowExist = False
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


    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Vendor List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.VendorTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                while (i < len(self.VendorIDArray)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.VendorIDArray[i],
                            self.VendorNameArray[i],
                            self.VendorTelArray[i],
                            self.VendorContractArray[i],
                            self.VendorPOCArray[i],
                            self.VendorTelPOCArray[i],
                            self.VendorEmailPOCArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    #Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.VendorFrame,title='Vendor List',
                            message = '!!! The File was saved !!!')
                    self.VendorWindow.title("Vendor List                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Vendor','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.VendorFrame,title='Vendor List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def on_VendorWindow_quit(self):
        if (self.VendorWindowExist):
            self.VendorWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Vendor','CLOSE Window']    
                Logging.Log(Parameter) 
            self.VendorWindow.destroy()

    def on_Vendor_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():

            self.VendorIDFrameEntry['state'] = 'readonly'
            self.VendorNameFrameEntry['state'] = 'readonly'
            self.VendorTelFrameEntry['state'] = 'readonly'
            self.VendorContractFrameEntry['state'] = 'readonly'
            self.VendorPOCFrameEntry['state'] = 'readonly'
            self.VendorTelPOCFrameEntry['state'] = 'readonly'
            self.VendorEmailPOCFrameEntry['state'] = 'readonly'
            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.VendorTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.VendorTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                    Vendor_ID                  CHAR(20) NOT NULL PRIMARY KEY,
                    Vendor_Name                CHAR(50),
                    Tel_Support                CHAR(50),
                    Contract_No                CHAR(50),
                    Executed_by_UserID         CHAR(20))"""
            '''
            # SQL Querry to the Vendor Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            # SQL Querry to the Vendor Table
            self.data_ready = False
            sql = """
                        SELECT * FROM Vendor
                        ORDER BY Vendor_Name ASC
                  """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.VendorIDArray = []
                self.VendorNameArray = []
                self.VendorTelArray = []
                self.VendorContractArray = []
                self.VendorPOCArray = []
                self.VendorTelPOCArray = []
                self.VendorEmailPOCArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.VendorIDArray.append(self.db.results[i][0].strip())
                    self.VendorNameArray.append(self.db.results[i][1].strip())
                    if (self.db.results[i][2] == None):
                        self.VendorTelArray.append("")
                    else:
                        self.VendorTelArray.append(self.db.results[i][2].strip())
                    if (self.db.results[i][3] == None):
                        self.VendorContractArray.append("")
                    else:
                        self.VendorContractArray.append(self.db.results[i][3].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][4].strip())
                    if (self.db.results[i][5] == None):
                        self.VendorPOCArray.append("")
                    else:    
                        self.VendorPOCArray.append(self.db.results[i][5].strip())
                    if (self.db.results[i][6] == None):
                        self.VendorTelPOCArray.append("")
                    else:    
                        self.VendorTelPOCArray.append(self.db.results[i][6].strip())
                    if (self.db.results[i][7] == None):
                        self.VendorEmailPOCArray.append("")
                    else:    
                        self.VendorEmailPOCArray.append(self.db.results[i][7].strip())                            
                    i = i + 1
                i = 0
                while (i < len(self.VendorIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            self.VendorIDArray[i],
                            self.VendorNameArray[i],
                            self.VendorTelArray[i],
                            self.VendorContractArray[i],                
                            self.VendorPOCArray[i],
                            self.VendorTelPOCArray[i],
                            self.VendorEmailPOCArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]                    
                    self.VendorTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=self.VendorFrame,title='Vendor',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.VendorFrame,title='Vendor',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_Vendor_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Vendor','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.VendorIDFrameEntry['state'] = 'normal'
        self.VendorIDFrameEntry.delete(0,END)
        self.VendorIDFrameEntry['state'] = 'readonly'
        self.VendorIDFrameEntry['state'] = 'normal'
            
        self.VendorNameFrameEntry['state'] = 'normal'
        self.VendorNameFrameEntry.delete(0,END)
        self.VendorNameFrameEntry['state'] = 'readonly'
        self.VendorNameFrameEntry['state'] = 'normal'

        self.VendorTelFrameEntry['state'] = 'normal'
        self.VendorTelFrameEntry.delete(0,END)
        self.VendorTelFrameEntry['state'] = 'readonly'
        self.VendorTelFrameEntry['state'] = 'normal'

        self.VendorContractFrameEntry['state'] = 'normal'
        self.VendorContractFrameEntry.delete(0,END)
        self.VendorContractFrameEntry['state'] = 'readonly'
        self.VendorContractFrameEntry['state'] = 'normal'

        self.VendorPOCFrameEntry['state'] = 'normal'
        self.VendorPOCFrameEntry.delete(0,END)
        self.VendorPOCFrameEntry['state'] = 'readonly'
        self.VendorPOCFrameEntry['state'] = 'normal'

        self.VendorTelPOCFrameEntry['state'] = 'normal'
        self.VendorTelPOCFrameEntry.delete(0,END)
        self.VendorTelPOCFrameEntry['state'] = 'readonly'
        self.VendorTelPOCFrameEntry['state'] = 'normal'

        self.VendorEmailPOCFrameEntry['state'] = 'normal'
        self.VendorEmailPOCFrameEntry.delete(0,END)
        self.VendorEmailPOCFrameEntry['state'] = 'readonly'
        self.VendorEmailPOCFrameEntry['state'] = 'normal'
        
        self.ButtonVendorAdd['state']    = DISABLED
        self.ButtonVendorEdit['state']   = DISABLED
        self.ButtonVendorRemove['state'] = DISABLED
        self.ButtonVendorOK['state']     = ACTIVE
        self.ButtonVendorCancel['state'] = ACTIVE


    def Call_Button_Vendor_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Vendor','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.VendorIDFrameEntry['state'] = 'readonly'
        self.VendorNameFrameEntry['state'] = 'normal'
        self.VendorTelFrameEntry['state'] = 'normal'
        self.VendorContractFrameEntry['state'] = 'normal'
        self.VendorPOCFrameEntry['state'] = 'normal'
        self.VendorTelPOCFrameEntry['state'] = 'normal'
        self.VendorEmailPOCFrameEntry['state'] = 'normal'

        self.ButtonVendorAdd['state']    = DISABLED
        self.ButtonVendorEdit['state']   = DISABLED
        self.ButtonVendorRemove['state'] = DISABLED
        self.ButtonVendorOK['state']     = ACTIVE
        self.ButtonVendorCancel['state'] = ACTIVE


    def Call_Button_Vendor_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Vendor','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.VendorFrame,title='Vendor',message = 'Are you Sure you want to Remove it?')):
                self.ButtonVendorAdd['state']    = ACTIVE
                self.ButtonVendorEdit['state']   = DISABLED
                self.ButtonVendorRemove['state'] = DISABLED
                self.ButtonVendorOK['state']     = DISABLED
                self.ButtonVendorCancel['state'] = DISABLED                
                VendorID = self.VendorIDFrameEntry.get()
                VendorName = self.VendorNameFrameEntry.get()
                #print (VendorVendorID)
                if (VendorID == "UNKNOWN"):
                    mbox.showerror(master=self.VendorFrame,title='Vendor',
                        message = '*** The Vendor ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM Device_Model
                            WHERE Vendor_ID = '%s'
                          """ % (VendorID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.VendorFrame,title='Vendor',
                                    message = '*** The Vendor ID you entered CANNOT be Removed, \r\n Becuase is used in a Device Model ***')
                            self.on_Vendor_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Vendor
                                WHERE Vendor_ID = '%s'
                              """ % (VendorID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Vendor WHERE Vendor_ID = '%s'" % (VendorID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.VendorFrame,title='Vendor',
                                        message = '*** The Vendor ID you entered was Removed ***')
                                self.on_Vendor_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.VendorFrame,title='Vendor',
                                        message = '*** The Vendor ID you entered was NOT Removed ***')
                                self.on_Vendor_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.VendorFrame,title='Vendor',
                            message = '*** The Vendor ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.VendorFrame,title='Vendor',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Vendor_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Vendor','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonVendorAdd['state']    = DISABLED
            self.ButtonVendorEdit['state']   = DISABLED
            self.ButtonVendorRemove['state'] = DISABLED
            self.ButtonVendorOK['state']     = DISABLED 
            self.ButtonVendorCancel['state'] = DISABLED 

            VendorID = self.VendorIDFrameEntry.get()
            VendorName = self.VendorNameFrameEntry.get()
            VendorTel = self.VendorTelFrameEntry.get()
            VendorContract = self.VendorContractFrameEntry.get()
            VendorPOC = self.VendorPOCFrameEntry.get()
            VendorTelPOC = self.VendorTelPOCFrameEntry.get()
            VendorEmailPOC = self.VendorEmailPOCFrameEntry.get()
            
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(VendorID) > 0) and (len(VendorName) > 0)):
                    sql = """
                            SELECT * FROM Vendor
                            WHERE Vendor_ID = '%s'
                          """ % (VendorID)
                    if (self.db.Execute(sql)):
                        self.ButtonVendorOK['state']     = ACTIVE 
                        self.ButtonVendorCancel['state'] = ACTIVE

                        mbox.showerror(master=self.VendorFrame,title='Vendor',
                        message = '*** The Vendor ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Vendor(Vendor_ID, Vendor_Name, Tel_Support, Contract_no, Point_of_contact, \
                                                  Tel_Point_of_Contact,Email_Point_of_Contact,Executed_by_UserID) \
                                           VALUES('%s', '%s', '%s', '%s','%s','%s','%s','%s')" % \
                                   (VendorID, VendorName, VendorTel,VendorContract,VendorPOC,VendorTelPOC,VendorEmailPOC,self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            mbox.showwarning(master=self.VendorFrame,title='Vendor',
                                message = '*** The Vendor ID you entered was Added ***')
                            self.on_Vendor_Table_Refresh()
                        else:
                            mbox.showerror(master=self.VendorFrame,title='Vendor',
                                message = '*** The Vendor ID you entered was NOT Added ***')
                            self.on_Vendor_Table_Refresh()
                else:
                    self.ButtonVendorOK['state']     = ACTIVE 
                    self.ButtonVendorCancel['state'] = ACTIVE
                    mbox.showerror(master=self.VendorFrame,title='Vendor',
                        message = '*** The Vendor ID and Vendor Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Vendor
                        WHERE Vendor_ID = '%s'
                      """ % (VendorID)
                if (len(VendorName) == 0):
                    self.ButtonVendorOK['state']     = ACTIVE 
                    self.ButtonVendorCancel['state'] = ACTIVE
                    self.ButtonVendorTypeAdd['state'] = ACTIVE
                    self.ButtonVendorTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.VendorFrame,title='Vendor',
                        message = '*** The Vendor Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):

                        sql = "UPDATE Vendor SET Vendor_Name = '%s', Tel_Support = '%s', Contract_no = '%s', \
                                                 Point_of_contact = '%s', Tel_Point_of_Contact  = '%s', Email_Point_of_Contact = '%s',\
                                                 Executed_by_UserID = '%s' WHERE Vendor_ID = '%s'" % (VendorName, VendorTel, VendorContract,
                                                VendorPOC, VendorTelPOC, VendorEmailPOC, self.Username, VendorID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.VendorFrame,title='Vendor',
                                message = '*** The Vendor ID you entered was Updated ***')
                            self.on_Vendor_Table_Refresh()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.VendorFrame,title='Vendor',
                                message = '*** The Vendor ID you entered was NOT Upadted ***')
                            self.on_Vendor_Table_Refresh()               
                    else:
                        mbox.showerror(master=self.VendorFrame,title='Vendor',
                        message = '*** The Vendor ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.VendorFrame,title='Vendor',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Vendor_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Vendor','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonVendorAdd['state']    = ACTIVE
        self.ButtonVendorEdit['state']   = DISABLED
        self.ButtonVendorRemove['state'] = DISABLED
        self.ButtonVendorOK['state']     = DISABLED
        self.ButtonVendorCancel['state'] = DISABLED

        self.VendorIDFrameEntry['state'] = 'normal'
        self.VendorIDFrameEntry.delete(0,END)
        self.VendorIDFrameEntry['state'] = 'readonly'
            
        self.VendorNameFrameEntry['state'] = 'normal'
        self.VendorNameFrameEntry.delete(0,END)
        self.VendorNameFrameEntry['state'] = 'readonly'

        self.VendorTelFrameEntry['state'] = 'normal'
        self.VendorTelFrameEntry.delete(0,END)
        self.VendorTelFrameEntry['state'] = 'readonly'       

        self.VendorContractFrameEntry['state'] = 'normal'
        self.VendorContractFrameEntry.delete(0,END)
        self.VendorContractFrameEntry['state'] = 'readonly'

        self.VendorPOCFrameEntry['state'] = 'normal'
        self.VendorPOCFrameEntry.delete(0,END)
        self.VendorPOCFrameEntry['state'] = 'readonly'

        self.VendorTelPOCFrameEntry['state'] = 'normal'
        self.VendorTelPOCFrameEntry.delete(0,END)
        self.VendorTelPOCFrameEntry['state'] = 'readonly'

        self.VendorEmailPOCFrameEntry['state'] = 'normal'
        self.VendorEmailPOCFrameEntry.delete(0,END)
        self.VendorEmailPOCFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_Vendor_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.VendorTreeview.focus()
        dic = self.VendorTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonVendorAdd['state']    = ACTIVE
            self.ButtonVendorEdit['state']   = ACTIVE
            self.ButtonVendorRemove['state'] = ACTIVE
            self.ButtonVendorOK['state']     = DISABLED
            self.ButtonVendorCancel['state'] = DISABLED

            self.VendorIDFrameEntry['state'] = 'normal'
            self.VendorIDFrameEntry.delete(0,END)
            self.VendorIDFrameEntry.insert(0,self.VendorIDArray[curItem])
            self.VendorIDFrameEntry['state'] = 'readonly'
            
            self.VendorNameFrameEntry['state'] = 'normal'
            self.VendorNameFrameEntry.delete(0,END)
            self.VendorNameFrameEntry.insert(0,self.VendorNameArray[curItem])
            self.VendorNameFrameEntry['state'] = 'readonly'

            self.VendorTelFrameEntry['state'] = 'normal'
            self.VendorTelFrameEntry.delete(0,END)
            self.VendorTelFrameEntry.insert(0,self.VendorTelArray[curItem])
            self.VendorTelFrameEntry['state'] = 'readonly'

            self.VendorContractFrameEntry['state'] = 'normal'
            self.VendorContractFrameEntry.delete(0,END)
            self.VendorContractFrameEntry.insert(0,self.VendorContractArray[curItem])
            self.VendorContractFrameEntry['state'] = 'readonly'

            self.VendorPOCFrameEntry['state'] = 'normal'
            self.VendorPOCFrameEntry.delete(0,END)
            self.VendorPOCFrameEntry.insert(0,self.VendorPOCArray[curItem])
            self.VendorPOCFrameEntry['state'] = 'readonly'
            
            self.VendorTelPOCFrameEntry['state'] = 'normal'
            self.VendorTelPOCFrameEntry.delete(0,END)
            self.VendorTelPOCFrameEntry.insert(0,self.VendorTelPOCArray[curItem])
            self.VendorTelPOCFrameEntry['state'] = 'readonly'

            self.VendorEmailPOCFrameEntry['state'] = 'normal'
            self.VendorEmailPOCFrameEntry.delete(0,END)
            self.VendorEmailPOCFrameEntry.insert(0,self.VendorEmailPOCArray[curItem])
            self.VendorEmailPOCFrameEntry['state'] = 'readonly'
            


    def Display_Vendor_Window(self):
        global VendorLabel
        if not self.VendorWindowExist:
            self.VendorWindowExist = True
            # Set up the Window
            self.VendorWindow = Tk()
            self.VendorWindow.geometry('1300x650+350+70')
            self.VendorWindow.title("Vendors")
            self.VendorWindow.protocol("WM_DELETE_WINDOW", self.on_VendorWindow_quit)
            self.VendorWindow.call('tk', 'scaling', self.Windows_Scaling)
            VendorLabel = Label(self.VendorWindow,text="Helvetica", font=("Helvetica", 19))
            VendorLabel["text"] = "Vendors"
            VendorLabel.pack()
                    
            # Setup Frame
            self.VendorFrame = Frame(self.VendorWindow)
            self.VendorFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.VendorFrame.rowconfigure(0, weight=1)
            self.VendorFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    VendorIDFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorIDFrameLabel["text"] = "Vendor ID:"
                    VendorIDFrameLabel.place(x=10, y=10)                    
                    self.VendorIDFrameEntry = Entry(self.VendorFrame)
                    self.VendorIDFrameEntry['width']=40
                    self.VendorIDFrameEntry.place(x=160, y=10)
                    self.VendorIDFrameEntry['state'] = 'readonly'
                    
                    VendorNameFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorNameFrameLabel["text"] = "Vendor Name:"
                    VendorNameFrameLabel.place(x=10, y=40)
                    self.VendorNameFrameEntry = Entry(self.VendorFrame)
                    self.VendorNameFrameEntry['width']=40
                    self.VendorNameFrameEntry.place(x=160, y=40)
                    self.VendorNameFrameEntry['state'] = 'readonly'                    

                    VendorTelFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorTelFrameLabel["text"] = "Vendor Tel. No:"
                    VendorTelFrameLabel.place(x=10, y=70)
                    self.VendorTelFrameEntry = Entry(self.VendorFrame)
                    self.VendorTelFrameEntry['width']=50
                    self.VendorTelFrameEntry.place(x=160, y=70)
                    self.VendorTelFrameEntry['state'] = 'readonly'

                    VendorContractFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorContractFrameLabel["text"] = "Vendor Contract No:"
                    VendorContractFrameLabel.place(x=10, y=100)
                    self.VendorContractFrameEntry = Entry(self.VendorFrame)
                    self.VendorContractFrameEntry['width']=50
                    self.VendorContractFrameEntry.place(x=160, y=100)
                    self.VendorContractFrameEntry['state'] = 'readonly'

                    VendorPOCFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorPOCFrameLabel["text"] = "Point Of Contact:"
                    VendorPOCFrameLabel.place(x=10, y=130)
                    self.VendorPOCFrameEntry = Entry(self.VendorFrame)
                    self.VendorPOCFrameEntry['width']=50
                    self.VendorPOCFrameEntry.place(x=160, y=130)
                    self.VendorPOCFrameEntry['state'] = 'readonly'

                    VendorTelPOCFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorTelPOCFrameLabel["text"] = "Tel. Point Of Contact:"
                    VendorTelPOCFrameLabel.place(x=10, y=160)
                    self.VendorTelPOCFrameEntry = Entry(self.VendorFrame)
                    self.VendorTelPOCFrameEntry['width']=50
                    self.VendorTelPOCFrameEntry.place(x=160, y=160)
                    self.VendorTelPOCFrameEntry['state'] = 'readonly'

                    VendorEmailPOCFrameLabel = Label(self.VendorFrame,text="Helvetica", font=("Helvetica", 10))
                    VendorEmailPOCFrameLabel["text"] = "E-Mail Point Of Contact:"
                    VendorEmailPOCFrameLabel.place(x=10, y=190)
                    self.VendorEmailPOCFrameEntry = Entry(self.VendorFrame)
                    self.VendorEmailPOCFrameEntry['width']=50
                    self.VendorEmailPOCFrameEntry.place(x=160, y=190)
                    self.VendorEmailPOCFrameEntry['state'] = 'readonly'
                   
                    # Setup Buttons
                    self.ButtonVendorAdd = Button(self.VendorFrame, text = 'Add', command = self.Call_Button_Vendor_Add, state=DISABLED)
                    self.ButtonVendorAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonVendorAdd['state'] = ACTIVE

                    self.ButtonVendorEdit = Button(self.VendorFrame, text = 'Edit', command = self.Call_Button_Vendor_Edit, state=DISABLED)
                    self.ButtonVendorEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonVendorRemove = Button(self.VendorFrame, text = 'Remove', command = self.Call_Button_Vendor_Remove, state=DISABLED)
                    self.ButtonVendorRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonVendorOK = Button(self.VendorFrame, text = 'OK / UPDATE', command = self.Call_Button_Vendor_OK, state=DISABLED)
                    self.ButtonVendorOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonVendorCancel = Button(self.VendorFrame, text = 'Cancel', command = self.Call_Button_Vendor_Cancel, state=DISABLED)
                    self.ButtonVendorCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.VendorFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For Vendor Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.VendorTreeviewDataColumns = ('Vendor ID','Name','Tel No','Contract No','Point Of Contact','Tel. No.','E-Mail Address','Added/Updated By')
                    
                    self.VendorTreeview = ttk.Treeview(self.VendorFrame,columns=self.VendorTreeviewDataColumns, height=17) # <--- Make sure the frame is correct !!
                    self.VendorTreeviewysb = Scrollbar(self.VendorFrame,orient=VERTICAL, command=self.VendorTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.VendorTreeviewxsb = Scrollbar(self.VendorFrame,orient=HORIZONTAL, command=self.VendorTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.VendorTreeview['yscroll'] = self.VendorTreeviewysb.set
                    self.VendorTreeview['xscroll'] = self.VendorTreeviewxsb.set

                    # setup headings and column 
                    self.VendorTreeview.heading('#0', text='Item No.',    anchor=W) # E for East and W for West
                    self.VendorTreeview.heading('#1', text='Vendor ID',   anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Vendor ID', False)) # E for East and W for West
                    self.VendorTreeview.heading('#2', text='Name',        anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Name', False)) 
                    self.VendorTreeview.heading('#3', text='Tel No',      anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Tel No', False)) 
                    self.VendorTreeview.heading('#4', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Contract No', False))
                    self.VendorTreeview.heading('#5', text='Point Of Contact', anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Contract No', False))
                    self.VendorTreeview.heading('#6', text='Tel. No.', anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Tel. No.', False))
                    self.VendorTreeview.heading('#7', text='E-Mail Address', anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'E-Mail Address', False))
                    self.VendorTreeview.heading('#8', text='Added/Updated By',    anchor=W,command=lambda: self.treeview_sort_column(self.VendorTreeview, 'Added/Updated By', False)) 
                    
                    self.VendorTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.VendorTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.VendorTreeview.column('#2', stretch=1, width=10)
                    self.VendorTreeview.column('#3', stretch=1, width=10)
                    self.VendorTreeview.column('#4', stretch=1, width=10)
                    self.VendorTreeview.column('#5', stretch=1, width=10)
                    self.VendorTreeview.column('#6', stretch=1, width=10)
                    self.VendorTreeview.column('#7', stretch=1, width=10)
                    self.VendorTreeview.column('#8', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.VendorTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.VendorTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.VendorTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.VendorTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.VendorTreeview.bind('<ButtonRelease-1>', self.on_Vendor_Tree_select_click) # When Select the Tree
                    #self.VendorTreeview.bind("<Double-1>", self.On_Vendor_Tree_Refresh)         

                    # SQL Querry to the Vendor Table
                    self.data_ready = False
                    self.on_Vendor_Table_Refresh()

                    #------------------ MENU <BEGIN> ----------------------------------------------------------
                    menubar = Menu(self.VendorWindow)
                    filemenu = Menu(menubar, tearoff=0)
                    menubar.add_cascade(label="File", menu=filemenu)
                    if Is_SaveAs_Available:
                        filemenu.add_command(label="Save As", command=self.Call_Save_As)
                    filemenu.add_separator()
                    filemenu.add_command(label="Exit", command=self.on_VendorWindow_quit)
                    self.VendorWindow.config(menu=menubar)
                    #------------------ MENU <END> ----------------------------------------------------------
                        
                    self.VendorWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.VendorFrame,title='Vendor',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Vendor Class....:")
    Vendor = Class_Vendor("BV",Windows_Scaling)
    Vendor.Display_Vendor_Window()


if __name__ == '__main__':
    Main()

