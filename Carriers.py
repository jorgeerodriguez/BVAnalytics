#==========================================================================
# Program: Carrier.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: April-22-2018
# Date Last Modified: April-22-2018
# Summary: This is Class to for the Carrier
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
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
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
    Parameter = ['Carrier','OPEN Window']    
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
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_Carrier:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.CarrierWindowExist = False
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
                Tab0 = "Carrier List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.CarrierTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                while (i < len(self.CarrierIDArray)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.CarrierIDArray[i],
                            self.CarrierNameArray[i],
                            self.CarrierTelArray[i],
                            self.CarrierContractArray[i],
                            self.CarrierPOCArray[i],
                            self.CarrierTelPOCArray[i],
                            self.CarrierEmailPOCArray[i],
                            self.CarrierWebArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    #Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.CarrierFrame,title='Carrier List',
                            message = '!!! The File was saved !!!')
                    self.CarrierWindow.title("Carrier List                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Carrier','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.CarrierFrame,title='Carrier List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def on_CarrierWindow_quit(self):
        if (self.CarrierWindowExist):
            self.CarrierWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Carrier','CLOSE Window']    
                Logging.Log(Parameter) 
            self.CarrierWindow.destroy()

    def on_Carrier_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():

            self.CarrierIDFrameEntry['state'] = 'readonly'
            self.CarrierNameFrameEntry['state'] = 'readonly'
            self.CarrierTelFrameEntry['state'] = 'readonly'
            self.CarrierContractFrameEntry['state'] = 'readonly'
            self.CarrierPOCFrameEntry['state'] = 'readonly'
            self.CarrierTelPOCFrameEntry['state'] = 'readonly'
            self.CarrierEmailPOCFrameEntry['state'] = 'readonly'
            self.CarrierWebFrameEntry['state'] = 'readonly'
            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.CarrierTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.CarrierTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                    Carrier_ID              CHAR(20) NOT NULL PRIMARY KEY,
                    Description             CHAR(100),
                    Point_Of_Contact        CHAR(50),                    
                    Telephone_NO            CHAR(50),
                    E_Mail_Address          CHAR(100),
                    Web_Site                CHAR(100),
                    Tech_Support_No         CHAR(50),
                    Master_Contract_No      CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
            '''
            # SQL Querry to the Carrier Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            # SQL Querry to the Carrier Table
            self.data_ready = False
            sql = """
                        SELECT * FROM Carrier
                        ORDER BY Description ASC
                  """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.CarrierIDArray = []
                self.CarrierNameArray = []
                self.CarrierTelArray = []
                self.CarrierContractArray = []
                self.CarrierPOCArray = []
                self.CarrierTelPOCArray = []
                self.CarrierEmailPOCArray = []
                self.CarrierWebArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                '''
                    0 Carrier_ID              CHAR(20) NOT NULL PRIMARY KEY,
                    1 Description             CHAR(100),
                    2 Point_Of_Contact        CHAR(50),                    
                    3 Telephone_NO            CHAR(50),
                    4 E_Mail_Address          CHAR(100),
                    5 Web_Site                CHAR(100),
                    6 Tech_Support_No         CHAR(50),
                    7 Master_Contract_No      CHAR(50),
                    8 Executed_by_UserID      CHAR(20))"""
                '''
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CarrierIDArray.append(self.db.results[i][0].strip())
                    self.CarrierNameArray.append(self.db.results[i][1].strip())
                    if (self.db.results[i][2] == None):
                        self.CarrierPOCArray.append("")
                    else:    
                        self.CarrierPOCArray.append(self.db.results[i][2].strip())
                    if (self.db.results[i][3] == None):
                        self.CarrierTelPOCArray.append("")
                    else:    
                        self.CarrierTelPOCArray.append(self.db.results[i][3].strip())
                    if (self.db.results[i][4] == None):
                        self.CarrierEmailPOCArray.append("")
                    else:    
                        self.CarrierEmailPOCArray.append(self.db.results[i][4].strip())                            
                    if (self.db.results[i][5] == None):
                        self.CarrierWebArray.append("")
                    else:    
                        self.CarrierWebArray.append(self.db.results[i][5].strip())                            
                    if (self.db.results[i][6] == None):
                        self.CarrierTelArray.append("")
                    else:
                        self.CarrierTelArray.append(self.db.results[i][6].strip())
                    if (self.db.results[i][7] == None):
                        self.CarrierContractArray.append("")
                    else:
                        self.CarrierContractArray.append(self.db.results[i][7].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][8].strip())
                    i = i + 1
                i = 0
                while (i < len(self.CarrierIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            self.CarrierIDArray[i],
                            self.CarrierNameArray[i],
                            self.CarrierTelArray[i],
                            self.CarrierContractArray[i],                
                            self.CarrierPOCArray[i],
                            self.CarrierTelPOCArray[i],
                            self.CarrierEmailPOCArray[i],
                            self.CarrierWebArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]                    
                    self.CarrierTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=self.CarrierFrame,title='Carrier',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_Carrier_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Carrier','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.CarrierIDFrameEntry['state'] = 'normal'
        self.CarrierIDFrameEntry.delete(0,END)
        self.CarrierIDFrameEntry['state'] = 'readonly'
        self.CarrierIDFrameEntry['state'] = 'normal'
            
        self.CarrierNameFrameEntry['state'] = 'normal'
        self.CarrierNameFrameEntry.delete(0,END)
        self.CarrierNameFrameEntry['state'] = 'readonly'
        self.CarrierNameFrameEntry['state'] = 'normal'

        self.CarrierTelFrameEntry['state'] = 'normal'
        self.CarrierTelFrameEntry.delete(0,END)
        self.CarrierTelFrameEntry['state'] = 'readonly'
        self.CarrierTelFrameEntry['state'] = 'normal'

        self.CarrierContractFrameEntry['state'] = 'normal'
        self.CarrierContractFrameEntry.delete(0,END)
        self.CarrierContractFrameEntry['state'] = 'readonly'
        self.CarrierContractFrameEntry['state'] = 'normal'

        self.CarrierPOCFrameEntry['state'] = 'normal'
        self.CarrierPOCFrameEntry.delete(0,END)
        self.CarrierPOCFrameEntry['state'] = 'readonly'
        self.CarrierPOCFrameEntry['state'] = 'normal'

        self.CarrierTelPOCFrameEntry['state'] = 'normal'
        self.CarrierTelPOCFrameEntry.delete(0,END)
        self.CarrierTelPOCFrameEntry['state'] = 'readonly'
        self.CarrierTelPOCFrameEntry['state'] = 'normal'

        self.CarrierEmailPOCFrameEntry['state'] = 'normal'
        self.CarrierEmailPOCFrameEntry.delete(0,END)
        self.CarrierEmailPOCFrameEntry['state'] = 'readonly'
        self.CarrierEmailPOCFrameEntry['state'] = 'normal'

        self.CarrierWebFrameEntry['state'] = 'normal'
        self.CarrierWebFrameEntry.delete(0,END)
        self.CarrierWebFrameEntry['state'] = 'readonly'
        self.CarrierWebFrameEntry['state'] = 'normal'
        
        self.ButtonCarrierAdd['state']    = DISABLED
        self.ButtonCarrierEdit['state']   = DISABLED
        self.ButtonCarrierRemove['state'] = DISABLED
        self.ButtonCarrierOK['state']     = ACTIVE
        self.ButtonCarrierCancel['state'] = ACTIVE


    def Call_Button_Carrier_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Carrier','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.CarrierIDFrameEntry['state'] = 'readonly'
        self.CarrierNameFrameEntry['state'] = 'normal'
        self.CarrierTelFrameEntry['state'] = 'normal'
        self.CarrierContractFrameEntry['state'] = 'normal'
        self.CarrierPOCFrameEntry['state'] = 'normal'
        self.CarrierTelPOCFrameEntry['state'] = 'normal'
        self.CarrierEmailPOCFrameEntry['state'] = 'normal'
        self.CarrierWebFrameEntry['state'] = 'normal'

        self.ButtonCarrierAdd['state']    = DISABLED
        self.ButtonCarrierEdit['state']   = DISABLED
        self.ButtonCarrierRemove['state'] = DISABLED
        self.ButtonCarrierOK['state']     = ACTIVE
        self.ButtonCarrierCancel['state'] = ACTIVE


    def Call_Button_Carrier_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Carrier','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.CarrierFrame,title='Carrier',message = 'Are you Sure you want to Remove it?')):
                self.ButtonCarrierAdd['state']    = ACTIVE
                self.ButtonCarrierEdit['state']   = DISABLED
                self.ButtonCarrierRemove['state'] = DISABLED
                self.ButtonCarrierOK['state']     = DISABLED
                self.ButtonCarrierCancel['state'] = DISABLED                
                CarrierID = self.CarrierIDFrameEntry.get()
                CarrierName = self.CarrierNameFrameEntry.get()
                #print (CarrierCarrierID)
                if (CarrierID == "UNKNOWN"):
                    mbox.showerror(master=self.CarrierFrame,title='Carrier',
                        message = '*** The Carrier ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM Circuits
                            WHERE Carrier_ID = '%s'
                          """ % (CarrierID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                                    message = '*** The Carrier ID you entered CANNOT be Removed, \r\n Becuase is used in Circuits ***')
                            self.on_Carrier_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Carrier
                                WHERE Carrier_ID = '%s'
                              """ % (CarrierID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Carrier WHERE Carrier_ID = '%s'" % (CarrierID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.CarrierFrame,title='Carrier',
                                        message = '*** The Carrier ID you entered was Removed ***')
                                self.on_Carrier_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.CarrierFrame,title='Carrier',
                                        message = '*** The Carrier ID you entered was NOT Removed ***')
                                self.on_Carrier_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                            message = '*** The Carrier ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Carrier_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Carrier','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonCarrierAdd['state']    = DISABLED
            self.ButtonCarrierEdit['state']   = DISABLED
            self.ButtonCarrierRemove['state'] = DISABLED
            self.ButtonCarrierOK['state']     = DISABLED 
            self.ButtonCarrierCancel['state'] = DISABLED 

            CarrierID = self.CarrierIDFrameEntry.get()
            CarrierName = self.CarrierNameFrameEntry.get()
            CarrierTel = self.CarrierTelFrameEntry.get()
            CarrierContract = self.CarrierContractFrameEntry.get()
            CarrierPOC = self.CarrierPOCFrameEntry.get()
            CarrierTelPOC = self.CarrierTelPOCFrameEntry.get()
            CarrierEmailPOC = self.CarrierEmailPOCFrameEntry.get()
            CarrierWeb = self.CarrierWebFrameEntry.get()
            
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(CarrierID) > 0) and (len(CarrierName) > 0)):
                    sql = """
                            SELECT * FROM Carrier
                            WHERE Carrier_ID = '%s'
                          """ % (CarrierID)
                    if (self.db.Execute(sql)):
                        self.ButtonCarrierOK['state']     = ACTIVE 
                        self.ButtonCarrierCancel['state'] = ACTIVE
                        mbox.showerror(master=self.CarrierFrame,title='Carrier',
                        message = '*** The Carrier ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Carrier(Carrier_ID, Description, Tech_Support_No, Master_Contract_No, Point_Of_Contact, \
                                                  Telephone_NO, E_Mail_Address, Web_Site, Executed_by_UserID) \
                                           VALUES('%s', '%s', '%s', '%s','%s','%s','%s','%s','%s')" % \
                                   (CarrierID, CarrierName, CarrierTel,CarrierContract,CarrierPOC,CarrierTelPOC,CarrierEmailPOC,CarrierWeb,self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            mbox.showwarning(master=self.CarrierFrame,title='Carrier',
                                message = '*** The Carrier ID you entered was Added ***')
                            self.on_Carrier_Table_Refresh()
                        else:
                            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                                message = '*** The Carrier ID you entered was NOT Added ***')
                            self.on_Carrier_Table_Refresh()
                else:
                    self.ButtonCarrierOK['state']     = ACTIVE 
                    self.ButtonCarrierCancel['state'] = ACTIVE
                    mbox.showerror(master=self.CarrierFrame,title='Carrier',
                        message = '*** The Carrier ID and Carrier Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Carrier
                        WHERE Carrier_ID = '%s'
                      """ % (CarrierID)
                if (len(CarrierName) == 0):
                    self.ButtonCarrierOK['state']     = ACTIVE 
                    self.ButtonCarrierCancel['state'] = ACTIVE
                    self.ButtonCarrierTypeAdd['state'] = ACTIVE
                    self.ButtonCarrierTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.CarrierFrame,title='Carrier',
                        message = '*** The Carrier Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):
                        '''
                            Carrier_ID              CHAR(20) NOT NULL PRIMARY KEY,
                            Description             CHAR(100),
                            Point_Of_Contact        CHAR(50),                    
                            Telephone_NO            CHAR(50),
                            E_Mail_Address          CHAR(100),
                            Web_Site                CHAR(100),
                            Tech_Support_No         CHAR(50),
                            Master_Contract_No      CHAR(50),
                            Executed_by_UserID      CHAR(20))"""
                        '''

                        sql = "UPDATE Carrier SET Description = '%s', Tech_Support_No = '%s', Master_Contract_No = '%s', \
                                                 Point_Of_Contact = '%s', Telephone_NO  = '%s', E_Mail_Address = '%s',\
                                                 Web_Site = '%s', Executed_by_UserID = '%s' WHERE Carrier_ID = '%s'" % (CarrierName, CarrierTel, CarrierContract,
                                                CarrierPOC, CarrierTelPOC, CarrierEmailPOC, CarrierWeb, self.Username, CarrierID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.CarrierFrame,title='Carrier',
                                message = '*** The Carrier ID you entered was Updated ***')
                            self.on_Carrier_Table_Refresh()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                                message = '*** The Carrier ID you entered was NOT Upadted ***')
                            self.on_Carrier_Table_Refresh()               
                    else:
                        mbox.showerror(master=self.CarrierFrame,title='Carrier',
                        message = '*** The Carrier ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CarrierFrame,title='Carrier',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Carrier_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Carrier','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonCarrierAdd['state']    = ACTIVE
        self.ButtonCarrierEdit['state']   = DISABLED
        self.ButtonCarrierRemove['state'] = DISABLED
        self.ButtonCarrierOK['state']     = DISABLED
        self.ButtonCarrierCancel['state'] = DISABLED

        self.CarrierIDFrameEntry['state'] = 'normal'
        self.CarrierIDFrameEntry.delete(0,END)
        self.CarrierIDFrameEntry['state'] = 'readonly'
            
        self.CarrierNameFrameEntry['state'] = 'normal'
        self.CarrierNameFrameEntry.delete(0,END)
        self.CarrierNameFrameEntry['state'] = 'readonly'

        self.CarrierTelFrameEntry['state'] = 'normal'
        self.CarrierTelFrameEntry.delete(0,END)
        self.CarrierTelFrameEntry['state'] = 'readonly'       

        self.CarrierContractFrameEntry['state'] = 'normal'
        self.CarrierContractFrameEntry.delete(0,END)
        self.CarrierContractFrameEntry['state'] = 'readonly'

        self.CarrierPOCFrameEntry['state'] = 'normal'
        self.CarrierPOCFrameEntry.delete(0,END)
        self.CarrierPOCFrameEntry['state'] = 'readonly'

        self.CarrierTelPOCFrameEntry['state'] = 'normal'
        self.CarrierTelPOCFrameEntry.delete(0,END)
        self.CarrierTelPOCFrameEntry['state'] = 'readonly'

        self.CarrierEmailPOCFrameEntry['state'] = 'normal'
        self.CarrierEmailPOCFrameEntry.delete(0,END)
        self.CarrierEmailPOCFrameEntry['state'] = 'readonly'

        self.CarrierWebFrameEntry['state'] = 'normal'
        self.CarrierWebFrameEntry.delete(0,END)
        self.CarrierWebFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_Carrier_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.CarrierTreeview.focus()
        dic = self.CarrierTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonCarrierAdd['state']    = ACTIVE
            self.ButtonCarrierEdit['state']   = ACTIVE
            self.ButtonCarrierRemove['state'] = ACTIVE
            self.ButtonCarrierOK['state']     = DISABLED
            self.ButtonCarrierCancel['state'] = DISABLED

            self.CarrierIDFrameEntry['state'] = 'normal'
            self.CarrierIDFrameEntry.delete(0,END)
            self.CarrierIDFrameEntry.insert(0,self.CarrierIDArray[curItem])
            self.CarrierIDFrameEntry['state'] = 'readonly'
            
            self.CarrierNameFrameEntry['state'] = 'normal'
            self.CarrierNameFrameEntry.delete(0,END)
            self.CarrierNameFrameEntry.insert(0,self.CarrierNameArray[curItem])
            self.CarrierNameFrameEntry['state'] = 'readonly'

            self.CarrierTelFrameEntry['state'] = 'normal'
            self.CarrierTelFrameEntry.delete(0,END)
            self.CarrierTelFrameEntry.insert(0,self.CarrierTelArray[curItem])
            self.CarrierTelFrameEntry['state'] = 'readonly'

            self.CarrierContractFrameEntry['state'] = 'normal'
            self.CarrierContractFrameEntry.delete(0,END)
            self.CarrierContractFrameEntry.insert(0,self.CarrierContractArray[curItem])
            self.CarrierContractFrameEntry['state'] = 'readonly'

            self.CarrierPOCFrameEntry['state'] = 'normal'
            self.CarrierPOCFrameEntry.delete(0,END)
            self.CarrierPOCFrameEntry.insert(0,self.CarrierPOCArray[curItem])
            self.CarrierPOCFrameEntry['state'] = 'readonly'
            
            self.CarrierTelPOCFrameEntry['state'] = 'normal'
            self.CarrierTelPOCFrameEntry.delete(0,END)
            self.CarrierTelPOCFrameEntry.insert(0,self.CarrierTelPOCArray[curItem])
            self.CarrierTelPOCFrameEntry['state'] = 'readonly'

            self.CarrierEmailPOCFrameEntry['state'] = 'normal'
            self.CarrierEmailPOCFrameEntry.delete(0,END)
            self.CarrierEmailPOCFrameEntry.insert(0,self.CarrierEmailPOCArray[curItem])
            self.CarrierEmailPOCFrameEntry['state'] = 'readonly'
            
            self.CarrierWebFrameEntry['state'] = 'normal'
            self.CarrierWebFrameEntry.delete(0,END)
            self.CarrierWebFrameEntry.insert(0,self.CarrierWebArray[curItem])
            self.CarrierWebFrameEntry['state'] = 'readonly'


    def Display_Carrier_Window(self):
        global CarrierLabel
        if not self.CarrierWindowExist:
            self.CarrierWindowExist = True
            # Set up the Window
            self.CarrierWindow = Tk()
            self.CarrierWindow.geometry('1300x650+350+70')
            self.CarrierWindow.title("Carriers")
            self.CarrierWindow.protocol("WM_DELETE_WINDOW", self.on_CarrierWindow_quit)
            self.CarrierWindow.call('tk', 'scaling', self.Windows_Scaling)
            CarrierLabel = Label(self.CarrierWindow,text="Helvetica", font=("Helvetica", 19))
            CarrierLabel["text"] = "Carriers"
            CarrierLabel.pack()
                    
            # Setup Frame
            self.CarrierFrame = Frame(self.CarrierWindow)
            self.CarrierFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.CarrierFrame.rowconfigure(0, weight=1)
            self.CarrierFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    CarrierIDFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierIDFrameLabel["text"] = "Carrier ID:"
                    CarrierIDFrameLabel.place(x=10, y=10)                    
                    self.CarrierIDFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierIDFrameEntry['width']=40
                    self.CarrierIDFrameEntry.place(x=160, y=10)
                    self.CarrierIDFrameEntry['state'] = 'readonly'
                    
                    CarrierNameFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierNameFrameLabel["text"] = "Carrier Name:"
                    CarrierNameFrameLabel.place(x=10, y=40)
                    self.CarrierNameFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierNameFrameEntry['width']=40
                    self.CarrierNameFrameEntry.place(x=160, y=40)
                    self.CarrierNameFrameEntry['state'] = 'readonly'                    

                    CarrierTelFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierTelFrameLabel["text"] = "Carrier Tel. No:"
                    CarrierTelFrameLabel.place(x=10, y=70)
                    self.CarrierTelFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierTelFrameEntry['width']=50
                    self.CarrierTelFrameEntry.place(x=160, y=70)
                    self.CarrierTelFrameEntry['state'] = 'readonly'

                    CarrierContractFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierContractFrameLabel["text"] = "Carrier Contract No:"
                    CarrierContractFrameLabel.place(x=10, y=100)
                    self.CarrierContractFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierContractFrameEntry['width']=50
                    self.CarrierContractFrameEntry.place(x=160, y=100)
                    self.CarrierContractFrameEntry['state'] = 'readonly'

                    CarrierPOCFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierPOCFrameLabel["text"] = "Point Of Contact:"
                    CarrierPOCFrameLabel.place(x=10, y=130)
                    self.CarrierPOCFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierPOCFrameEntry['width']=50
                    self.CarrierPOCFrameEntry.place(x=160, y=130)
                    self.CarrierPOCFrameEntry['state'] = 'readonly'

                    CarrierTelPOCFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierTelPOCFrameLabel["text"] = "Tel. Point Of Contact:"
                    CarrierTelPOCFrameLabel.place(x=10, y=160)
                    self.CarrierTelPOCFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierTelPOCFrameEntry['width']=50
                    self.CarrierTelPOCFrameEntry.place(x=160, y=160)
                    self.CarrierTelPOCFrameEntry['state'] = 'readonly'

                    CarrierEmailPOCFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierEmailPOCFrameLabel["text"] = "E-Mail Point Of Contact:"
                    CarrierEmailPOCFrameLabel.place(x=10, y=190)
                    self.CarrierEmailPOCFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierEmailPOCFrameEntry['width']=50
                    self.CarrierEmailPOCFrameEntry.place(x=160, y=190)
                    self.CarrierEmailPOCFrameEntry['state'] = 'readonly'

                    CarrierWebFrameLabel = Label(self.CarrierFrame,text="Helvetica", font=("Helvetica", 10))
                    CarrierWebFrameLabel["text"] = "Carrier URL/WEB Site:"
                    CarrierWebFrameLabel.place(x=10, y=220)
                    self.CarrierWebFrameEntry = Entry(self.CarrierFrame)
                    self.CarrierWebFrameEntry['width']=50
                    self.CarrierWebFrameEntry.place(x=160, y=220)
                    self.CarrierWebFrameEntry['state'] = 'readonly'
                   
                    # Setup Buttons
                    self.ButtonCarrierAdd = Button(self.CarrierFrame, text = 'Add', command = self.Call_Button_Carrier_Add, state=DISABLED)
                    self.ButtonCarrierAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonCarrierAdd['state'] = ACTIVE

                    self.ButtonCarrierEdit = Button(self.CarrierFrame, text = 'Edit', command = self.Call_Button_Carrier_Edit, state=DISABLED)
                    self.ButtonCarrierEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonCarrierRemove = Button(self.CarrierFrame, text = 'Remove', command = self.Call_Button_Carrier_Remove, state=DISABLED)
                    self.ButtonCarrierRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonCarrierOK = Button(self.CarrierFrame, text = 'OK / UPDATE', command = self.Call_Button_Carrier_OK, state=DISABLED)
                    self.ButtonCarrierOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonCarrierCancel = Button(self.CarrierFrame, text = 'Cancel', command = self.Call_Button_Carrier_Cancel, state=DISABLED)
                    self.ButtonCarrierCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.CarrierFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For Carrier Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.CarrierTreeviewDataColumns = ('Carrier ID','Name','Tel No','Contract No','Point Of Contact','Tel. No.','E-Mail Address','URL','Added/Updated By')
                    
                    self.CarrierTreeview = ttk.Treeview(self.CarrierFrame,columns=self.CarrierTreeviewDataColumns, height=16) # <--- Make sure the frame is correct !!
                    self.CarrierTreeviewysb = Scrollbar(self.CarrierFrame,orient=VERTICAL, command=self.CarrierTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.CarrierTreeviewxsb = Scrollbar(self.CarrierFrame,orient=HORIZONTAL, command=self.CarrierTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.CarrierTreeview['yscroll'] = self.CarrierTreeviewysb.set
                    self.CarrierTreeview['xscroll'] = self.CarrierTreeviewxsb.set

                    # setup headings and column 
                    self.CarrierTreeview.heading('#0', text='Item No.',    anchor=W) # E for East and W for West
                    self.CarrierTreeview.heading('#1', text='Carrier ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Carrier ID', False)) # E for East and W for West
                    self.CarrierTreeview.heading('#2', text='Name',        anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Name', False)) 
                    self.CarrierTreeview.heading('#3', text='Tel No',      anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Tel No', False)) 
                    self.CarrierTreeview.heading('#4', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Contract No', False))
                    self.CarrierTreeview.heading('#5', text='Point Of Contact', anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Contract No', False))
                    self.CarrierTreeview.heading('#6', text='Tel. No.', anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Tel. No.', False))
                    self.CarrierTreeview.heading('#7', text='E-Mail Address', anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'E-Mail Address', False))
                    self.CarrierTreeview.heading('#8', text='URL',    anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'URL', False)) 
                    self.CarrierTreeview.heading('#9', text='Added/Updated By',    anchor=W,command=lambda: self.treeview_sort_column(self.CarrierTreeview, 'Added/Updated By', False)) 
                    
                    self.CarrierTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.CarrierTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.CarrierTreeview.column('#2', stretch=1, width=10)
                    self.CarrierTreeview.column('#3', stretch=1, width=10)
                    self.CarrierTreeview.column('#4', stretch=1, width=10)
                    self.CarrierTreeview.column('#5', stretch=1, width=10)
                    self.CarrierTreeview.column('#6', stretch=1, width=10)
                    self.CarrierTreeview.column('#7', stretch=1, width=10)
                    self.CarrierTreeview.column('#8', stretch=1, width=10)
                    self.CarrierTreeview.column('#9', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.CarrierTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.CarrierTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.CarrierTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.CarrierTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.CarrierTreeview.bind('<ButtonRelease-1>', self.on_Carrier_Tree_select_click) # When Select the Tree
                    #self.CarrierTreeview.bind("<Double-1>", self.On_Carrier_Tree_Refresh)         

                    # SQL Querry to the Carrier Table
                    self.data_ready = False
                    self.on_Carrier_Table_Refresh()

                    #------------------ MENU <BEGIN> ----------------------------------------------------------
                    menubar = Menu(self.CarrierWindow)
                    filemenu = Menu(menubar, tearoff=0)
                    menubar.add_cascade(label="File", menu=filemenu)
                    if Is_SaveAs_Available:
                        filemenu.add_command(label="Save As", command=self.Call_Save_As)
                    filemenu.add_separator()
                    filemenu.add_command(label="Exit", command=self.on_CarrierWindow_quit)
                    self.CarrierWindow.config(menu=menubar)
                    #------------------ MENU <END> ----------------------------------------------------------
                        
                    self.CarrierWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.CarrierFrame,title='Carrier',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Carrier Class....:")
    Carrier = Class_Carrier("BV",Windows_Scaling)
    Carrier.Display_Carrier_Window()


if __name__ == '__main__':
    Main()

