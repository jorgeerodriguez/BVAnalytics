#==========================================================================
# Program: NetworkTemplatesBreakOut.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: May-21-2018
# Date Last Modified: May-24-2018
# Summary: This is Class to for the NetworkTemplatesBreakOut
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

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Network Templates Breakout','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

try:
    from NetworkTemplates import *
    Is_NetworkTemplates_Available = True
    NetworkTemplates = Class_NetworkTemplates(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO NetworkTemplates Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_NetworkTemplates_Available = False  

try:
    from NetworkTemplatesBreakoutSettings import *
    Is_NetworkTemplatesBreakoutSettings_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO NetworkTemplates Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_NetworkTemplatesBreakoutSettings_Available = False  


#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_NetworkTemplatesBreakOut:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.NetworkTemplatesBreakOutWindowExist = False
        self.NetworkTemplatesBreakOutSettingsWindowExist = False # <- This one controls the Settign Sub-Window
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

#-------------------------------------------------- Network Template Window <BEGIN> -------------------------------------------
    def Display_NetworkTemplate_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','DeviceType Window']    
            Logging.Log(Parameter)
        NetworkTemplates.Display_NetworkTemplates_Window()

    def on_Network_Template_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','Template Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():
            '''
            Template_ID             CHAR(100) NOT NULL PRIMARY KEY,
            Template_Name           CHAR(100) NOT NULL,
            Base_Network_Slash      INT,
            Split_in_Slash          INT,
            Executed_by_UserID      CHAR(20)
            '''

            self.TemplateTable_TemplateIDArray = []
            self.TemplateTable_TemplateNameArray = []
            self.TemplateTable_BaseNetworkSlashkArray = []
            self.TemplateTable_SplitInSlashArray = []
            self.TemplateTable_AddedByArray = []

            # SQL Querry to the NetworkTemplatesBreakOut Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """ SELECT * FROM Network_Templates ORDER BY Template_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                while (i < len(self.db.results)):
                    num = i + 1
                    self.TemplateTable_TemplateIDArray.append(self.db.results[i][0].strip())
                    self.TemplateTable_TemplateNameArray.append(self.db.results[i][1].strip())
                    self.TemplateTable_BaseNetworkSlashkArray.append(self.db.results[i][2])
                    self.TemplateTable_SplitInSlashArray.append(self.db.results[i][3])
                    self.TemplateTable_AddedByArray.append(self.db.results[i][4].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxTemplateID['values'] = self.TemplateTable_TemplateNameArray
                if (len(self.TemplateTable_TemplateNameArray)== 0):
                    self.ComboBoxTemplateID['state'] = DISABLED
                else:
                    self.ComboBoxTemplateID.current(0)
                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.current(self.TemplateTable_BaseNetworkSlashkArray[0])
                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.current(self.TemplateTable_SplitInSlashArray[0])
                    self.ComboBoxTemplateID['state'] = 'readonly'
                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['state'] = 'disabled'
                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['state'] = 'disabled'
            else:
                mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                message = 'No Templates Records found')
        else:
            mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

#-------------------------------------------------- Network Template Window <END> -------------------------------------------

    def on_Network_Template_combo_changed(self,event):
        self.ComboBoxTemplateID.current()
        self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.current(self.TemplateTable_BaseNetworkSlashkArray[self.ComboBoxTemplateID.current()])
        self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.current(self.TemplateTable_SplitInSlashArray[self.ComboBoxTemplateID.current()])

        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.NetworkTemplatesBreakOutTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.NetworkTemplatesBreakOutTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------
 
        if (Is_NetworkTemplates_Available and self.NetworkTemplatesBreakOutSettingsWindowExist):
            self.NetworkTemplatesBreakoutSettings.on_NetworkTemplatesBreakoutSettingsWindow_quit()
            self.NetworkTemplatesBreakOutSettingsWindowExist = False

        if (self.TemplateTable_BaseNetworkSlashkArray[self.ComboBoxTemplateID.current()] == 0 and
           self.TemplateTable_SplitInSlashArray[self.ComboBoxTemplateID.current()] == 0):

            self.ButtonNetworkTemplatesBreakOutSplit['state']  = DISABLED
            self.ButtonNetworkTemplatesBreakOutEdit['state']   = ACTIVE
            self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
            self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
            self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED
        else:
            if self.db.Connect():
                # SQL Querry to the NetworkTemplatesBreakOut Table
                sql = """
                        SELECT * FROM Network_Templates_Break_Out
                        WHERE Template_ID = '%s'
                        ORDER BY Sequence_1 ASC, Sequence_2 ASC
                      """ % (self.TemplateTable_TemplateIDArray[self.ComboBoxTemplateID.current()])
                
                if (self.db.Execute(sql)):
                    self.TemplateIDSEQ1SEQ2Array = []
                    self.TemplateIDArray = []
                    self.Sequence1Array = []
                    self.Sequence2Array = []
                    self.SplitinSlashArray = []
                    self.VLANL3TypeArray = []
                    self.VLAN_L3_NameArray = []
                    self.DescriptionArray = []
                    self.DHCPArray = []
                    self.SplitMoreArray = []
                    self.SplitMoreSlashArray = []
                    self.ExecutedbyUserIDArray = []
                    self.results = []
                    self.results2 = []
                    self.data_ready = True   
                    i = 0
                    '''

                        0  Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                        1  Template_ID              CHAR(100) NOT NULL,
                        2  Sequence_1               INT,
                        3  Sequence_2               INT,
                        4  Split_in_Slash           INT,
                        5  VLAN_L3_Type             CHAR(20),
                        6  VLAN_L3_Name             CHAR(50),
                        7  Description              CHAR(100),
                        8  DHCP                     CHAR(10),
                        9  Split_More               CHAR(10),
                        10 Split_More_Slash         INT,
                        11 Executed_by_UserID       CHAR(20)
                    '''

                    while (i < len(self.db.results)):
                        num = i + 1
                        self.TemplateIDSEQ1SEQ2Array.append(self.db.results[i][0].strip())
                        self.TemplateIDArray.append(self.db.results[i][1].strip())
                        self.Sequence1Array.append(self.db.results[i][2])
                        self.Sequence2Array.append(self.db.results[i][3])
                        self.SplitinSlashArray.append(self.db.results[i][4])
                        self.VLANL3TypeArray.append(self.db.results[i][5].strip())
                        if (self.db.results[i][6] == None):
                            self.VLAN_L3_NameArray.append("")
                        else:
                            self.VLAN_L3_NameArray.append(self.db.results[i][6].strip())
                        if (self.db.results[i][7] == None):
                            self.DescriptionArray.append("")
                        else:
                            self.DescriptionArray.append(self.db.results[i][7].strip())
                        self.DHCPArray.append(self.db.results[i][8].strip())
                        self.SplitMoreArray.append(self.db.results[i][9].strip())
                        self.SplitMoreSlashArray.append(self.db.results[i][10])
                        self.ExecutedbyUserIDArray.append(self.db.results[i][11].strip())
                        i = i + 1
                       # ('Sequence','Slash','Type','Description','DHCP','Split More','In Slash','Added By')
                    i = 0
                    while (i < len(self.TemplateIDSEQ1SEQ2Array)):
                        num = i + 1
                        
                        tags = self.VLANL3TypeArray[i]
                        if (self.Sequence2Array[i] > 0):
                            tags = "SummarySplit"
                        item = [
                                (str(self.Sequence1Array[i])+"."+str(self.Sequence2Array[i])),
                                self.SplitinSlashArray[i],
                                self.VLANL3TypeArray[i],
                                self.VLAN_L3_NameArray[i],
                                self.DescriptionArray[i],
                                self.DHCPArray[i],
                                self.SplitMoreArray[i],
                                self.SplitMoreSlashArray[i],
                                self.ExecutedbyUserIDArray[i]
                                ]
                        self.results.append(item)
                        item2 = [
                                self.TemplateIDSEQ1SEQ2Array[i],
                                self.TemplateIDArray[i],
                                self.Sequence1Array[i],
                                self.Sequence2Array[i],
                                self.SplitinSlashArray[i],
                                self.VLANL3TypeArray[i],
                                self.VLAN_L3_NameArray[i],
                                self.DescriptionArray[i],
                                self.DHCPArray[i],
                                self.SplitMoreArray[i],
                                self.SplitMoreSlashArray[i],
                                self.ExecutedbyUserIDArray[i]
                                ]
                        self.results2.append(item2)
                        self.NetworkTemplatesBreakOutTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                        i = i + 1
                        
                    self.ButtonNetworkTemplatesBreakOutSplit['state']    = ACTIVE
                    self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
                    self.ButtonNetworkTemplatesBreakOutRemove['state'] = ACTIVE
                    self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
                    self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
                    self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
                    self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED

                else:
                    mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates BrakOut',
                    message = 'No Records found')
                    self.data_ready = False
                    #------------------------------- Deleting Tree View <BEGIN> --------
                    x = self.NetworkTemplatesBreakOutTreeview.get_children()
                    if x != '()': # checks if there is something in the first row
                        for child in x:
                            #print (child)
                            self.NetworkTemplatesBreakOutTreeview.delete(child)
                    #------------------------------- Deleting Tree View <END> --------
                    self.ButtonNetworkTemplatesBreakOutSplit['state']  = ACTIVE
                    self.ButtonNetworkTemplatesBreakOutEdit['state']   = ACTIVE
                    self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
                    self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
                    self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
                    self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
                    self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED

                    # need to add the Combobox here                 
            else:
                mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates BreakOut',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

                
    def on_NetworkTemplatesBreakOutWindow_quit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','CLOSE Window']    
            Logging.Log(Parameter)

        if (self.NetworkTemplatesBreakOutWindowExist):
            self.NetworkTemplatesBreakOutWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Network Templates Breakout','CLOSE Window']    
                Logging.Log(Parameter)
            if (Is_NetworkTemplates_Available and self.NetworkTemplatesBreakOutSettingsWindowExist):
                self.NetworkTemplatesBreakoutSettings.on_NetworkTemplatesBreakoutSettingsWindow_quit()
            self.NetworkTemplatesBreakOutWindow.destroy()


    def on_NetworkTemplatesBreakOut_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','Template Refresh Button']    
            Logging.Log(Parameter)
        self.on_Network_Template_combo_changed("event")


    def Call_Button_NetworkTemplatesBreakOut_ChangeSettings(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','Change Settings']    
            Logging.Log(Parameter)
  
        if Is_NetworkTemplates_Available:
            if not self.NetworkTemplatesBreakOutSettingsWindowExist:
                self.NetworkTemplatesBreakOutSettingsWindowExist = True
                Treeview = self.NetworkTemplatesBreakOutTreeview
                self.NetworkTemplatesBreakoutSettings = Class_NetworkTemplatesBreakoutSettings(ODBC_DSN_name,Windows_Scaling,self.results2,Treeview,self.index)
                self.NetworkTemplatesBreakoutSettings.Display_NetworkTemplatesBreakoutSettings_Window()   


    def Call_Button_NetworkTemplatesBreakOut_Split_Summary(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','SPLIT Summary Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.ComboBoxTemplateID['state'] = DISABLED

        self.ButtonNetworkTemplatesBreakOutSplit['state']    = DISABLED
        self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
        self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED
        
        self.TemplateIDSEQ1SEQ2 = self.TemplateIDSEQ1SEQ2Array[self.index]
        self.TemplateID = self.TemplateIDArray[self.index]
        self.Sequence1 = self.Sequence1Array[self.index]
        self.Sequence2 = self.Sequence2Array[self.index]
        self.SplitinSlash = self.SplitinSlashArray[self.index]
        self.VLANL3Type = self.VLANL3TypeArray[self.index]
        self.VLAN_L3_Name = self.VLAN_L3_NameArray[self.index]
        self.Description = self.DescriptionArray[self.index]
        self.DHCP = self.DHCPArray[self.index]
        self.SplitMore = self.SplitMoreArray[self.index]
        self.SplitMoreSlash = self.SplitMoreSlashArray[self.index]
        self.ExecutedbyUserID = self.ExecutedbyUserIDArray[self.index]

        Base_Network = self.SplitinSlash
        Brake_it_into = self.SplitMoreSlash
        
        if (Brake_it_into < Base_Network ):            
                mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                message = '*** The Divided in Slash needs to be Greater than the Base Network Slash ***')
        else:
            Total = 2 ** (Brake_it_into - Base_Network)
            seq1 = self.Sequence1
            seq2 = 1
            #print (Total)
            self.progress['maximum'] = Total
            self.progress['value'] = 0
            while (seq2 <= Total):
                TemplateID = self.TemplateTable_TemplateIDArray[self.ComboBoxTemplateID.current()]
                PrimaryKey = TemplateID+"-"+str(seq1)+"-"+str(seq2)
                sql = """
                        SELECT * FROM Network_Templates_Break_Out
                        WHERE Template_ID_SEQ1_SEQ2 = '%s'
                      """ % (PrimaryKey)
                if not (self.db.Execute(sql)):
                    '''
                    Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                    Template_ID              CHAR(100) NOT NULL,
                    Sequence_1               INT,
                    Sequence_2               INT,
                    Split_in_Slash           INT,
                    VLAN_L3_Type             CHAR(20), # VLAN | Layer3 | Summary
                    VLAN_L3_Name             CHAR(50),
                    Description              CHAR(100),
                    DHCP                     CHAR(10),
                    Split_More               CHAR(10),
                    Split_More_Slash         INT,
                    Executed_by_UserID       CHAR(20) )"""
                    '''
                    sql = "INSERT INTO Network_Templates_Break_Out(Template_ID_SEQ1_SEQ2, Template_ID, Sequence_1, Sequence_2, Split_in_Slash, VLAN_L3_Type, \
                                   VLAN_L3_Name, Description, DHCP, Split_More, Split_More_Slash,Executed_by_UserID) \
                                   VALUES ('%s','%s','%d', '%d','%d','%s','%s','%s','%s','%s','%d','%s')" % \
                                   (PrimaryKey, TemplateID, seq1, seq2, Brake_it_into, "VLAN", "0","NO Description", "NO", "NO",0,self.Username )
                    self.db.Add_Move_Change_Data(sql) # test a limitation on 65535
                seq2 = seq2 + 1
                self.progress['value'] = seq2
            #print(seq1)
            self.on_Network_Template_combo_changed("event")                
        self.ComboBoxTemplateID['state'] = 'readonly'

    
    def Call_Button_NetworkTemplatesBreakOut_Split(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','SPLIT Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.ComboBoxTemplateID['state'] = DISABLED
        #self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['state'] = 'readonly'
        #self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['state'] = 'readonly'

        self.ButtonNetworkTemplatesBreakOutSplit['state']    = ACTIVE
        self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutRemove['state'] = ACTIVE
        self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
        self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED
        
        self.ComboBoxTemplateID.current()
        Base_Network = self.TemplateTable_BaseNetworkSlashkArray[self.ComboBoxTemplateID.current()]
        Brake_it_into = self.TemplateTable_SplitInSlashArray[self.ComboBoxTemplateID.current()]
        if (Brake_it_into < Base_Network ):            
                mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                message = '*** The Divided in Slash needs to be Greater than the Base Network Slash ***')
        else:
            Total = 2 ** (Brake_it_into - Base_Network)
            seq1 = 1
            seq2 = 0
            #print (Total)
            self.progress['maximum'] = Total
            self.progress['value'] = 0
            while (seq1 <= Total):
                TemplateID = self.TemplateTable_TemplateIDArray[self.ComboBoxTemplateID.current()]
                PrimaryKey = TemplateID+"-"+str(seq1)+"-"+str(seq2)
                sql = """
                        SELECT * FROM Network_Templates_Break_Out
                        WHERE Template_ID_SEQ1_SEQ2 = '%s'
                      """ % (PrimaryKey)
                if not (self.db.Execute(sql)):
                    '''
                    Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                    Template_ID              CHAR(100) NOT NULL,
                    Sequence_1               INT,
                    Sequence_2               INT,
                    Split_in_Slash           INT,
                    VLAN_L3_Type             CHAR(20), # VLAN | Layer3 | Summary
                    VLAN_L3_Name             CHAR(50),
                    Description              CHAR(100),
                    DHCP                     CHAR(10),
                    Split_More               CHAR(10),
                    Split_More_Slash         INT,
                    Executed_by_UserID       CHAR(20) )"""
                    '''
                    sql = "INSERT INTO Network_Templates_Break_Out(Template_ID_SEQ1_SEQ2, Template_ID, Sequence_1, Sequence_2, Split_in_Slash, VLAN_L3_Type, \
                                   VLAN_L3_Name, Description, DHCP, Split_More, Split_More_Slash,Executed_by_UserID) \
                                   VALUES ('%s','%s','%d', '%d','%d','%s','%s','%s','%s','%s','%d','%s')" % \
                                   (PrimaryKey, TemplateID, seq1, seq2, Brake_it_into, "VLAN", "0","NO Description", "NO", "NO",0,self.Username )
                    self.db.Add_Move_Change_Data(sql) # test a limitation on 65535
                seq1 = seq1 + 1
                self.progress['value'] = seq1
            #print(seq1)
            self.on_Network_Template_combo_changed("event")                
        self.ComboBoxTemplateID['state'] = 'readonly'



    def Call_Button_NetworkTemplatesBreakOut_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'

        self.ComboBoxTemplateID['state'] = DISABLED
        
        self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['state'] = 'readonly'
        self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['state'] = 'readonly'

        self.ButtonNetworkTemplatesBreakOutSplit['state']    = DISABLED
        self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutOK['state']     = ACTIVE
        self.ButtonNetworkTemplatesBreakOutCancel['state'] = ACTIVE
        self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED


    def Call_Button_NetworkTemplatesBreakOut_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove' 
            if (mbox.askyesnocancel(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',message = 'Are you Sure you want to Remove it?')):
                PrimaryKey = self.TemplateTable_TemplateIDArray[self.ComboBoxTemplateID.current()]
                sql = "DELETE FROM Network_Templates_Break_out WHERE Template_ID = '%s'" % (PrimaryKey)
                if (self.db.Add_Move_Change_Data(sql)):
                    mbox.showwarning(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                        message = '*** The Network Template ID you entered was Removed ***')
                    self.on_NetworkTemplatesBreakOut_Table_Refresh()
                else:
                    mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                        message = '*** The Network Template ID you entered was NOT Removed ***')
                    self.on_NetworkTemplatesBreakOut_Table_Refresh()               
        else:
            mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='NetworkTemplatesBreakOut',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_NetworkTemplatesBreakOut_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ComboBoxTemplateID['state'] = 'readonly'

            self.ButtonNetworkTemplatesBreakOutSplit['state']    = DISABLED
            self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
            self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED 
            self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
            self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED
            
            PrimaryKey = self.TemplateTable_TemplateIDArray[self.ComboBoxTemplateID.current()]
            #print (PrimaryKey)
            if (self.Selection == 'edit'):
                #Base_Network = self.TemplateTable_BaseNetworkSlashkArray[self.ComboBoxTemplateID.current()]
                #Brake_it_into = self.TemplateTable_SplitInSlashArray[self.ComboBoxTemplateID.current()]
                #if (Brake_it_into < Base_Network ):            
                #    mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                #    message = '*** The Divided in Slash needs to be Greater than the Base Network Slash ***')

                if (self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.current() < self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.current()):
                    mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates BreakOut',
                    message = '*** The Divided in Slash needs to be Greater than the Base Network Slash ***')
                    self.ButtonNetworkTemplatesBreakOutOK['state']     = ACTIVE 
                    self.ButtonNetworkTemplatesBreakOutCancel['state'] = ACTIVE
                else:
                    sql = """
                            SELECT * FROM Network_Templates
                            WHERE Template_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)): 
                        sql = "UPDATE Network_Templates SET Base_Network_Slash = '%d', Split_in_Slash = '%d', \
                               Executed_by_UserID = '%s' WHERE Template_ID = '%s'" % (self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.current(),
                                                            self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.current(),self.Username,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                                message = '*** The Template ID you entered was Updated ***')
                            self.on_Network_Template_Table_Refresh()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                                message = '*** The Template ID you entered was NOT Upadted ***')
                            self.on_Network_Template_Table_Refresh()
                    else:
                        mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                        message = '*** The Template ID you try to Edit Does not exist Anymore ***')
                        self.on_Network_Template_Table_Refresh()
        else:
            mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Templates Breakout',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_NetworkTemplatesBreakOut_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Templates Breakout','CANCEL Button']    
            Logging.Log(Parameter)
        self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['state'] = 'disabled'
        self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['state'] = 'disabled'

        self.ButtonNetworkTemplatesBreakOutSplit['state']    = DISABLED
        self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
        self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
        self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = DISABLED
        self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED

        self.ComboBoxTemplateID['state'] = 'readonly'

        self.Selection = 'cancel'
        self.on_Network_Template_combo_changed("event")


    def on_NetworkTemplatesBreakOut_Tree_select_click(self,event):
        #print ("CLICK")
        self.ComboBoxTemplateID['state'] = 'readonly'
        curItem = self.NetworkTemplatesBreakOutTreeview.focus()
        dic = self.NetworkTemplatesBreakOutTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            self.index = curItem
            #print (values)
            #print (values[0])
            #print (values[1])

            self.TemplateIDSEQ1SEQ2 = self.TemplateIDSEQ1SEQ2Array[curItem]
            self.TemplateID = self.TemplateIDArray[curItem]
            self.Sequence1 = self.Sequence1Array[curItem]
            self.Sequence2 = self.Sequence2Array[curItem]
            self.SplitinSlash = self.SplitinSlashArray[curItem]
            self.VLANL3Type = self.VLANL3TypeArray[curItem]
            self.VLAN_L3_Name = self.VLAN_L3_NameArray[curItem]
            self.Description = self.DescriptionArray[curItem]
            self.DHCP = self.DHCPArray[curItem]
            self.SplitMore = self.SplitMoreArray[curItem]
            self.SplitMoreSlash = self.SplitMoreSlashArray[curItem]
            self.ExecutedbyUserID = self.ExecutedbyUserIDArray[curItem]
            
            self.ButtonNetworkTemplatesBreakOutSplit['state']    = DISABLED
            self.ButtonNetworkTemplatesBreakOutEdit['state']   = DISABLED
            self.ButtonNetworkTemplatesBreakOutRemove['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutOK['state']     = DISABLED
            self.ButtonNetworkTemplatesBreakOutCancel['state'] = DISABLED
            self.ButtonNetworkTemplatesBreakOutChangeSettings['state']   = ACTIVE
            #print (self.VLANL3Type)
            #print (self.SplitMore)
            #print (self.SplitMoreSlash)
            #print (self.SplitinSlash)
            if ((self.VLANL3Type == "Summary") and (self.SplitMore == "YES") and (self.SplitMoreSlash > self.SplitinSlash) and self.Sequence2 == 0):
                self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = ACTIVE
            else:
                self.ButtonNetworkTemplatesBreakOutSplitSummary['state'] = DISABLED
                

    def On_NetworkTemplatesBreakOut_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_NetworkTemplatesBreakOut_Table_Refresh()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                TemplateName = self.TemplateTable_TemplateNameArray[self.ComboBoxTemplateID.current()]
                Tab0 = TemplateName
                Tab1 = "NONE"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.NetworkTemplatesBreakOutTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                #ExcellFile.Add_DataToWorksheet(self.NetworkTemplatesBreakOutTreeviewDataColumns,Row,Column,Tab1,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                Row = 3
                '''
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab1,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                '''
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.NetworkTemplatesBreakOutFrame,title='Network Template Breakout List',
                            message = '!!! The File was saved !!!')
                    self.NetworkTemplatesBreakOutWindow.title("Network Templates Provision                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Network Templates Breakout','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='Network Template Breakout List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_NetworkTemplatesBreakOut_Window(self):
        #global NetworkTemplatesBreakOutLabel
 
        if not self.NetworkTemplatesBreakOutWindowExist:
            # Set up the Window
            self.NetworkTemplatesBreakOutWindowExist = True
            self.NetworkTemplatesBreakOutWindow = Tk()
            self.NetworkTemplatesBreakOutWindow.geometry('1000x700+350+70')
            self.NetworkTemplatesBreakOutWindow.title("Network Templates Provision")
            self.NetworkTemplatesBreakOutWindow.protocol("WM_DELETE_WINDOW", self.on_NetworkTemplatesBreakOutWindow_quit)
            self.NetworkTemplatesBreakOutWindow.call('tk', 'scaling', self.Windows_Scaling)
            NetworkTemplatesBreakOutLabel = Label(self.NetworkTemplatesBreakOutWindow,text="Helvetica", font=("Helvetica", 19))
            NetworkTemplatesBreakOutLabel["text"] = "Network Templates Provision"
            NetworkTemplatesBreakOutLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.NetworkTemplatesBreakOutWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_NetworkTemplatesBreakOutWindow_quit)
            self.NetworkTemplatesBreakOutWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.NetworkTemplatesBreakOutFrame = Frame(self.NetworkTemplatesBreakOutWindow)
            self.NetworkTemplatesBreakOutFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.NetworkTemplatesBreakOutFrame.rowconfigure(0, weight=1)
            self.NetworkTemplatesBreakOutFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                    self.slash = []
                    i = 0
                    while (i < 33):
                        self.slash.append("Slash: ["+str(i)+"]")
                        i = i + 1

                     # Setup Buttons
                    if (Is_NetworkTemplates_Available):
                        self.ButtonDeviceTypeAdd = Button(self.NetworkTemplatesBreakOutFrame, text = '+ Template', command = self.Display_NetworkTemplate_Window, state=ACTIVE)
                        self.ButtonDeviceTypeAdd.place(x = 450, y = 7, width=85, height=25)

                        self.ButtonDeviceTypeRefresh = Button(self.NetworkTemplatesBreakOutFrame, text = 'Refresh', command = self.on_Network_Template_Table_Refresh, state=ACTIVE)
                        self.ButtonDeviceTypeRefresh.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonNetworkTemplatesBreakOutSplit = Button(self.NetworkTemplatesBreakOutFrame, text = 'SPLIT', command = self.Call_Button_NetworkTemplatesBreakOut_Split, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutSplit.place(x = 450, y = 37, width=85, height=25)

                    self.ButtonNetworkTemplatesBreakOutEdit = Button(self.NetworkTemplatesBreakOutFrame, text = 'Edit', command = self.Call_Button_NetworkTemplatesBreakOut_Edit, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutEdit.place(x = 550, y = 37, width=75, height=25)

                    self.ButtonNetworkTemplatesBreakOutRemove = Button(self.NetworkTemplatesBreakOutFrame, text = 'Remove', command = self.Call_Button_NetworkTemplatesBreakOut_Remove, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutRemove.place(x = 650, y = 37, width=75, height=25)

                    self.ButtonNetworkTemplatesBreakOutOK = Button(self.NetworkTemplatesBreakOutFrame, text = 'OK / UPDATE', command = self.Call_Button_NetworkTemplatesBreakOut_OK, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutOK.place(x = 750, y = 37, width=100, height=25)

                    self.ButtonNetworkTemplatesBreakOutCancel = Button(self.NetworkTemplatesBreakOutFrame, text = 'Cancel', command = self.Call_Button_NetworkTemplatesBreakOut_Cancel, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutCancel.place(x = 875, y = 37, width=75, height=25)

                    self.ButtonNetworkTemplatesBreakOutChangeSettings = Button(self.NetworkTemplatesBreakOutFrame, text = 'Change Settings', command = self.Call_Button_NetworkTemplatesBreakOut_ChangeSettings, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutChangeSettings.place(x = 450, y = 67, width=100, height=25)

                    self.ButtonNetworkTemplatesBreakOutSplitSummary = Button(self.NetworkTemplatesBreakOutFrame, text = 'SPLIT Summary', command = self.Call_Button_NetworkTemplatesBreakOut_Split_Summary, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakOutSplitSummary.place(x = 575, y = 67, width=100, height=25)


                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.NetworkTemplatesBreakOutFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=97)

                    # Setup Labels
                    NetworkTemplatesBreakOutIDFrameLabel = Label(self.NetworkTemplatesBreakOutFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakOutIDFrameLabel["text"] = "Template:"
                    NetworkTemplatesBreakOutIDFrameLabel.place(x=10, y=10)

                    NetworkTemplatesBreakOutBaseNetworkLabel = Label(self.NetworkTemplatesBreakOutFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakOutBaseNetworkLabel["text"] = "Base Network:"
                    NetworkTemplatesBreakOutBaseNetworkLabel.place(x=10, y=40)

                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork = ttk.Combobox(self.NetworkTemplatesBreakOutFrame, state='disabled', width = 20)
                    #self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.place(x = 140, y = 40)
                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['values'] = self.slash
                    #self.ComboBoxNetworkTemplatesBreakOutBaseNetwork['state'] = 'readonly'
                    self.ComboBoxNetworkTemplatesBreakOutBaseNetwork.current(0)

                    NetworkTemplatesBreakOutDivideNetworkLabel = Label(self.NetworkTemplatesBreakOutFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakOutDivideNetworkLabel["text"] = "Divided in:"
                    NetworkTemplatesBreakOutDivideNetworkLabel.place(x=10, y=70)

                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork = ttk.Combobox(self.NetworkTemplatesBreakOutFrame, state='disabled', width = 20)
                    #self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.place(x = 140, y = 70)
                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['values'] = self.slash
                    #self.ComboBoxNetworkTemplatesBreakOutDivideNetwork['state'] = 'readonly'
                    self.ComboBoxNetworkTemplatesBreakOutDivideNetwork.current(0)
                    
                    
                    #------------------ TREE VIEW For Network Templates Breakout Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.NetworkTemplatesBreakOutTreeviewDataColumns = ('Sequence','Slash','Type','VLAN No.','Description','DHCP','Split More','In Slash','Added By')
                    
                    self.NetworkTemplatesBreakOutTreeview = ttk.Treeview(self.NetworkTemplatesBreakOutFrame,columns=self.NetworkTemplatesBreakOutTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesBreakOutTreeviewysb = Scrollbar(self.NetworkTemplatesBreakOutFrame,orient=VERTICAL, command=self.NetworkTemplatesBreakOutTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesBreakOutTreeviewxsb = Scrollbar(self.NetworkTemplatesBreakOutFrame,orient=HORIZONTAL, command=self.NetworkTemplatesBreakOutTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesBreakOutTreeview['yscroll'] = self.NetworkTemplatesBreakOutTreeviewysb.set
                    self.NetworkTemplatesBreakOutTreeview['xscroll'] = self.NetworkTemplatesBreakOutTreeviewxsb.set

                    # setup headings and column 
                    self.NetworkTemplatesBreakOutTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.NetworkTemplatesBreakOutTreeview.heading('#1', text='Sequence',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Sequence', False)) # E for East and W for West
                    self.NetworkTemplatesBreakOutTreeview.heading('#2', text='Slash',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Slash', False)) # E for East and W for West
                    self.NetworkTemplatesBreakOutTreeview.heading('#3', text='Type',    anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Type', False)) # E for East and W for West
                    self.NetworkTemplatesBreakOutTreeview.heading('#4', text='VLAN No.',  anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'VLAN No.', False)) 
                    self.NetworkTemplatesBreakOutTreeview.heading('#5', text='Description',  anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Description', False)) 
                    self.NetworkTemplatesBreakOutTreeview.heading('#6', text='DHCP',  anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'DHCP', False)) 
                    self.NetworkTemplatesBreakOutTreeview.heading('#7', text='Split More',  anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Split More', False)) 
                    self.NetworkTemplatesBreakOutTreeview.heading('#8', text='In Slash',  anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Slash', False)) 
                    self.NetworkTemplatesBreakOutTreeview.heading('#9', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesBreakOutTreeview, 'Added By', False)) 
                    
                    self.NetworkTemplatesBreakOutTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.NetworkTemplatesBreakOutTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.NetworkTemplatesBreakOutTreeview.column('#2', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#3', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#4', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#5', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#6', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#7', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#8', stretch=1, width=10)
                    self.NetworkTemplatesBreakOutTreeview.column('#9', stretch=1, width=10)

                    self.NetworkTemplatesBreakOutTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.NetworkTemplatesBreakOutTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.NetworkTemplatesBreakOutTreeviewxsb.grid(row=2, column=0, sticky=EW)

                    # create fonts and tags
                    self.NetworkTemplatesBreakOutTreeview.tag_configure('None', font=('Helvetica', 10), background='gray1')
                    self.NetworkTemplatesBreakOutTreeview.tag_configure('Summary', font=('Helvetica', 10), background='tan1')
                    self.NetworkTemplatesBreakOutTreeview.tag_configure('SummarySplit', font=('Helvetica', 8), background='gray90')
                   
                    # Bind the double Click
                    self.NetworkTemplatesBreakOutTreeview.bind('<ButtonRelease-1>', self.on_NetworkTemplatesBreakOut_Tree_select_click) # When Select the Tree
                    self.NetworkTemplatesBreakOutTreeview.bind("<Double-1>", self.On_NetworkTemplatesBreakOut_Tree_Refresh)         

                    #------------------ TREE VIEW For Device Model Country Cost Database <END> -----------------------------------                    

                    # Setup ComboBox
                    self.ComboBoxTemplateID = ttk.Combobox(self.NetworkTemplatesBreakOutFrame, state='readonly', width = 37)
                    self.ComboBoxTemplateID.bind("<<ComboboxSelected>>", self.on_Network_Template_combo_changed)
                    self.ComboBoxTemplateID.place(x = 140, y = 10)
                    self.on_Network_Template_Table_Refresh()
                    self.on_Network_Template_combo_changed("event")
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.NetworkTemplatesBreakOutWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.NetworkTemplatesBreakOutFrame,title='NetworkTemplatesBreakOut',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the NetworkTemplatesBreakOut Class....:")
    NetworkTemplatesBreakOut = Class_NetworkTemplatesBreakOut("BV",Windows_Scaling)
    NetworkTemplatesBreakOut.Display_NetworkTemplatesBreakOut_Window()


if __name__ == '__main__':
    Main()

