#==========================================================================
# Program: NetworkAllocations.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: May-25-2018
# Date Last Modified: May-25-2018
# Summary: This is Class to for the NetworkAllocations
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
    Parameter = ['Network Allocation','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

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
    from IPAddressCalculator import *
    Is_IPAddressCalculator_Available = True
except:
    print ("************************************************************************************************* \n")
    print ("*** NO IPAddressCalculator Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************************* \n")
    Is_IPAddressCalculator_Available = False


#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_NetworkAllocations:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.NetworkAllocationsWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.GetPasswordWindowsExists = False

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


    def on_NetworkAllocationsWindow_quit(self):
        if (self.NetworkAllocationsWindowExist):
            self.NetworkAllocationsWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Network Allocation','CLOSE Window']    
                Logging.Log(Parameter) 
            self.NetworkAllocationsWindow.destroy()

    def on_NetworkAllocations_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.NetworkAllocationsTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.NetworkAllocationsTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                                    0.- Network_Slash            CHAR(50) NOT NULL PRIMARY KEY,
                                    1.- Network                  CHAR(20) NOT NULL,
                                    2.- Slash                    INT,
                                    3.- Subnet_Mask              CHAR(50) NOT NULL,
                                    4.- IPA                      INT,
                                    5.- IPB                      INT,
                                    6.- IPC                      INT,
                                    7.- IPD                      INT,
                                    8.- SMA                      INT,
                                    9.- SMB                      INT,
                                    10.-SMC                      INT,
                                    11.-SMD                      INT,
                                    12.-No_Of_Hosts              INT,
                                    13.-IP_From                  CHAR(20),
                                    14.-IP_To                    CHAR(20),
                                    15.-IP_From_Num              INT,
                                    16.-IP_To_Num                INT,
                                    17.-Assigned                 CHAR(5),
                                    18.-Network_Region_ID        CHAR(20),
                                    19.-Office_Size              CHAR(20) NOT NULL,
                                    20.-Status                   CHAR(20),                        
                                    21.-VLAN_L3_Type             CHAR(20),
                                    22.-VLAN_L3_Name             CHAR(50),
                                    23.-Description              CHAR(100),
                                    24.-DHCP                     CHAR(10),
                                    25.-Allocation_Date          CHAR(20),
                                    26.-Provision_Date           CHAR(20),
                                    27.-Executed_by_UserID       CHAR(20)
                                    28.-Country_ID               CHAR(20),
                                    29.-Region_ID                CHAR(20),
                                    30.-Facility_ID              CHAR(20),
                                    31.-Site_ID                  CHAR(20))"""
            '''
            self.Get_Region()
            
            # SQL Querry to the NetworkAllocations Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """
                        SELECT * FROM Networks
                        ORDER BY IP_From_Num ASC, Slash ASC
                  """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.NetworkTable_Network_Slash_Array = []
                self.NetworkTable_Network_Array = []
                self.NetworkTable_Slash_Array = []
                self.NetworkTable_Subnet_Mask_Array = []
                self.NetworkTable_IPA_Array = []
                self.NetworkTable_IPB_Array = []
                self.NetworkTable_IPC_Array = []
                self.NetworkTable_IPD_Array = []
                self.NetworkTable_SMA_Array = []
                self.NetworkTable_SMB_Array = []
                self.NetworkTable_SMC_Array = []
                self.NetworkTable_SMD_Array = []
                self.NetworkTable_No_Of_Hosts_Array = []
                self.NetworkTable_IP_From_Array = []
                self.NetworkTable_IP_To_Array = []
                self.NetworkTable_IP_From_Num_Array = []
                self.NetworkTable_IP_To_Num_Array = []
                self.NetworkTable_ssigned_Array = []
                self.NetworkTable_Network_Region_ID_Array = []
                self.NetworkTable_Office_Size_Array = []
                self.NetworkTable_Status_Array = []                        
                self.NetworkTable_VLAN_L3_Type_Array = []
                self.NetworkTable_VLAN_L3_Name_Array = []
                self.NetworkTable_Description_Array = []
                self.NetworkTable_DHCP_Array = []
                self.NetworkTable_Allocation_Date_Array = []
                self.NetworkTable_Provision_Date_Array = []
                self.NetworkTable_Executed_by_UserID_Array = []
                self.NetworkTable_Country_ID_Array = []
                self.NetworkTable_Region_ID_Array = []
                self.NetworkTable_Facility_ID_Array = []
                self.NetworkTable_Site_ID_Array = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    '''
                    self.NetworkAllocationsIDArray.append(self.db.results[i][0].strip())
                    self.NetworkAllocationsNameArray.append(self.db.results[i][1].strip())
                    if (self.db.results[i][2] == None):
                        self.RegionIDArray.append("UNKNOWN")
                    else:
                        self.RegionIDArray.append(self.db.results[i][2].strip())
                    '''
                    self.NetworkTable_Network_Slash_Array.append(self.db.results[i][0].strip())
                    self.NetworkTable_Network_Array.append(self.db.results[i][1].strip())
                    self.NetworkTable_Slash_Array.append(self.db.results[i][2])
                    self.NetworkTable_Subnet_Mask_Array.append(self.db.results[i][3].strip())
                    self.NetworkTable_IPA_Array.append(self.db.results[i][4])
                    self.NetworkTable_IPB_Array.append(self.db.results[i][5])
                    self.NetworkTable_IPC_Array.append(self.db.results[i][6])
                    self.NetworkTable_IPD_Array.append(self.db.results[i][7])
                    self.NetworkTable_SMA_Array.append(self.db.results[i][8])
                    self.NetworkTable_SMB_Array.append(self.db.results[i][9])
                    self.NetworkTable_SMC_Array.append(self.db.results[i][10])
                    self.NetworkTable_SMD_Array.append(self.db.results[i][11])
                    self.NetworkTable_No_Of_Hosts_Array.append(self.db.results[i][12])
                    self.NetworkTable_IP_From_Array.append(self.db.results[i][13].strip())
                    self.NetworkTable_IP_To_Array.append(self.db.results[i][14].strip())
                    self.NetworkTable_IP_From_Num_Array.append(self.db.results[i][15])
                    self.NetworkTable_IP_To_Num_Array.append(self.db.results[i][16])
                    self.NetworkTable_ssigned_Array.append(self.db.results[i][17].strip())
                    self.NetworkTable_Network_Region_ID_Array.append(self.db.results[i][18].strip())
                    self.NetworkTable_Office_Size_Array.append(self.db.results[i][19].strip())
                    self.NetworkTable_Status_Array.append(self.db.results[i][20].strip())
                    if (self.db.results[i][21] == None):
                        self.NetworkTable_VLAN_L3_Type_Array.append("")
                    else:    
                        self.NetworkTable_VLAN_L3_Type_Array.append(self.db.results[i][21].strip())
                    if (self.db.results[i][22] == None):
                        self.NetworkTable_VLAN_L3_Name_Array.append("")
                    else:
                        self.NetworkTable_VLAN_L3_Name_Array.append(self.db.results[i][22].strip())
                    self.NetworkTable_Description_Array.append(self.db.results[i][23].strip())
                    if (self.db.results[i][24] == None):
                        self.NetworkTable_DHCP_Array.append("NO")
                    else:    
                        self.NetworkTable_DHCP_Array.append(self.db.results[i][24].strip())
                    self.NetworkTable_Allocation_Date_Array.append(self.db.results[i][25].strip())
                    if (self.db.results[i][26] == None):
                        self.NetworkTable_Provision_Date_Array.append("")
                    else:
                        self.NetworkTable_Provision_Date_Array.append(self.db.results[i][26].strip())
                    self.NetworkTable_Executed_by_UserID_Array.append(self.db.results[i][27].strip())
                    if (self.db.results[i][28] == None):
                         self.NetworkTable_Country_ID_Array.append("")
                    else:
                        self.NetworkTable_Country_ID_Array.append(self.db.results[i][28].strip())
                    if (self.db.results[i][29] == None):
                        self.NetworkTable_Region_ID_Array.append("")
                    else:
                        self.NetworkTable_Region_ID_Array.append(self.db.results[i][29].strip())
                    if (self.db.results[i][30] == None):
                        self.NetworkTable_Facility_ID_Array.append("")
                    else:
                        self.NetworkTable_Facility_ID_Array.append(self.db.results[i][30].strip())
                    if (self.db.results[i][31] == None):
                        self.NetworkTable_Site_ID_Array.append("")
                    else:
                        self.NetworkTable_Site_ID_Array.append(self.db.results[i][31].strip())
                    i = i + 1
                                                            
                self.results = []
                self.results2 = []
                i = 0
                while (i < len(self.NetworkTable_Network_Slash_Array)):
                    num = i + 1
                    tags = "-" # To use in the future
                    self.NetworkAllocationsNetworkTableColumns = ('Network-Slash','Network','Slash','Subnet Mask','IPA','IPB','IPC','IPD','SMA','SMB','SMC','SMD',
                                                                  'No. Of Host','Network From','Network To','Network From Num','Network To Num','Asigned','Network Region',
                                                                  'Office Size','Status','Type','Vlan No.','Description','DHCP','Allocated','Provisioned','Added By',
                                                                  'Country','Region','Facility','Site')
                    if (self.NetworkTable_ssigned_Array[i] == 'YES') and (self.NetworkTable_VLAN_L3_Type_Array[i] != 'Summary'):
                        Network = "     "+self.NetworkTable_Network_Array[i]
                    else:
                        Network = self.NetworkTable_Network_Array[i]
                    item = [
                            Network,
                            #self.NetworkTable_Network_Array[i],
                            self.NetworkTable_Slash_Array[i],
                            self.NetworkTable_Description_Array[i],
                            self.NetworkTable_Allocation_Date_Array[i],
                            self.NetworkTable_Provision_Date_Array[i],
                            self.NetworkTable_Country_ID_Array[i],
                            self.NetworkTable_Network_Region_ID_Array[i],
                            #self.NetworkTable_Region_ID_Array[i],
                            self.NetworkTable_Facility_ID_Array[i],
                            self.NetworkTable_VLAN_L3_Type_Array[i],
                            self.NetworkTable_VLAN_L3_Name_Array[i],
                            self.NetworkTable_DHCP_Array[i],
                            self.NetworkTable_Executed_by_UserID_Array[i]
                            ]
                    self.results.append(item)
                    self.NetworkAllocationsTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    self.NetworkAllocationsTable = [
                            self.NetworkTable_Network_Slash_Array[i],
                            self.NetworkTable_Network_Array[i],
                            self.NetworkTable_Slash_Array[i],
                            self.NetworkTable_Subnet_Mask_Array[i],
                            self.NetworkTable_IPA_Array[i],
                            self.NetworkTable_IPB_Array[i],
                            self.NetworkTable_IPC_Array[i],
                            self.NetworkTable_IPD_Array[i],
                            self.NetworkTable_SMA_Array[i],
                            self.NetworkTable_SMB_Array[i],
                            self.NetworkTable_SMC_Array[i],
                            self.NetworkTable_SMD_Array[i],
                            self.NetworkTable_No_Of_Hosts_Array[i],
                            self.NetworkTable_IP_From_Array[i],
                            self.NetworkTable_IP_To_Array[i],
                            self.NetworkTable_IP_From_Num_Array[i],
                            self.NetworkTable_IP_To_Num_Array[i],
                            self.NetworkTable_ssigned_Array[i],
                            self.NetworkTable_Network_Region_ID_Array[i],
                            self.NetworkTable_Office_Size_Array[i],
                            self.NetworkTable_Status_Array[i],
                            self.NetworkTable_VLAN_L3_Type_Array[i],
                            self.NetworkTable_VLAN_L3_Name_Array[i],
                            self.NetworkTable_Description_Array[i],
                            self.NetworkTable_DHCP_Array[i],
                            self.NetworkTable_Allocation_Date_Array[i],
                            self.NetworkTable_Provision_Date_Array[i],
                            self.NetworkTable_Executed_by_UserID_Array[i],
                            self.NetworkTable_Country_ID_Array[i],
                            self.NetworkTable_Region_ID_Array[i],
                            self.NetworkTable_Facility_ID_Array[i],
                            self.NetworkTable_Site_ID_Array[i]
                        ]
                    self.results2.append(self.NetworkAllocationsTable)
                    i = i + 1
                    self.progress['value'] = i
            else:
                mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                message = 'No Records found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")



    def Call_Button_Region_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','Region Button']    
            Logging.Log(Parameter)
        Region.Display_Region_Window()

    def Call_Button_Region_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','Region Refresh Button']    
            Logging.Log(Parameter)
        self.Get_Region()
        self.ComboBoxNetworkRegionID['state'] = 'readonly'
        self.ComboBoxNetworkRegionID.current(0)

        
    def Call_Button_NetworkAllocations_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.NetworkAllocationsNetworkFrameEntry['state'] = 'normal'
        self.NetworkAllocationsNetworkFrameEntry.delete(0,END)
        self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
        self.NetworkAllocationsNetworkFrameEntry['state'] = 'normal'
            
        self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'normal'
        self.NetworkAllocationsNetworkDescriptionFrameEntry.delete(0,END)
        self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'readonly'
        self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'normal'

        self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'readonly'
        self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'readonly'


        self.NetworkAllocationsFromFrameEntry['state'] = 'normal'
        self.NetworkAllocationsFromFrameEntry.delete(0,END)
        self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'

        self.NetworkAllocationsToFrameEntry['state'] = 'normal'
        self.NetworkAllocationsToFrameEntry.delete(0,END)
        self.NetworkAllocationsToFrameEntry['state'] = 'readonly'

        self.NetworkAllocationsNoHostsFrameEntry['state'] = 'normal'
        self.NetworkAllocationsNoHostsFrameEntry.delete(0,END)
        self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
        
        self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'active'

        self.ComboBoxNetworkRegionID['state'] = 'active'
        self.ComboBoxNetworkRegionID.current(0)

        self.ButtonNetworkAllocationsAdd['state']    = DISABLED
        self.ButtonNetworkAllocationsEdit['state']   = DISABLED
        self.ButtonNetworkAllocationsRemove['state'] = DISABLED
        self.ButtonNetworkAllocationsOK['state']     = ACTIVE
        self.ButtonNetworkAllocationsCancel['state'] = ACTIVE

        self.ButtonRegionTypeAdd['state'] = ACTIVE
        self.ButtonRegionTypeRefresh['state'] = ACTIVE



    def Call_Button_NetworkAllocations_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'

        if (self.assigned == 'YES'):
            self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
                
            self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'normal'

            self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'disabled'

            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'
            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
            
            self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'disabled'
            self.ComboBoxNetworkRegionID['state'] = 'disabled'

            self.ButtonNetworkAllocationsAdd['state']    = DISABLED
            self.ButtonNetworkAllocationsEdit['state']   = DISABLED
            self.ButtonNetworkAllocationsRemove['state'] = DISABLED
            self.ButtonNetworkAllocationsOK['state']     = ACTIVE
            self.ButtonNetworkAllocationsCancel['state'] = ACTIVE

            self.ButtonRegionTypeAdd['state'] = DISABLED
            self.ButtonRegionTypeRefresh['state'] = DISABLED
            
        else:
            self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
                
            self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'normal'

            self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'disabled'

            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'
            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
            
            self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'active'

            self.ComboBoxNetworkRegionID['state'] = 'active'

            self.ButtonNetworkAllocationsAdd['state']    = DISABLED
            self.ButtonNetworkAllocationsEdit['state']   = DISABLED
            self.ButtonNetworkAllocationsRemove['state'] = DISABLED
            self.ButtonNetworkAllocationsOK['state']     = ACTIVE
            self.ButtonNetworkAllocationsCancel['state'] = ACTIVE

            self.ButtonRegionTypeAdd['state'] = ACTIVE
            self.ButtonRegionTypeRefresh['state'] = ACTIVE
       

    #-------------- Using a Password Question to make sure it was the intent to be deleted <Begin> ---------------

    def Remove_Networks_From_DB(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.NetworkAllocationsFrame,title='Network Allocation',message = 'Are you Sure you want to Remove it?')):
                self.ButtonNetworkAllocationsAdd['state']    = ACTIVE
                self.ButtonNetworkAllocationsEdit['state']   = DISABLED
                self.ButtonNetworkAllocationsRemove['state'] = DISABLED
                self.ButtonNetworkAllocationsOK['state']     = DISABLED
                self.ButtonNetworkAllocationsCancel['state'] = DISABLED                
                self.ButtonRegionTypeAdd['state'] = DISABLED
                self.ButtonRegionTypeRefresh['state'] = DISABLED
                #print (NetworkAllocationsRegionID)
                if (self.assigned == 'YES'):
                    mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                        message = '*** The Network CANNOT be Removed since it has been assigned already ***')
                else:
                    PrimaryKey = self.PrimaryKey
                    sql = "DELETE FROM Networks WHERE Network_Slash = '%s'" % (PrimaryKey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        mbox.showwarning(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was Removed ***')
                    else:
                        mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was NOT Removed ***')
                    self.on_NetworkAllocations_Table_Refresh()               
        else:
            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.GetPasswordWindow.destroy()
            self.Remove_Networks_From_DB()
            self.GetPasswordWindowsExists = False
        else:
            mbox.showerror(master=self.GetPasswordWindow,title='Username and Password',
                             message = '*** ERROR *** - Please Enter a Valid Information')
            self.GetPasswordWindow.destroy()
            self.GetPasswordWindowsExists = False
            
    def try_login_Enter(self,event):
        self.try_login()
        
    def on_GetPasswordWindow_quit(self):
        self.GetPasswordWindowsExists = False
        self.GetPasswordWindow.destroy()

    def Get_Usernanme_and_Password(self):
        if not self.GetPasswordWindowsExists:
            self.password = ""
            self.username = ""
            self.GetPasswordWindowsExists = True
            self.GetPasswordWindow = Tk()
            self.GetPasswordWindow.resizable(width=FALSE, height=FALSE)
            self.GetPasswordWindow.protocol("WM_DELETE_WINDOW", self.on_GetPasswordWindow_quit)
            self.GetPasswordWindow.title("Log-In")
            self.GetPasswordWindow.geometry("200x150")
            #Creating the username & password entry boxes
            self.username_text = Label(self.GetPasswordWindow, text="Username:")
            self.username_guess = Entry(self.GetPasswordWindow)
            self.password_text = Label(self.GetPasswordWindow, text="Password:")
            self.password_guess = Entry(self.GetPasswordWindow, show="*")
            self.password_guess.bind('<Return>',self.try_login_Enter)
            self.attempt_login = Button(self.GetPasswordWindow,text="Login", command = self.try_login)                 
            self.username_text.pack()
            self.username_guess.pack()
            self.password_text.pack()
            self.password_guess.pack()
            self.attempt_login.pack()
            self.GetPasswordWindow.mainloop()
            
    def Call_Button_NetworkAllocations_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','REMOVE Button']    
            Logging.Log(Parameter)
        self.Get_Usernanme_and_Password()

    #-------------- Using a Password Question to make sure it was the intent to be deleted <End> ---------------

    def Call_Button_NetworkAllocations_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonNetworkAllocationsAdd['state']    = DISABLED
            self.ButtonNetworkAllocationsEdit['state']   = DISABLED
            self.ButtonNetworkAllocationsRemove['state'] = DISABLED
            self.ButtonNetworkAllocationsOK['state']     = DISABLED 
            self.ButtonNetworkAllocationsCancel['state'] = DISABLED 
            self.ButtonRegionTypeAdd['state'] = DISABLED
            self.ButtonRegionTypeRefresh['state'] = DISABLED

            self.ButtonNetworkAllocationsAdd['state']    = DISABLED
            self.ButtonNetworkAllocationsEdit['state']   = DISABLED
            self.ButtonNetworkAllocationsRemove['state'] = DISABLED
            self.ButtonNetworkAllocationsOK['state']     = ACTIVE
            self.ButtonNetworkAllocationsCancel['state'] = ACTIVE

            self.ButtonRegionTypeAdd['state'] = ACTIVE
            self.ButtonRegionTypeRefresh['state'] = ACTIVE

            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                self.NetworkAllocationsNetworkDescription = self.NetworkAllocationsNetworkDescriptionFrameEntry.get()
                if ((len(self.NetworkAllocationsNetworkDescription) > 0) and (self.Is_Network_Valid)):
                    self.NetworkAllocationsNetwork = self.NetworkAllocationsNetworkFrameEntry.get()
                    self.NetworkAllocationsNetworkDescription = self.NetworkAllocationsNetworkDescriptionFrameEntry.get()
                    self.NetworkAllocationsNetworkSlash = self.ComboBoxNetworkAllocationsNetworkSlash.current()
                    self.NetworkAllocationsNetworkSubnetMask = self.Subnetmask[self.ComboBoxNetworkAllocationsNetworkSlash.current()]
                    self.NetowrkAllocationsSubnetMaskArray = self.NetworkAllocationsNetworkSubnetMask.split('.')
                    self.NetworkAllocationsOfficeSize = self.NetworkAllocationsOfficeSizeValues[self.ComboBoxNetworkAllocationsOfficeSize.current()]
                    self.NetworkAllocationsNetworkRegionID = self.RegionTableRegionIDArray[self.ComboBoxNetworkRegionID.current()]
                    PrimaryKey = self.network_From_Str + "-" + str(self.NetworkAllocationsNetworkSlash)

                    self.now = datetime.datetime.now()
                    self.day = self.now.day
                    self.month = self.now.month
                    self.year = self.now.year
                    self.date = str(self.month) + "/" + str(self.day) + "/" + str(self.year)
                    
                    sql = """
                            SELECT * FROM Networks
                            WHERE Network_Slash = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        self.ButtonNetworkAllocationsOK['state']     = ACTIVE 
                        self.ButtonNetworkAllocationsCancel['state'] = ACTIVE
                        self.ButtonRegionTypeAdd['state'] = ACTIVE
                        self.ButtonRegionTypeRefresh['state'] = ACTIVE

                        mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                        message = '*** The Network you entered already exist ***')
                    else:
                        '''
                                    Network_Slash            CHAR(50) NOT NULL PRIMARY KEY,
                                    Network                  CHAR(20) NOT NULL,
                                    Slash                    INT,
                                    Subnet_Mask              CHAR(50) NOT NULL,
                                    IPA                      INT,
                                    IPB                      INT,
                                    IPC                      INT,
                                    IPD                      INT,
                                    SMA                      INT,
                                    SMB                      INT,
                                    SMC                      INT,
                                    SMD                      INT,
                                    No_Of_Hosts              INT,
                                    IP_From                  CHAR(20),
                                    IP_To                    CHAR(20),
                                    IP_From_Num              INT,
                                    IP_To_Num                INT,
                                    Assigned                 CHAR(5),
                                    Network_Region_ID        CHAR(20),
                                    Office_Size              CHAR(20) NOT NULL,
                                    Status                   CHAR(20),                        
                                    VLAN_L3_Type             CHAR(20),
                                    VLAN_L3_Name             CHAR(50),
                                    Description              CHAR(100),
                                    DHCP                     CHAR(10),
                                    Allocation_Date          CHAR(20),
                                    Provision_Date           CHAR(20),
                                    Executed_by_UserID       CHAR(20),
                                    Country_ID               CHAR(20),
                                    Region_ID                CHAR(20),
                                    Facility_ID              CHAR(20),
                                    Site_ID                  CHAR(20))"""
                        '''
                        sql = "INSERT INTO Networks (Network_Slash, Network, Slash, Subnet_Mask,IPA, IPB, IPC, IPD, SMA, SMB, SMC, SMD, No_Of_Hosts, IP_From, IP_To,\
                                   IP_From_Num, IP_To_Num, Assigned, Network_Region_ID, Office_Size, Status, VLAN_L3_Type, VLAN_L3_Name, Description, \
                                   DHCP, Allocation_Date, Provision_Date, Executed_by_UserID, Country_ID, Region_ID, Facility_ID, Site_ID) \
                                   VALUES ('%s','%s',%d,'%s',%d,%d,%d,%d,%d,%d,%d,%d,%d,'%s','%s',%f,%f,'%s','%s','%s','%s','%s','%s',\
                                   '%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.network_From_Str, int(self.NetworkAllocationsNetworkSlash), self.NetworkAllocationsNetworkSubnetMask,
                                    int(self.Network_Range_Results[0]),int(self.Network_Range_Results[1]),int(self.Network_Range_Results[2]),int(self.Network_Range_Results[3]),
                                    int(self.NetowrkAllocationsSubnetMaskArray[0]),int(self.NetowrkAllocationsSubnetMaskArray[1]),int(self.NetowrkAllocationsSubnetMaskArray[2]),
                                    int(self.NetowrkAllocationsSubnetMaskArray[3]),int(self.No_of_Hosts),self.network_From_Str,self.network_To_Str,self.Network_From_Num,
                                    self.Network_To_Num,'NO',self.NetworkAllocationsNetworkRegionID,self.NetworkAllocationsOfficeSize,'Allocation','Summary','VLAN 0',
                                    self.NetworkAllocationsNetworkDescription,'NO',self.date,'',self.Username,'','','','')
                        #print (sql)
                        #print (self.NetworkAllocationsNetworkSlash)
                        #print (type(self.NetworkAllocationsNetworkSlash))
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was Added ***')
                            self.on_NetworkAllocations_Table_Refresh()
                            
                            self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
                            self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'readonly'
                            self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'disabled'
                            self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'disabled'
                            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'
                            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'
                            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'                            
                            self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'disabled'
                            self.ComboBoxNetworkRegionID['state'] = 'disabled'
                        else:
                            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was NOT Added ***')
                            self.on_NetworkAllocations_Table_Refresh()
                            self.ComboBoxNetworkRegionID['state'] = 'active'
                else:
                    self.ButtonNetworkAllocationsOK['state']     = ACTIVE 
                    self.ButtonNetworkAllocationsCancel['state'] = ACTIVE
                    self.ButtonRegionTypeAdd['state'] = ACTIVE
                    self.ButtonRegionTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                        message = '*** Needs to be a Valida Network and the Description Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                self.NetworkAllocationsNetworkDescription = self.NetworkAllocationsNetworkDescriptionFrameEntry.get()
                if ((len(self.NetworkAllocationsNetworkDescription) == 0)):
                    self.ButtonNetworkAllocationsOK['state']     = ACTIVE 
                    self.ButtonNetworkAllocationsCancel['state'] = ACTIVE
                    self.ButtonRegionTypeAdd['state'] = ACTIVE
                    self.ButtonRegionTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                        message = '*** The Description Cannot be BLANK ***')                     
                else:
                    self.NetworkAllocationsNetworkDescription = self.NetworkAllocationsNetworkDescriptionFrameEntry.get()
                    self.NetworkAllocationsOfficeSize = self.NetworkAllocationsOfficeSizeValues[self.ComboBoxNetworkAllocationsOfficeSize.current()]
                    self.NetworkAllocationsNetworkRegionID = self.RegionTableRegionIDArray[self.ComboBoxNetworkRegionID.current()]
                    PrimaryKey = self.PrimaryKey
                    
                    sql = """
                            SELECT * FROM Networks
                            WHERE Network_Slash = '%s'
                          """ % (PrimaryKey)
                    
                    if (self.db.Execute(sql)):
                        sql = "UPDATE Networks SET Network_Region_ID = '%s', Office_Size = '%s', Description = '%s', Executed_by_UserID = '%s' \
                               WHERE Network_Slash = '%s'" % (self.NetworkAllocationsNetworkRegionID,self.NetworkAllocationsOfficeSize,
                                    self.NetworkAllocationsNetworkDescription,self.Username,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            mbox.showwarning(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was Updated ***')
                        else:
                            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                message = '*** The Network you entered was NOT Upadted ***')
                        self.on_NetworkAllocations_Table_Refresh()               
                    else:
                        mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                        message = '*** The Network you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_NetworkAllocations_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonNetworkAllocationsAdd['state']    = ACTIVE
        self.ButtonNetworkAllocationsEdit['state']   = DISABLED
        self.ButtonNetworkAllocationsRemove['state'] = DISABLED
        self.ButtonNetworkAllocationsOK['state']     = DISABLED
        self.ButtonNetworkAllocationsCancel['state'] = DISABLED
        self.ButtonRegionTypeAdd['state'] = ACTIVE
        self.ButtonRegionTypeRefresh['state'] = DISABLED

        self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
        self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'readonly'
        self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'disabled'
        self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'disabled'
        self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'
        self.NetworkAllocationsToFrameEntry['state'] = 'readonly'
        self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
        self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'disabled'
        self.ComboBoxNetworkRegionID['state'] = 'disabled'
        
        self.Selection = 'cancel'

    def on_NetworkAllocations_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.NetworkAllocationsTreeview.focus()
        dic = self.NetworkAllocationsTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.data_ready = True
            
            self.ButtonNetworkAllocationsAdd['state']    = ACTIVE
            self.ButtonNetworkAllocationsEdit['state']   = ACTIVE
            self.ButtonNetworkAllocationsRemove['state'] = ACTIVE
            self.ButtonNetworkAllocationsOK['state']     = DISABLED
            self.ButtonNetworkAllocationsCancel['state'] = DISABLED
            self.ButtonRegionTypeAdd['state'] = ACTIVE
            self.ButtonRegionTypeRefresh['state'] = DISABLED


            self.NetworkAllocationsNetworkFrameEntry['state'] = 'normal'
            self.NetworkAllocationsNetworkFrameEntry.delete(0,END)
            self.NetworkAllocationsNetworkFrameEntry.insert(0,self.NetworkTable_Network_Array[curItem])
            self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'
                
            self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'normal'
            self.NetworkAllocationsNetworkDescriptionFrameEntry.delete(0,END)
            self.NetworkAllocationsNetworkDescriptionFrameEntry.insert(0,self.NetworkTable_Description_Array[curItem])
            self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'readonly'

            self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsNetworkSlash.current(self.NetworkTable_Slash_Array[curItem])
            self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsNetworkSubnetMask.current(self.NetworkTable_Slash_Array[curItem])


            self.NetworkAllocationsFromFrameEntry['state'] = 'normal'
            self.NetworkAllocationsFromFrameEntry.delete(0,END)
            self.NetworkAllocationsFromFrameEntry.insert(0,self.NetworkTable_IP_From_Array[curItem])
            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'

            self.NetworkAllocationsToFrameEntry['state'] = 'normal'
            self.NetworkAllocationsToFrameEntry.delete(0,END)
            self.NetworkAllocationsToFrameEntry.insert(0,self.NetworkTable_IP_To_Array[curItem])
            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'

            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'normal'
            self.NetworkAllocationsNoHostsFrameEntry.delete(0,END)
            self.NetworkAllocationsNoHostsFrameEntry.insert(0,str(self.NetworkTable_No_Of_Hosts_Array[curItem]))
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
            
            #self.OfficeSizeValues = ['Small','Medium','Large','Corporate','Datacenter']
            self.ComboBoxNetworkAllocationsOfficeSize['state'] = 'disabled'
            if (self.NetworkTable_Office_Size_Array[curItem] == 'Small'):
                self.ComboBoxNetworkAllocationsOfficeSize.current(0)
            if (self.NetworkTable_Office_Size_Array[curItem] == 'Medium'):
                self.ComboBoxNetworkAllocationsOfficeSize.current(1)
            if (self.NetworkTable_Office_Size_Array[curItem] == 'Large'):
                self.ComboBoxNetworkAllocationsOfficeSize.current(2)
            if (self.NetworkTable_Office_Size_Array[curItem] == 'Corporate'):
                self.ComboBoxNetworkAllocationsOfficeSize.current(3)
            if (self.NetworkTable_Office_Size_Array[curItem] == 'Datacenter'):
                self.ComboBoxNetworkAllocationsOfficeSize.current(4)            

            self.ComboBoxNetworkRegionID['state'] = 'disabled'
            i = 0
            while (i < len(self.RegionTableRegionIDArray)):
                if (self.NetworkTable_Network_Region_ID_Array[curItem] == self.RegionTableRegionIDArray[i]):
                    self.ComboBoxNetworkRegionID.current(i)
                    i = len(self.RegionTableRegionIDArray) + 1
                else:
                    i = i + 1
            self.assigned = self.NetworkTable_ssigned_Array[curItem]
            self.PrimaryKey = self.NetworkTable_Network_Slash_Array[curItem]

    def On_NetworkAllocations_Tree_Refresh(self,event):
        #--- Double Click --
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Allocation','REFRESH Button']    
            Logging.Log(Parameter)            
        self.on_NetworkAllocations_Table_Refresh()

    def Calculate_Network(self,pos):
        Network = self.NetworkAllocationsNetworkFrameEntry.get()
        Slash = pos
        IPCalc = IPAddressCalculator(Network,Slash)
        Hosts = IPCalc.Get_Hosts()
        Network_Results = IPCalc.Get_Networks() # Number of Networks
        self.Network_Range_Results = IPCalc.Get_Network_Range() # [-1] == Error
        if (self.Network_Range_Results[0] != -1):
            self.Is_Network_Valid = True
            self.Network_From_Num = int(self.Network_Range_Results[0])*1000000000 + int(self.Network_Range_Results[1])*1000000 + int(self.Network_Range_Results[2])*1000 + int(self.Network_Range_Results[3])
            self.network_From_Str = str(self.Network_Range_Results[0])+"."+str(self.Network_Range_Results[1])+"."+str(self.Network_Range_Results[2])+"."+str(self.Network_Range_Results[3])

            self.Network_To_Num = int(self.Network_Range_Results[4])*1000000000 + int(self.Network_Range_Results[5])*1000000 + int(self.Network_Range_Results[6])*1000 + int(self.Network_Range_Results[7])
            self.network_To_Str = str(self.Network_Range_Results[4])+"."+str(self.Network_Range_Results[5])+"."+str(self.Network_Range_Results[6])+"."+str(self.Network_Range_Results[7])

            self.NetworkAllocationsFromFrameEntry['state'] = 'normal'
            self.NetworkAllocationsFromFrameEntry.delete(0,END)
            self.NetworkAllocationsFromFrameEntry.insert(0,self.network_From_Str)
            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'

            self.NetworkAllocationsToFrameEntry['state'] = 'normal'
            self.NetworkAllocationsToFrameEntry.delete(0,END)
            self.NetworkAllocationsToFrameEntry.insert(0,self.network_To_Str)
            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'

            self.No_of_Hosts = Hosts[0]
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'normal'
            self.NetworkAllocationsNoHostsFrameEntry.delete(0,END)
            self.NetworkAllocationsNoHostsFrameEntry.insert(0,str(self.No_of_Hosts))
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
        else:
            self.Is_Network_Valid = False
            self.NetworkAllocationsFromFrameEntry['state'] = 'normal'
            self.NetworkAllocationsFromFrameEntry.delete(0,END)
            self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'

            self.NetworkAllocationsToFrameEntry['state'] = 'normal'
            self.NetworkAllocationsToFrameEntry.delete(0,END)
            self.NetworkAllocationsToFrameEntry['state'] = 'readonly'

            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'normal'
            self.NetworkAllocationsNoHostsFrameEntry.delete(0,END)
            self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'
            
            mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
            message = 'Not a Valid Network Format')


    def on_slash_combo_changed(self,event):
        pos = self.ComboBoxNetworkAllocationsNetworkSlash.current()
        self.ComboBoxNetworkAllocationsNetworkSubnetMask.current(pos)
        self.Calculate_Network(pos)
                    
        
    def on_subnet_mask_combo_changed(self,event):
        pos = self.ComboBoxNetworkAllocationsNetworkSubnetMask.current()
        self.ComboBoxNetworkAllocationsNetworkSlash.current(pos)
        self.Calculate_Network(pos)
                    
    
    def Get_Region(self):
        self.Is_Get_Region = False
        if self.db.Connect():
            # SQL Querry to the Device Type
            sql = """
                    SELECT * FROM Region
                    ORDER BY Region_Name ASC
                  """
            #print (sql)
            self.RegionTableRegionIDArray = []
            self.RegionTableRegionNameArray = []
            self.ComboBoxNetworkRegionID['values'] = "Empty"
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                num = 0
                self.progress['maximum'] = len(self.db.results)
                self.RegionTableRegionIDArray = []
                self.RegionTableRegionNameArray = []
                while (i < len(self.db.results)):
                    if (self.db.results[i][2].strip() not in self.RegionTableRegionIDArray):
                        num = num + 1
                        self.RegionTableRegionIDArray.append(self.db.results[i][2].strip())
                        self.RegionTableRegionNameArray.append(self.db.results[i][3].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxNetworkRegionID['values'] = self.RegionTableRegionNameArray
                if (len(self.RegionTableRegionNameArray)== 0):
                    self.ComboBoxNetworkRegionID['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.ComboBoxNetworkRegionID['state'] = DISABLED
                    self.ComboBoxNetworkRegionID.current(0)
                    #self.DeviceComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                message = 'No Records found for Regions')

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Netwroks List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.NetworkAllocationsNetworkTableColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.NetworkAllocationsFrame,title='Network Allocation',
                            message = '!!! The File was saved !!!')
                    self.NetworkAllocationsWindow.title("Network Allocation                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Network Allocation','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_NetworkAllocations_Window(self):
        global NetworkAllocationsLabel
        if not self.NetworkAllocationsWindowExist:
            self.NetworkAllocationsWindowExist = True
            # Set up the Window
            self.NetworkAllocationsWindow = Tk()
            self.NetworkAllocationsWindow.geometry('1450x650+350+70')
            self.NetworkAllocationsWindow.title("Network Allocation")
            self.NetworkAllocationsWindow.protocol("WM_DELETE_WINDOW", self.on_NetworkAllocationsWindow_quit)
            self.NetworkAllocationsWindow.call('tk', 'scaling', self.Windows_Scaling)
            NetworkAllocationsLabel = Label(self.NetworkAllocationsWindow,text="Helvetica", font=("Helvetica", 19))
            NetworkAllocationsLabel["text"] = "Network Allocation"
            NetworkAllocationsLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.NetworkAllocationsWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_NetworkAllocationsWindow_quit)
            self.NetworkAllocationsWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.NetworkAllocationsFrame = Frame(self.NetworkAllocationsWindow)
            self.NetworkAllocationsFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.NetworkAllocationsFrame.rowconfigure(0, weight=1)
            self.NetworkAllocationsFrame.columnconfigure(0, weight=1)
            if self.db.Connect():
                    self.slash = []
                    i = 0
                    while (i < 33):
                        self.slash.append("Slash: ["+str(i)+"]")
                        i = i + 1
                        
                    self.Subnetmask = []
                    self.Subnetmask.append("0.0.0.0")   #0
                    self.Subnetmask.append("128.0.0.0") #1
                    self.Subnetmask.append("192.0.0.0") #2
                    self.Subnetmask.append("224.0.0.0") #3
                    self.Subnetmask.append("240.0.0.0") #4
                    self.Subnetmask.append("248.0.0.0") #5
                    self.Subnetmask.append("252.0.0.0") #6
                    self.Subnetmask.append("254.0.0.0") #7
                    self.Subnetmask.append("255.0.0.0") #8
                    
                    self.Subnetmask.append("255.128.0.0") #9
                    self.Subnetmask.append("255.192.0.0") #10
                    self.Subnetmask.append("255.224.0.0") #11
                    self.Subnetmask.append("255.240.0.0") #12
                    self.Subnetmask.append("255.248.0.0") #13
                    self.Subnetmask.append("255.252.0.0") #14
                    self.Subnetmask.append("255.254.0.0") #15
                    self.Subnetmask.append("255.255.0.0") #16
                    
                    self.Subnetmask.append("255.255.128.0") #17
                    self.Subnetmask.append("255.255.192.0") #18
                    self.Subnetmask.append("255.255.224.0") #19
                    self.Subnetmask.append("255.255.240.0") #20
                    self.Subnetmask.append("255.255.248.0") #21
                    self.Subnetmask.append("255.255.252.0") #22
                    self.Subnetmask.append("255.255.254.0") #23
                    self.Subnetmask.append("255.255.255.0") #24
                    
                    self.Subnetmask.append("255.255.255.128") #25
                    self.Subnetmask.append("255.255.255.192") #26
                    self.Subnetmask.append("255.255.255.224") #27
                    self.Subnetmask.append("255.255.255.240") #28
                    self.Subnetmask.append("255.255.255.248") #29
                    self.Subnetmask.append("255.255.255.252") #30
                    self.Subnetmask.append("255.255.255.254") #31
                    self.Subnetmask.append("255.255.255.255") #32


                    # Setup Labels
                    NetworkAllocationsNetworkFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsNetworkFrameLabel["text"] = "Network:"
                    NetworkAllocationsNetworkFrameLabel.place(x=10, y=10)
                    self.NetworkAllocationsNetworkFrameEntry = Entry(self.NetworkAllocationsFrame)
                    self.NetworkAllocationsNetworkFrameEntry['width']=50
                    self.NetworkAllocationsNetworkFrameEntry.place(x=140, y=10)
                    self.NetworkAllocationsNetworkFrameEntry['state'] = 'readonly'

                    NetworkAllocationsNetworkSlashLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsNetworkSlashLabel["text"] = "Slash/Subnet Mask:"
                    NetworkAllocationsNetworkSlashLabel.place(x=10, y=40)
                    self.ComboBoxNetworkAllocationsNetworkSlash = ttk.Combobox(self.NetworkAllocationsFrame, state='disabled', width = 10)
                    self.ComboBoxNetworkAllocationsNetworkSlash.bind("<<ComboboxSelected>>", self.on_slash_combo_changed)
                    self.ComboBoxNetworkAllocationsNetworkSlash.place(x = 140, y = 40)
                    self.ComboBoxNetworkAllocationsNetworkSlash['values'] = self.slash
                    #self.ComboBoxNetworkAllocationsNetworkSlash['state'] = 'readonly'
                    self.ComboBoxNetworkAllocationsNetworkSlash.current(0)

                    self.ComboBoxNetworkAllocationsNetworkSubnetMask = ttk.Combobox(self.NetworkAllocationsFrame, state='disabled', width = 20)
                    self.ComboBoxNetworkAllocationsNetworkSubnetMask.bind("<<ComboboxSelected>>", self.on_subnet_mask_combo_changed)
                    self.ComboBoxNetworkAllocationsNetworkSubnetMask.place(x = 240, y = 40)
                    self.ComboBoxNetworkAllocationsNetworkSubnetMask['values'] = self.Subnetmask
                    #self.ComboBoxNetworkAllocationsNetworkSubnetMask['state'] = 'readonly'
                    self.ComboBoxNetworkAllocationsNetworkSubnetMask.current(0)
                    
                    NetworkAllocationsNetworkDescriptionFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsNetworkDescriptionFrameLabel["text"] = "Description:"
                    NetworkAllocationsNetworkDescriptionFrameLabel.place(x=10, y=70)
                    self.NetworkAllocationsNetworkDescriptionFrameEntry = Entry(self.NetworkAllocationsFrame)
                    self.NetworkAllocationsNetworkDescriptionFrameEntry['width']=50
                    self.NetworkAllocationsNetworkDescriptionFrameEntry.place(x=140, y=70)
                    self.NetworkAllocationsNetworkDescriptionFrameEntry['state'] = 'readonly'                    

                    NetworkAllocationsNetworkRegionFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsNetworkRegionFrameLabel["text"] = "Network Region:"
                    NetworkAllocationsNetworkRegionFrameLabel.place(x=10, y=100)
                    self.ComboBoxNetworkRegionID = ttk.Combobox(self.NetworkAllocationsFrame, state='disabled', width = 47)
                    self.ComboBoxNetworkRegionID.place(x = 140, y = 100)

                    self.ButtonRegionTypeAdd = Button(self.NetworkAllocationsFrame, text = '+ Region', command = self.Call_Button_Region_Add, state=DISABLED)
                    self.ButtonRegionTypeAdd.place(x = 450, y = 97, width=75, height=25)
                    self.ButtonRegionTypeAdd['state'] = ACTIVE

                    self.ButtonRegionTypeRefresh = Button(self.NetworkAllocationsFrame, text = 'Refresh', command = self.Call_Button_Region_Refresh, state=DISABLED)
                    self.ButtonRegionTypeRefresh.place(x = 550, y = 97, width=75, height=25)
                    self.ButtonRegionTypeRefresh['state'] = DISABLED

                    NetworkAllocationsFromFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsFromFrameLabel["text"] = "From:"
                    NetworkAllocationsFromFrameLabel.place(x=10, y=130)
                    self.NetworkAllocationsFromFrameEntry = Entry(self.NetworkAllocationsFrame)
                    self.NetworkAllocationsFromFrameEntry['width']=18
                    self.NetworkAllocationsFromFrameEntry.place(x=140, y=130)
                    self.NetworkAllocationsFromFrameEntry['state'] = 'readonly'

                    NetworkAllocationsToFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsToFrameLabel["text"] = "To:"
                    NetworkAllocationsToFrameLabel.place(x=260, y=130)
                    self.NetworkAllocationsToFrameEntry = Entry(self.NetworkAllocationsFrame)
                    self.NetworkAllocationsToFrameEntry['width']=18
                    self.NetworkAllocationsToFrameEntry.place(x=300, y=130)
                    self.NetworkAllocationsToFrameEntry['state'] = 'readonly'

                    NetworkAllocationsNoHostsFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsNoHostsFrameLabel["text"] = "No. Of Hosts:"
                    NetworkAllocationsNoHostsFrameLabel.place(x=420, y=130)
                    self.NetworkAllocationsNoHostsFrameEntry = Entry(self.NetworkAllocationsFrame)
                    self.NetworkAllocationsNoHostsFrameEntry['width']=18
                    self.NetworkAllocationsNoHostsFrameEntry.place(x=510, y=130)
                    self.NetworkAllocationsNoHostsFrameEntry['state'] = 'readonly'

                    NetworkAllocationsOfficeSizeFrameLabel = Label(self.NetworkAllocationsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsOfficeSizeFrameLabel["text"] = "Office Size:"
                    NetworkAllocationsOfficeSizeFrameLabel.place(x=10, y=160)
                    self.NetworkAllocationsOfficeSizeValues = ['Small','Medium','Large','Corporate','Datacenter']
                    self.ComboBoxNetworkAllocationsOfficeSize = ttk.Combobox(self.NetworkAllocationsFrame, state='disabled', width = 25)
                    self.ComboBoxNetworkAllocationsOfficeSize['values'] = self.NetworkAllocationsOfficeSizeValues
                    self.ComboBoxNetworkAllocationsOfficeSize.current(0)
                    self.ComboBoxNetworkAllocationsOfficeSize.place(x = 140, y = 162)


                    # Setup Buttons
                    self.ButtonNetworkAllocationsAdd = Button(self.NetworkAllocationsFrame, text = 'Add', command = self.Call_Button_NetworkAllocations_Add, state=DISABLED)
                    self.ButtonNetworkAllocationsAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonNetworkAllocationsAdd['state'] = ACTIVE

                    self.ButtonNetworkAllocationsEdit = Button(self.NetworkAllocationsFrame, text = 'Edit', command = self.Call_Button_NetworkAllocations_Edit, state=DISABLED)
                    self.ButtonNetworkAllocationsEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonNetworkAllocationsRemove = Button(self.NetworkAllocationsFrame, text = 'Remove', command = self.Call_Button_NetworkAllocations_Remove, state=DISABLED)
                    self.ButtonNetworkAllocationsRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonNetworkAllocationsOK = Button(self.NetworkAllocationsFrame, text = 'OK / UPDATE', command = self.Call_Button_NetworkAllocations_OK, state=DISABLED)
                    self.ButtonNetworkAllocationsOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonNetworkAllocationsCancel = Button(self.NetworkAllocationsFrame, text = 'Cancel', command = self.Call_Button_NetworkAllocations_Cancel, state=DISABLED)
                    self.ButtonNetworkAllocationsCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.NetworkAllocationsFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)

                    #------------------ TREE VIEW For NetworkAllocations Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.NetworkAllocationsTreeviewDataColumns = ('Network','Slash','Description','Allocated','Provisioned','Country','Region','Facility','Type','Vlan No.','DHCP','Added By')
                    
                    self.NetworkAllocationsTreeview = ttk.Treeview(self.NetworkAllocationsFrame,columns=self.NetworkAllocationsTreeviewDataColumns, height=19) # <--- Make sure the frame is correct !!
                    self.NetworkAllocationsTreeviewysb = Scrollbar(self.NetworkAllocationsFrame,orient=VERTICAL, command=self.NetworkAllocationsTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.NetworkAllocationsTreeviewxsb = Scrollbar(self.NetworkAllocationsFrame,orient=HORIZONTAL, command=self.NetworkAllocationsTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.NetworkAllocationsTreeview['yscroll'] = self.NetworkAllocationsTreeviewysb.set
                    self.NetworkAllocationsTreeview['xscroll'] = self.NetworkAllocationsTreeviewxsb.set

                    # setup headings and column
                    self.NetworkAllocationsTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West
                    self.NetworkAllocationsTreeview.heading('#1', text='Network',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Network', False)) # E for East and W for West
                    self.NetworkAllocationsTreeview.heading('#2', text='Slash',         anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Slash', False)) # E for East and W for West
                    self.NetworkAllocationsTreeview.heading('#3', text='Description',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Description', False)) 
                    self.NetworkAllocationsTreeview.heading('#4', text='Allocated',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Allocated', False)) 
                    self.NetworkAllocationsTreeview.heading('#5', text='Provisioned',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Provisioned', False)) 
                    self.NetworkAllocationsTreeview.heading('#6', text='Country',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Country', False)) 
                    self.NetworkAllocationsTreeview.heading('#7', text='Region',        anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Region', False)) 
                    self.NetworkAllocationsTreeview.heading('#8', text='Facility',      anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Facility', False)) 
                    self.NetworkAllocationsTreeview.heading('#9', text='Type',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Type', False)) 
                    self.NetworkAllocationsTreeview.heading('#10', text='Vlan No.',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Vlan No.', False)) 
                    self.NetworkAllocationsTreeview.heading('#11',text='DHCP',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'DHCP', False)) 
                    self.NetworkAllocationsTreeview.heading('#12', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Added By', False)) 
                    
                    self.NetworkAllocationsTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.NetworkAllocationsTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.NetworkAllocationsTreeview.column('#2', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#3', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#4', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#5', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#6', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#7', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#8', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#9', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#10', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#11', stretch=1, width=10)
                    self.NetworkAllocationsTreeview.column('#12', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.NetworkAllocationsTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.NetworkAllocationsTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.NetworkAllocationsTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.NetworkAllocationsTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.NetworkAllocationsTreeview.bind('<ButtonRelease-1>', self.on_NetworkAllocations_Tree_select_click) # When Select the Tree
                    self.NetworkAllocationsTreeview.bind("<Double-1>", self.On_NetworkAllocations_Tree_Refresh)
                    self.on_NetworkAllocations_Table_Refresh()
                    self.NetworkAllocationsWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the NetworkAllocations Class....:")
    NetworkAllocations = Class_NetworkAllocations("BV",Windows_Scaling)
    NetworkAllocations.Display_NetworkAllocations_Window()


if __name__ == '__main__':
    Main()

