#==========================================================================
# Program: NetworkProvisioning.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-14-2018
# Date Last Modified: Jan-14-2018
# Summary: This is Class to for the NetworkProvisioning
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


try:
    from IPAddressCalculator import *
    Is_IPAddressCalculator_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO IPAddressCalculator Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_IPAddressCalculator = False



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
    print ("*** NO Facility Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Facility_Available = False  

try:
    from IPAddressCalculator import *
    Is_IPAddressCalculator_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO IPAddressCalculator Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_IPAddressCalculator_Available = False

try:
    from NetworkAllocations import *
    Is_NetworkAllocations_Available = True
    NetworkAllocations = Class_NetworkAllocations(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************************* \n")
    print ("*** NO NetworkProvisioning Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************************* \n")
    Is_NetworkAllocations_Available = False

try:
    from NetworkTemplatesBreakOut import *
    Is_NetworkTemplatesBreakout_Available = True
    NetworkTemplatesBreakOut = Class_NetworkTemplatesBreakOut(ODBC_DSN_name,Windows_Scaling)
    #NetworkTemplatesBreakout = Class_NetworkTemplatesBreakout(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************************** \n")
    print ("*** NO NetworkTemplatesBreakOut Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************************** \n")
    Is_NetworkTemplatesBreakout_Available = False

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['NetworkProvisioning','OPEN Window']    
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

class Class_NetworkProvisioning:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.NetworkProvisioningWindowExist = False
        self.NetworkProvisioningCalendarExist = False
        self.Username = os.getlogin()
        self.date = ""
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

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************
        
    def Clean_Screen(self,option,option2):
        # Setup Buttons

        #self.NetworkProvisioningBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False
        
        if (option == 'country'): ## The option are country,region and NetworkProvisioning
            self.ComboBoxRegionID.set("")
            self.ComboBoxRegionID['state'] = DISABLED
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            if (option2 != 'country-combo'):
                self.ComboBoxCoutryID.set("")
                self.ComboBoxRegionID['state'] = DISABLED
                self.ComboBoxFacilityID['state'] = DISABLED
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = DISABLED
                self.ButtonRegionRefresh['state'] = DISABLED
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = DISABLED
                self.ButtonFacilityRefresh['state'] = DISABLED
            if (Is_NetworkAllocations_Available):
                self.ButtonAllocationsAdd['state'] = DISABLED
            if (Is_NetworkTemplatesBreakout_Available):
                self.ButtonTemplatesAdd['state'] = DISABLED

        if (option == 'region'):
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            if (option2 != 'region-combo'):
                self.ComboBoxRegionID.set("")
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = DISABLED
                self.ButtonFacilityRefresh['state'] = DISABLED
            if (Is_NetworkAllocations_Available):
                self.ButtonAllocationsAdd['state'] = DISABLED
            if (Is_NetworkTemplatesBreakout_Available):
                self.ButtonTemplatesAdd['state'] = DISABLED

        if (option2 == 'facility-combo'):
            self.ButtonNetworkProvisioningAdd['state'] = ACTIVE
        else:
            self.ButtonNetworkProvisioningAdd['state'] = DISABLED
            
        self.ButtonNetworkProvisioningRemove['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.NetworkProvisioningTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.NetworkProvisioningTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------

    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.NetworkProvisioningID = self.NetworkProvisioningIDFrameEntry.get()
        self.NetworkProvisioningName = self.NetworkProvisioningNameFrameEntry.get()
        self.NetworkProvisioningDescription = self.NetworkProvisioningDescriptionFrameEntry.get()

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <END> ------------------------*
#****************************************************************************************


#****************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************

    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # Disable the Network and Templares Availables
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['values'] = ["NONE"]
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current(0)
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)

            # SQL Querry to the NetworkProvisioning Table
            sql = """
                    SELECT * FROM Region
                    WHERE Country_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = ACTIVE
                self.ButtonRegionRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.RegionIDArray = []
                self.RegionNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.RegionIDArray.append(self.db.results[i][2].strip())
                    self.RegionNameArray.append(self.db.results[i][3].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxRegionID['values'] = self.RegionNameArray
                if (len(self.RegionNameArray)== 0):
                    self.ComboBoxRegionID['state'] = DISABLED
                    self.ComboBoxFacilityID['state'] = DISABLED
                else:
                    self.ComboBoxRegionID['state'] = 'readonly'
                    self.ComboBoxRegionID.set("")
                    self.ComboBoxFacilityID.set("")
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the NetworkProvisioning Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CountryIDArray.append(self.db.results[i][0].strip())
                    self.CountryNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxCoutryID['values'] = self.CountryNameArray
                if (len(self.CountryNameArray)== 0):
                    self.ComboBoxCoutryID['state'] = DISABLED
                else:
                    self.Clean_Screen('country','all')
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Country Records found')
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        

#**************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <END> ------------------------*
#**************************************************************************************

#***************************************************************************************
#---------------------------- REGION SELECTION SECTION <BEGIN> ------------------------*
#***************************************************************************************
        
    def Display_Region_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Region Window']    
            Logging.Log(Parameter)
        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        self.Clean_Screen('region','region-combo')
        if self.db.Connect():
            # Disable the Network and Templares Availables
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['values'] = ["NONE"]
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current(0)
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)

            # SQL Querry to the NetworkProvisioning Table
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                    ORDER BY Facility_Name ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (Is_NetworkAllocations_Available):
                self.ButtonAllocationsAdd['state'] = ACTIVE
            if (Is_NetworkTemplatesBreakout_Available):
                self.ButtonTemplatesAdd['state'] = ACTIVE
                
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.FacilityIDArray = []
                self.FacilityNameArray = []
                self.FacilitySizeArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.FacilityIDArray.append(self.db.results[i][3].strip())
                    self.FacilityNameArray.append(self.db.results[i][4].strip())
                    self.FacilitySizeArray.append(self.db.results[i][7].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxFacilityID['values'] = self.FacilityNameArray
                if (len(self.FacilityNameArray)== 0):
                    self.ComboBoxFacilityID['state'] = DISABLED
                else:
                    self.ComboBoxFacilityID['state'] = 'readonly'
                    self.ComboBoxFacilityID.set("")
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Region Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the NetworkProvisioning Table
            sql = """
                    SELECT * FROM Region
                    WHERE Country_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.RegionIDArray = []
                self.RegionNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.RegionIDArray.append(self.db.results[i][2].strip())
                    self.RegionNameArray.append(self.db.results[i][3].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxRegionID['values'] = self.RegionNameArray
                if (len(self.RegionNameArray)== 0):
                    self.ComboBoxRegionID['state'] = DISABLED
                    self.ComboBoxFacilityID['state'] = DISABLED
                else:
                    self.ComboBoxRegionID['state'] = 'readonly'
                    self.ComboBoxRegionID.set("")
                    if (Is_Region_Available):
                        self.ButtonRegionAdd['state'] = 'active'
                        self.ButtonRegionRefresh['state'] = 'active'
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Records found')
                self.ComboBoxRegionID['state'] = DISABLED
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonNetworkProvisioningAdd['state']    = DISABLED
                self.ButtonNetworkProvisioningRemove['state'] = DISABLED

                self.NetworkProvisioningIDFrameEntry['state'] = 'normal'
                self.NetworkProvisioningIDFrameEntry.delete(0,END)
                self.NetworkProvisioningIDFrameEntry['state'] = 'readonly'
                
                self.NetworkProvisioningNameFrameEntry['state'] = 'normal'
                self.NetworkProvisioningNameFrameEntry.delete(0,END)
                self.NetworkProvisioningNameFrameEntry['state'] = 'readonly'

        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- REGION SELECTION SECTION <END> ------------------------*
#*************************************************************************************


#***************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <BEGIN> ------------------------*
#***************************************************************************************
        
    def Display_Facility_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Facility Window']    
            Logging.Log(Parameter)
        Facility.Display_Facility_Window()

    def On_Find_Available_Networks(self):
        #time.sleep(2)
        if self.db.Connect():            
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
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """
                        SELECT * FROM Networks
                        WHERE Network_Region_ID = '%s' AND Office_Size = '%s' AND Assigned = '%s'
                        ORDER BY IP_From_Num ASC, Slash ASC""" % (self.RegionIDArray[self.ComboBoxRegionID.current()],
                        self.FacilitySizeArray[self.ComboBoxFacilityID.current()],'NO')

            if (self.db.Execute(sql)):
                self.data2_ready = True
                self.NetworkTable_Available_Network_Slash_Array = []
                self.NetworkTable_Available_Network_Array = []
                self.NetworkTable_Available_Slash_Array = []
                self.NetworkTable_Available_Subnet_Mask_Array = []
                self.NetworkTable_Available_IPA_Array = []
                self.NetworkTable_Available_IPB_Array = []
                self.NetworkTable_Available_IPC_Array = []
                self.NetworkTable_Available_IPD_Array = []
                self.NetworkTable_Available_SMA_Array = []
                self.NetworkTable_Available_SMB_Array = []
                self.NetworkTable_Available_SMC_Array = []
                self.NetworkTable_Available_SMD_Array = []
                self.NetworkTable_Available_No_Of_Hosts_Array = []
                self.NetworkTable_Available_IP_From_Array = []
                self.NetworkTable_Available_IP_To_Array = []
                self.NetworkTable_Available_IP_From_Num_Array = []
                self.NetworkTable_Available_IP_To_Num_Array = []
                self.NetworkTable_Available_Assigned_Array = []
                self.NetworkTable_Available_Network_Region_ID_Array = []
                self.NetworkTable_Available_Office_Size_Array = []
                self.NetworkTable_Available_Status_Array = []                        
                self.NetworkTable_Available_VLAN_L3_Type_Array = []
                self.NetworkTable_Available_VLAN_L3_Name_Array = []
                self.NetworkTable_Available_Description_Array = []
                self.NetworkTable_Available_DHCP_Array = []
                self.NetworkTable_Available_Allocation_Date_Array = []
                self.NetworkTable_Available_Provision_Date_Array = []
                self.NetworkTable_Available_Executed_by_UserID_Array = []
                self.NetworkTable_Available_Country_ID_Array = []
                self.NetworkTable_Available_Region_ID_Array = []
                self.NetworkTable_Available_Facility_ID_Array = []
                self.NetworkTable_Available_Site_ID_Array = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.NetworkTable_Available_Network_Slash_Array.append(self.db.results[i][0].strip())
                    self.NetworkTable_Available_Network_Array.append(self.db.results[i][1].strip())
                    self.NetworkTable_Available_Slash_Array.append(self.db.results[i][2])
                    self.NetworkTable_Available_Subnet_Mask_Array.append(self.db.results[i][3].strip())
                    self.NetworkTable_Available_IPA_Array.append(self.db.results[i][4])
                    self.NetworkTable_Available_IPB_Array.append(self.db.results[i][5])
                    self.NetworkTable_Available_IPC_Array.append(self.db.results[i][6])
                    self.NetworkTable_Available_IPD_Array.append(self.db.results[i][7])
                    self.NetworkTable_Available_SMA_Array.append(self.db.results[i][8])
                    self.NetworkTable_Available_SMB_Array.append(self.db.results[i][9])
                    self.NetworkTable_Available_SMC_Array.append(self.db.results[i][10])
                    self.NetworkTable_Available_SMD_Array.append(self.db.results[i][11])
                    self.NetworkTable_Available_No_Of_Hosts_Array.append(self.db.results[i][12])
                    self.NetworkTable_Available_IP_From_Array.append(self.db.results[i][13].strip())
                    self.NetworkTable_Available_IP_To_Array.append(self.db.results[i][14].strip())
                    self.NetworkTable_Available_IP_From_Num_Array.append(self.db.results[i][15])
                    self.NetworkTable_Available_IP_To_Num_Array.append(self.db.results[i][16])
                    self.NetworkTable_Available_Assigned_Array.append(self.db.results[i][17].strip())
                    self.NetworkTable_Available_Network_Region_ID_Array.append(self.db.results[i][18].strip())
                    self.NetworkTable_Available_Office_Size_Array.append(self.db.results[i][19].strip())
                    self.NetworkTable_Available_Status_Array.append(self.db.results[i][20].strip())
                    if (self.db.results[i][21] == None):
                        self.NetworkTable_Available_VLAN_L3_Type_Array.append("")
                    else:    
                        self.NetworkTable_Available_VLAN_L3_Type_Array.append(self.db.results[i][21].strip())
                    if (self.db.results[i][22] == None):
                        self.NetworkTable_Available_VLAN_L3_Name_Array.append("")
                    else:
                        self.NetworkTable_Available_VLAN_L3_Name_Array.append(self.db.results[i][22].strip())
                    self.NetworkTable_Available_Description_Array.append(self.db.results[i][23].strip())
                    if (self.db.results[i][24] == None):
                        self.NetworkTable_Available_DHCP_Array.append("NO")
                    else:    
                        self.NetworkTable_Available_DHCP_Array.append(self.db.results[i][24].strip())
                    self.NetworkTable_Available_Allocation_Date_Array.append(self.db.results[i][25].strip())
                    if (self.db.results[i][26] == None):
                        self.NetworkTable_Available_Provision_Date_Array.append("")
                    else:
                        self.NetworkTable_Available_Provision_Date_Array.append(self.db.results[i][26].strip())
                    self.NetworkTable_Available_Executed_by_UserID_Array.append(self.db.results[i][27].strip())
                    if (self.db.results[i][28] == None):
                         self.NetworkTable_Available_Country_ID_Array.append("")
                    else:
                        self.NetworkTable_Available_Country_ID_Array.append(self.db.results[i][28].strip())
                    if (self.db.results[i][29] == None):
                        self.NetworkTable_Available_Region_ID_Array.append("")
                    else:
                        self.NetworkTable_Available_Region_ID_Array.append(self.db.results[i][29].strip())
                    if (self.db.results[i][30] == None):
                        self.NetworkTable_Available_Facility_ID_Array.append("")
                    else:
                        self.NetworkTable_Available_Facility_ID_Array.append(self.db.results[i][30].strip())
                    if (self.db.results[i][31] == None):
                        self.NetworkTable_Available_Site_ID_Array.append("")
                    else:
                        self.NetworkTable_Available_Site_ID_Array.append(self.db.results[i][31].strip())
                    i = i + 1
                    
                i = 0
                self.available_results = []
                while (i < len(self.NetworkTable_Available_Network_Slash_Array)):
                    item = [
                            self.NetworkTable_Available_Network_Array[i] + '/' + str(self.NetworkTable_Available_Slash_Array[i]),
                            ]
                    self.available_results.append(item)
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxNetworkAllocationsAvailableNetworkSlash.set("Select Network")
                self.ComboBoxNetworkAllocationsAvailableNetworkSlash['values'] = self.available_results
                self.ComboBoxNetworkAllocationsAvailableNetworkSlash['state'] = 'readonly'
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)

            else:
                self.ComboBoxNetworkAllocationsAvailableNetworkSlash.set("NO Network Found")
                self.ComboBoxNetworkAllocationsAvailableNetworkSlash['state'] = 'disable'
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)

                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                message = 'No Available Networks for this Region and Office Size were found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def On_Find_Networks(self):
        #time.sleep(2)
        if self.db.Connect():            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.NetworkProvisioningTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.NetworkProvisioningTreeview.delete(child)
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
            
            # SQL Querry to the NetworkProvisioning Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """
                        SELECT * FROM Networks
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                        ORDER BY IP_From_Num ASC, Slash ASC""" % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                         self.FacilityIDArray[self.ComboBoxFacilityID.current()])
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
                self.NetworkTable_Assigned_Array = []
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
                    self.NetworkTable_Assigned_Array.append(self.db.results[i][17].strip())
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
                    self.NetworkProvisioningNetworkTableColumns = ('Network-Slash','Network','Slash','Subnet Mask','IPA','IPB','IPC','IPD','SMA','SMB','SMC','SMD',
                                                                  'No. Of Host','Network From','Network To','Network From Num','Network To Num','Asigned','Network Region',
                                                                  'Office Size','Status','Type','Vlan No.','Description','DHCP','Allocated','Provisioned','Added By',
                                                                  'Country','Region','Facility','Site')
                    if (self.NetworkTable_VLAN_L3_Type_Array[i] != 'Summary'):
                        self.NetworkTableNetowrk = "     " + self.NetworkTable_Network_Array[i]
                    else:
                        self.NetworkTableNetowrk = self.NetworkTable_Network_Array[i]
                        tags = "Summary"
                    item = [
                            #self.NetworkTable_Network_Array[i],
                            self.NetworkTableNetowrk,
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
                    self.NetworkProvisioningTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    self.NetworkProvisioningTable = [
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
                            self.NetworkTable_Assigned_Array[i],
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
                    self.results2.append(self.NetworkProvisioningTable)
                    i = i + 1
                    self.progress['value'] = i
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                message = 'No Records found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def on_facility_combo_changed(self,event):
        self.Clean_Screen('facility','facility-combo')
        self.FacilitySizeFrameEntry['state'] = 'normal'
        self.FacilitySizeFrameEntry.delete(0,END)
        self.FacilitySizeFrameEntry.insert(0,self.FacilitySizeArray[self.ComboBoxFacilityID.current()])
        self.FacilitySizeFrameEntry['state'] = 'readonly'

        self.ButtonNetworkProvisioningAdd['state']    = DISABLED
        self.ButtonNetworkProvisioningRemove['state'] = DISABLED
        
        self.On_Find_Networks()
        self.On_Find_Available_Networks()


    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Facility Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('facility','all')
        self.on_facility_combo_changed("event") #Jorge R
        if self.db.Connect():
            # Disable the Network and Templares Availables
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['values'] = ["NONE"]
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current(0)
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
            self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)
            # SQL Querry to the NetworkProvisioning Table
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.FacilityIDArray = []
                self.FacilityNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.FacilityIDArray.append(self.db.results[i][3].strip())
                    self.FacilityNameArray.append(self.db.results[i][4].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxFacilityID['values'] = self.FacilityNameArray
                if (len(self.FacilityNameArray)== 0):
                    self.ComboBoxFacilityID['state'] = DISABLED
                else:
                    self.ComboBoxFacilityID['state'] = 'readonly'
                    self.ComboBoxFacilityID.set("")
                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd['state'] = 'active'
                        self.ButtonFacilityRefresh['state'] = 'active'
                    if (Is_NetworkAllocations_Available):
                        self.ButtonAllocationsAdd['state'] = ACTIVE
                    if (Is_NetworkTemplatesBreakout_Available):
                        self.ButtonTemplatesAdd['state'] = ACTIVE       
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Records found')
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonNetworkProvisioningAdd['state']    = DISABLED
                self.ButtonNetworkProvisioningRemove['state'] = DISABLED

                self.NetworkProvisioningIDFrameEntry['state'] = 'normal'
                self.NetworkProvisioningIDFrameEntry.delete(0,END)
                self.NetworkProvisioningIDFrameEntry['state'] = 'readonly'
                
                self.NetworkProvisioningNameFrameEntry['state'] = 'normal'
                self.NetworkProvisioningNameFrameEntry.delete(0,END)
                self.NetworkProvisioningNameFrameEntry['state'] = 'readonly'

        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <END> ------------------------*
#*************************************************************************************


#*********************************************************************************************************
#---------------------------- Allocation and Templates SELECTION SECTION <BEGIN> ------------------------*
#*********************************************************************************************************

    def on_Display_NetworkAllocations_Window(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Region Window']    
            Logging.Log(Parameter)
        NetworkAllocations.Display_NetworkAllocations_Window()

    def Display_NetworkTemplatesBreakout_Window(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','Region Window']    
            Logging.Log(Parameter)
        NetworkTemplatesBreakOut.Display_NetworkTemplatesBreakOut_Window()
        

#*********************************************************************************************************
#---------------------------- Allocation and Templates SELECTION SECTION <END> ------------------------*
#*********************************************************************************************************
        
    def on_NetworkProvisioningWindow_quit(self):
        if (self.NetworkProvisioningWindowExist):
            self.NetworkProvisioningWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['NetworkProvisioning','CLOSE Window']    
                Logging.Log(Parameter)
            self.NetworkProvisioningWindow.destroy()


    def on_NetworkProvisioning_Table_Refresh(self): 
        self.on_country_combo_changed("event")

       
    def Call_Button_NetworkProvisioning_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','ADD Button']    
            Logging.Log(Parameter)
        self.Selection = 'add'
        #----------- using the IP Calulator to get all subnets for the selected netowork -------------------------------------
        self.Summary_Network = self.NetworkTable_Available_Network_Array[self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current()]
        self.Summary_Slash = self.TemplateBaseNetworkSlashArray[self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current()]
        Brake_into_Slash = self.TemplateSplitInSlashArray[self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current()]
        Number_of_Consecutive_Networks = 2**(Brake_into_Slash - self.Summary_Slash)
        
        #----------> Getting All Subnets  <---------------        
        IPCalc = IPAddressCalculator(self.Summary_Network,Brake_into_Slash)
        results = IPCalc.Get_Network_Range_All(Number_of_Consecutive_Networks)
        i = 0
        its_clean = 0
        while (i < len(results)):
            if (results[i][:len(results[i])] == '-1'):
                i = i + len(results)
                its_clean = 1
            else:
                i = i + 1        
        if (self.db.Connect()) and (its_clean == 0):
            # SQL Querry to the NetworkTemplatesBreakOut Table
            sql = """
                    SELECT * FROM Network_Templates_Break_Out
                    WHERE Template_ID = '%s'
                    ORDER BY Sequence_1 ASC, Sequence_2 ASC
                  """ % (self.TemplateIDArray[self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current()])        
            if (self.db.Execute(sql)):
                self.Template_Break_Out_Table_TemplateIDSEQ1SEQ2Array = []
                self.Template_Break_Out_Table_TemplateIDArray = []
                self.Template_Break_Out_Table_Sequence1Array = []
                self.Template_Break_Out_Table_Sequence2Array = []
                self.Template_Break_Out_Table_SplitinSlashArray = []
                self.Template_Break_Out_Table_VLANL3TypeArray = []
                self.Template_Break_Out_Table_VLAN_L3_NameArray = []
                self.Template_Break_Out_Table_DescriptionArray = []
                self.Template_Break_Out_Table_DHCPArray = []
                self.Template_Break_Out_Table_SplitMoreArray = []
                self.Template_Break_Out_Table_SplitMoreSlashArray = []
                self.Template_Break_Out_Table_ExecutedbyUserIDArray = []
                self.results = []
                self.results2 = []
                self.data_ready = True   
                i = 0
                TemplatesBreakOut = []
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
                    self.Template_Break_Out_Table_TemplateIDSEQ1SEQ2Array.append(self.db.results[i][0].strip())
                    self.Template_Break_Out_Table_TemplateIDArray.append(self.db.results[i][1].strip())
                    self.Template_Break_Out_Table_Sequence1Array.append(self.db.results[i][2])
                    self.Template_Break_Out_Table_Sequence2Array.append(self.db.results[i][3])
                    self.Template_Break_Out_Table_SplitinSlashArray.append(self.db.results[i][4])
                    self.Template_Break_Out_Table_VLANL3TypeArray.append(self.db.results[i][5].strip())
                    if (self.db.results[i][6] == None):
                        self.Template_Break_Out_Table_VLAN_L3_NameArray.append("")
                    else:
                        self.Template_Break_Out_Table_VLAN_L3_NameArray.append(self.db.results[i][6].strip())
                    if (self.db.results[i][7] == None):
                        self.Template_Break_Out_Table_DescriptionArray.append("")
                    else:
                        self.Template_Break_Out_Table_DescriptionArray.append(self.db.results[i][7].strip())
                    self.Template_Break_Out_Table_DHCPArray.append(self.db.results[i][8].strip())
                    self.Template_Break_Out_Table_SplitMoreArray.append(self.db.results[i][9].strip())
                    self.Template_Break_Out_Table_SplitMoreSlashArray.append(self.db.results[i][10])
                    self.Template_Break_Out_Table_ExecutedbyUserIDArray.append(self.db.results[i][11].strip())
                    i = i + 1
                    # ('Sequence','Slash','Type','Description','DHCP','Split More','In Slash','Added By')
                i = 0
                while (i < len(self.Template_Break_Out_Table_TemplateIDSEQ1SEQ2Array)):
                    num = i + 1
                        
                    tags = self.Template_Break_Out_Table_VLANL3TypeArray[i]
                    if (self.Template_Break_Out_Table_Sequence2Array[i] > 0):
                        tags = "SummarySplit"
                    item = [
                            (str(self.Template_Break_Out_Table_Sequence1Array[i])+"."+str(self.Template_Break_Out_Table_Sequence2Array[i])),
                            self.Template_Break_Out_Table_SplitinSlashArray[i],
                            self.Template_Break_Out_Table_VLANL3TypeArray[i],
                            self.Template_Break_Out_Table_VLAN_L3_NameArray[i],
                            self.Template_Break_Out_Table_DescriptionArray[i],
                            self.Template_Break_Out_Table_DHCPArray[i],
                            self.Template_Break_Out_Table_SplitMoreArray[i],
                            self.Template_Break_Out_Table_SplitMoreSlashArray[i],
                            self.Template_Break_Out_Table_ExecutedbyUserIDArray[i]
                            ]
                    self.results.append(item)
                    item2 = [
                            self.Template_Break_Out_Table_TemplateIDSEQ1SEQ2Array[i],
                            self.Template_Break_Out_Table_TemplateIDArray[i],
                            self.Template_Break_Out_Table_Sequence1Array[i],
                            self.Template_Break_Out_Table_Sequence2Array[i],
                            self.Template_Break_Out_Table_SplitinSlashArray[i],
                            self.Template_Break_Out_Table_VLANL3TypeArray[i],
                            self.Template_Break_Out_Table_VLAN_L3_NameArray[i],
                            self.Template_Break_Out_Table_DescriptionArray[i],
                            self.Template_Break_Out_Table_DHCPArray[i],
                            self.Template_Break_Out_Table_SplitMoreArray[i],
                            self.Template_Break_Out_Table_SplitMoreSlashArray[i],
                            self.Template_Break_Out_Table_ExecutedbyUserIDArray[i]
                            ]
                    self.results2.append(item2)
                    TemplatesBreakOut.append(item2)
                    i = i + 1
                self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
                self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
                self.NetworkProvisionOfficeSize = self.FacilitySizeArray[self.ComboBoxFacilityID.current()]
                i = 0
                while (i < len(TemplatesBreakOut)):
                    if (TemplatesBreakOut[i][5] == 'Summary'):
                        #----------------------- Now we need to obtain the Sumary break out subnets -------- <BEGIN>-----------------
                        Network = str(results[TemplatesBreakOut[i][2]-1][0]) +"." + \
                                  str(results[TemplatesBreakOut[i][2]-1][1]) +"." + \
                                  str(results[TemplatesBreakOut[i][2]-1][2]) +"." + \
                                  str(results[TemplatesBreakOut[i][2]-1][3])
                        Slash = TemplatesBreakOut[i][4]
                        Brake_into_Slash = TemplatesBreakOut[i][10]
                        Number_of_Consecutive_Networks = 2**(Brake_into_Slash - Slash)
                        
                        IPCalc = IPAddressCalculator(Network,Brake_into_Slash)
                        results_summary = IPCalc.Get_Network_Range_All(Number_of_Consecutive_Networks)
                        j = 0
                        while (j < len(results_summary)):
                            i = i + 1
                            #-------------Adding to the Database -------------------------------
                            self.Network_From_Num = int(results_summary[j][0])*1000000000 + int(results_summary[j][1])*1000000 + int(results_summary[j][2])*1000 + int(results_summary[j][3])
                            self.network_From_Str = str(results_summary[j][0])+ "."+ str(results_summary[j][1])+"."+str(results_summary[j][2])+ "."+ str(results_summary[j][3])
                            self.No_of_Hosts = 2 ** (32 - TemplatesBreakOut[i][4])
                            self.Network_To_Num = int(results_summary[j][4])*1000000000 + int(results_summary[j][5])*1000000 + int(results_summary[j][6])*1000 + int(results_summary[j][7])
                            self.network_To_Str = str(results_summary[j][4])+ "."+ str(results_summary[j][5])+"."+str(results_summary[j][6])+ "."+ str(results_summary[j][7])
                            self.NetworkProvisionNetwork = self.network_From_Str
                            self.NetworkProvisionNetworkDescription = TemplatesBreakOut[i][7]
                            self.NetworkProvisionNetworkSlash = TemplatesBreakOut[i][4]
                            self.NetworkProvisionNetworkSubnetMask = self.Subnetmask[TemplatesBreakOut[i][4]]
                            self.NetowrkProvisionSubnetMaskArray = self.NetworkProvisionNetworkSubnetMask.split('.')
                            self.NetworkProvisionOfficeSize = self.FacilitySizeArray[self.ComboBoxFacilityID.current()]
                            self.NetworkProvisionNetworkRegionID = self.RegionID
                            PrimaryKey = self.network_From_Str + "-" + str(self.NetworkProvisionNetworkSlash)

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
                                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                                message = '*** The Network you entered [ '+ self.network_From_Str + "/" + str(self.NetworkProvisionNetworkSlash) +' ] already exist ***')
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
                                           '%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.network_From_Str,
                                                                                             int(self.NetworkProvisionNetworkSlash),
                                                                                             self.NetworkProvisionNetworkSubnetMask,
                                                                                             int(results_summary[j][0]),int(results_summary[j][1]), int(results_summary[j][2]),
                                int(results_summary[j][3]),int(self.NetowrkProvisionSubnetMaskArray[0]),int(self.NetowrkProvisionSubnetMaskArray[1]),int(self.NetowrkProvisionSubnetMaskArray[2]),
                                int(self.NetowrkProvisionSubnetMaskArray[3]),int(self.No_of_Hosts),self.network_From_Str,self.network_To_Str,self.Network_From_Num,
                                self.Network_To_Num,'YES',self.NetworkProvisionNetworkRegionID,self.NetworkProvisionOfficeSize,'Provision',TemplatesBreakOut[i][5],
                                TemplatesBreakOut[i][6],TemplatesBreakOut[i][7],TemplatesBreakOut[i][8],self.date,self.date,self.Username,self.CountryID,self.RegionID,self.FacilityID,'MDF')
                                
                                if not(self.db.Add_Move_Change_Data(sql)):
                                    mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                        message = '*** The Network you entered was NOT Added ***')
                            j = j + 1                            
                        #----------------------- Now we need to obtain the Sumary break out subnets -------- <END>-----------------
                    else:
                        k = TemplatesBreakOut[i][2] - 1
                        #-------------------Adding to the Database -------------------------
                        self.Network_From_Num = int(results[k][0])*1000000000 + int(results[k][1])*1000000 + int(results[k][2])*1000 + int(results[k][3])
                        self.network_From_Str = str(results[k][0])+ "."+ str(results[k][1])+"."+str(results[k][2])+ "."+ str(results[k][3])
                        self.No_of_Hosts = 2 ** (32 - TemplatesBreakOut[i][4])
                        self.Network_To_Num = int(results[k][4])*1000000000 + int(results[k][5])*1000000 + int(results[k][6])*1000 + int(results[k][7])
                        self.network_To_Str = str(results[k][4])+ "."+ str(results[k][5])+"."+str(results[k][6])+ "."+ str(results[k][7])
                        self.NetworkProvisionNetwork = self.network_From_Str
                        self.NetworkProvisionNetworkDescription = TemplatesBreakOut[i][7]
                        self.NetworkProvisionNetworkSlash = TemplatesBreakOut[i][4]
                        self.NetworkProvisionNetworkSubnetMask = self.Subnetmask[TemplatesBreakOut[i][4]]
                        self.NetowrkProvisionSubnetMaskArray = self.NetworkProvisionNetworkSubnetMask.split('.')
                        self.NetworkProvisionOfficeSize = self.FacilitySizeArray[self.ComboBoxFacilityID.current()]
                        self.NetworkProvisionNetworkRegionID = self.RegionID
                        PrimaryKey = self.network_From_Str + "-" + str(self.NetworkProvisionNetworkSlash)

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
                            mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provisioning',
                            message = '*** The Network you entered [ '+ self.network_From_Str + "/" + str(self.NetworkProvisionNetworkSlash) +' ] already exist ***')
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
                                           '%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.network_From_Str,
                                                                                             int(self.NetworkProvisionNetworkSlash),
                                                                                             self.NetworkProvisionNetworkSubnetMask,
                                                                                             int(results[k][0]),int(results[k][1]), int(results[k][2]),
                                int(results[k][3]),int(self.NetowrkProvisionSubnetMaskArray[0]),int(self.NetowrkProvisionSubnetMaskArray[1]),int(self.NetowrkProvisionSubnetMaskArray[2]),
                                int(self.NetowrkProvisionSubnetMaskArray[3]),int(self.No_of_Hosts),self.network_From_Str,self.network_To_Str,self.Network_From_Num,
                                self.Network_To_Num,'YES',self.NetworkProvisionNetworkRegionID,self.NetworkProvisionOfficeSize,'Provision',TemplatesBreakOut[i][5],
                                TemplatesBreakOut[i][6],TemplatesBreakOut[i][7],TemplatesBreakOut[i][8],self.date,self.date,self.Username,self.CountryID,self.RegionID,self.FacilityID,'MDF')
                                
                            if not (self.db.Add_Move_Change_Data(sql)):
                                mbox.showerror(master=self.NetworkAllocationsFrame,title='Network Allocation',
                                    message = '*** The Network you entered was NOT Added ***')
                    i = i + 1
                #------------------ Updating the status of the Summary <BEIGN> ------------------------
                PrimaryKey = self.Summary_Network + "-" + str(self.Summary_Slash)
                sql = """
                        SELECT * FROM Networks
                        WHERE Network_Slash = '%s'
                      """ % (PrimaryKey)
                    
                if (self.db.Execute(sql)):
                    sql = "UPDATE Networks SET Assigned = '%s', Country_ID = '%s', Region_ID = '%s', Facility_ID = '%s', Site_ID = '%s', Provision_Date = '%s', \
                                    Executed_by_UserID = '%s' \
                           WHERE Network_Slash = '%s'" % ('YES',self.CountryID,self.RegionID,self.FacilityID,'MDF',
                                self.date,self.Username,PrimaryKey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        mbox.showwarning(master=self.NetworkProvisioningFrame,title='Network Provision',
                            message = '*** The Summary Network you entered was Updated ***')
                    else:
                        mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provision',
                            message = '*** The Summary Network you entered was NOT Upadted ***')              
                else:
                    mbox.showerror(master=self.NetworkProvisioningFrame,title='Network provison',
                    message = '*** The Summary Network you try to Use Does not exist Anymore ***')
                #------------------ Updating tge status of the Summary <END> --------------------------  
                self.on_facility_combo_changed("event")
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provision',
                message = 'The Network Template was NOT found')
        else:
            if (its_clean == 1):
                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provision',
                                message = '*** ERROR *** - Finding the Network was NOT Succesful')
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Provision',
                                message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
            

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
            #print ("Select")
            curItem = self.NetworkProvisioningTreeview.focus()  
            dic = self.NetworkProvisioningTreeview.item(curItem)
            #print (dic)
            #print (dic.get('text'))
            curItem = int(dic.get('text')) - 1
            #print ("*" * 50)
            #print (self.NetworkTable_Network_Slash_Array[curItem])
            #print ("*" * 50)
            PrimaryKey = self.NetworkTable_Network_Slash_Array[curItem]
            self.Selection = 'remove'
            if (mbox.askyesnocancel(master=self.NetworkProvisioningFrame,title='Network Provision',message = 'Are you Sure you want to Remove it?')):
                self.ButtonNetworkProvisioningRemove['state'] = DISABLED
                #print (PrimaryKey)
                #PrimaryKey = self.PrimaryKey
                sql = "DELETE FROM Networks WHERE Network_Slash = '%s'" % (PrimaryKey)
                if (self.db.Add_Move_Change_Data(sql)):
                    mbox.showwarning(master=self.NetworkProvisioningFrame,title='Network Allocation',
                            message = '*** The Network you entered was Removed ***')
                else:
                    mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Allocation',
                            message = '*** The Network you entered was NOT Removed ***')
                self.on_facility_combo_changed("event")
            self.progress['value'] = 100

        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Allocation',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.GetPasswordWindow.destroy()
            self.Remove_Networks_From_DB() #<--------------- This is the Delete routine
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
            
    def Call_Button_NetworkProvisioning_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['NetworkProvisioning','REMOVE Button']    
            Logging.Log(Parameter)
        self.Get_Usernanme_and_Password()

    #-------------- Using a Password Question to make sure it was the intent to be deleted <End> ---------------


    def on_NetworkProvisioning_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.NetworkProvisioningTreeview.focus()  
        dic = self.NetworkProvisioningTreeview.item(curItem)
        #print (dic)
        #print (dic.get('text'))
        values = dic.get('values')
        if (len(values) > 0):
            #print (values)
            #print (values[0])
            #print (values[1])
            #print (curItem)
            self.ComboBoxCoutryID['state'] = 'readonly'
            self.ComboBoxRegionID['state'] = 'readonly'
            self.ComboBoxFacilityID['state'] = 'readonly'
            if (Is_Country_Available):
                self.ButtonCountryAdd['state'] = ACTIVE
                self.ButtonCountryRefresh['state'] = ACTIVE
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = ACTIVE
                self.ButtonRegionRefresh['state'] = ACTIVE            
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (Is_NetworkAllocations_Available):
                self.ButtonAllocationsAdd['state'] = ACTIVE
            if (Is_NetworkTemplatesBreakout_Available):
                self.ButtonTemplatesAdd['state'] = ACTIVE

            curItem = int(dic.get('text')) - 1
            #self.ButtonNetworkProvisioningAdd['state']    = ACTIVE
            self.ButtonNetworkProvisioningRemove['state'] = ACTIVE

    def On_NetworkProvisioning_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_NetworkProvisioning_Table_Refresh()

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
                ExcellFile.Add_DataToWorksheet(self.NetworkProvisioningNetworkTableColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.NetworkProvisioningFrame,title='Network Provision',
                            message = '!!! The File was saved !!!')
                    self.NetworkProvisioningWindow.title("Network Provision                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Network Allocation','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.NetworkProvisioningFrame,title='Network Allocation',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def on_network_combo_changed(self,event):
        self.ButtonNetworkProvisioningAdd['state']    = DISABLED
        self.ButtonNetworkProvisioningRemove['state'] = DISABLED

        if self.db.Connect():
            # SQL Querry to the NetworkProvisioning Table
            sql = """
                    SELECT * FROM Network_Templates
                    WHERE Base_Network_Slash = %d
                    ORDER BY Template_Name ASC
                  """ % (self.NetworkTable_Available_Slash_Array[self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current()])
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.TemplateIDArray = []
                self.TemplateNameArray = []
                self.TemplateBaseNetworkSlashArray = []
                self.TemplateSplitInSlashArray = []
                while (i < len(self.db.results)):
                    self.TemplateIDArray.append(self.db.results[i][0].strip())
                    self.TemplateNameArray.append(self.db.results[i][1].strip())
                    self.TemplateBaseNetworkSlashArray.append(self.db.results[i][2])
                    self.TemplateSplitInSlashArray.append(self.db.results[i][3])
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.set("Select Template")
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = self.TemplateNameArray
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'readonly'
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                message = 'No Records found')
                self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['state'] = 'disabled'
        else:
            mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
            

    def on_network_template_combo_changed(self,event):
        self.ButtonNetworkProvisioningAdd['state']    = ACTIVE
        self.ButtonNetworkProvisioningRemove['state'] = DISABLED

    def on_slash_combo_changed(self,event):
        pos = self.ComboBoxNetworkAllocationsNetworkSlash.current()
        self.ComboBoxNetworkAllocationsNetworkSubnetMask.current(pos)
        self.Calculate_Network(pos)
                    
        
    def on_subnet_mask_combo_changed(self,event):
        pos = self.ComboBoxNetworkAllocationsNetworkSubnetMask.current()
        self.ComboBoxNetworkAllocationsNetworkSlash.current(pos)
        self.Calculate_Network(pos)

    def Display_NetworkProvisioning_Window(self): 
        if not self.NetworkProvisioningWindowExist:
            # Set up the Window
            self.NetworkProvisioningWindowExist = True
            self.NetworkProvisioningWindow = Tk()
            self.NetworkProvisioningWindow.geometry('1300x670+350+70')
            self.NetworkProvisioningWindow.title("Network Provisioning")
            self.NetworkProvisioningWindow.protocol("WM_DELETE_WINDOW", self.on_NetworkProvisioningWindow_quit)
            self.NetworkProvisioningWindow.call('tk', 'scaling', self.Windows_Scaling)
            NetworkProvisioningLabel = Label(self.NetworkProvisioningWindow,text="Helvetica", font=("Helvetica", 19))
            NetworkProvisioningLabel["text"] = "Network Provisioning"
            NetworkProvisioningLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.NetworkProvisioningWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_NetworkProvisioningWindow_quit)
            self.NetworkProvisioningWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.NetworkProvisioningFrame = Frame(self.NetworkProvisioningWindow)
            self.NetworkProvisioningFrame.pack(side=TOP, fill=BOTH, expand=Y)

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

                    # set frame resizing priorities
                    self.NetworkProvisioningFrame.rowconfigure(0, weight=1)
                    self.NetworkProvisioningFrame.columnconfigure(0, weight=1)

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.NetworkProvisioningFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 8, width=75, height=24)

                        self.ButtonCountryRefresh = Button(self.NetworkProvisioningFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 8, width=75, height=24)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.NetworkProvisioningFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 38, width=75, height=24)

                        self.ButtonRegionRefresh = Button(self.NetworkProvisioningFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 38, width=75, height=24)

                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd = Button(self.NetworkProvisioningFrame, text = '+ Facility', command = self.Display_Facility_Window, state=DISABLED)
                        self.ButtonFacilityAdd.place(x = 450, y = 68, width=75, height=24)

                        self.ButtonFacilityRefresh = Button(self.NetworkProvisioningFrame, text = 'Refresh', command = self.on_Facility_Table_Refresh, state=DISABLED)
                        self.ButtonFacilityRefresh.place(x = 550, y = 68, width=75, height=24)

                    if (Is_NetworkAllocations_Available ):
                        self.ButtonAllocationsAdd = Button(self.NetworkProvisioningFrame, text = '+ Networks', command = self.on_Display_NetworkAllocations_Window, state=DISABLED)
                        self.ButtonAllocationsAdd.place(x = 450, y = 128, width=75, height=24)

                    if (Is_NetworkTemplatesBreakout_Available ):
                        self.ButtonTemplatesAdd = Button(self.NetworkProvisioningFrame, text = '+ Templates', command = self.Display_NetworkTemplatesBreakout_Window, state=DISABLED)
                        self.ButtonTemplatesAdd.place(x = 450, y = 158, width=75, height=24)

                    self.ButtonNetworkProvisioningAdd = Button(self.NetworkProvisioningFrame, text = 'Add', command = self.Call_Button_NetworkProvisioning_Add, state=DISABLED)
                    self.ButtonNetworkProvisioningAdd.place(x = 450, y = 98, width=75, height=25)

                    self.ButtonNetworkProvisioningRemove = Button(self.NetworkProvisioningFrame, text = 'Remove', command = self.Call_Button_NetworkProvisioning_Remove, state=DISABLED)
                    self.ButtonNetworkProvisioningRemove.place(x = 550, y = 98, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.NetworkProvisioningFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=650, y=158)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=40)

                    FacilityIDFrameLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityIDFrameLabel["text"] = "Facility Name:"
                    FacilityIDFrameLabel.place(x=10, y=70)

                    FacilitySizeFrameLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilitySizeFrameLabel["text"] = "Facility Size:"
                    FacilitySizeFrameLabel.place(x=10, y=100)
                    self.FacilitySizeFrameEntry = Entry(self.NetworkProvisioningFrame)
                    self.FacilitySizeFrameEntry['width']=50
                    self.FacilitySizeFrameEntry.place(x=110, y=100)
                    self.FacilitySizeFrameEntry['state'] = DISABLED

                    # Setup Labels and Entry

                    NetworkAllocationsAvailableNetworkSlashLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsAvailableNetworkSlashLabel["text"] = "Available Networks:"
                    NetworkAllocationsAvailableNetworkSlashLabel.place(x=10, y=130)
                    self.ComboBoxNetworkAllocationsAvailableNetworkSlash = ttk.Combobox(self.NetworkProvisioningFrame, state='disabled', width = 42)
                    self.ComboBoxNetworkAllocationsAvailableNetworkSlash.bind("<<ComboboxSelected>>", self.on_network_combo_changed)
                    self.ComboBoxNetworkAllocationsAvailableNetworkSlash.place(x = 140, y = 130)
                    self.ComboBoxNetworkAllocationsAvailableNetworkSlash['values'] = ["NONE"] 
                    self.ComboBoxNetworkAllocationsAvailableNetworkSlash.current(0)


                    NetworkAllocationsAvailableNetworkTemplatesLabel = Label(self.NetworkProvisioningFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkAllocationsAvailableNetworkTemplatesLabel["text"] = "Available Templates:"
                    NetworkAllocationsAvailableNetworkTemplatesLabel.place(x=10, y=160)
                    self.ComboBoxNetworkAllocationsAvailableNetworkTemplates = ttk.Combobox(self.NetworkProvisioningFrame, state='disabled', width = 42)
                    self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.bind("<<ComboboxSelected>>", self.on_network_template_combo_changed)
                    self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.place(x = 140, y = 160)
                    self.ComboBoxNetworkAllocationsAvailableNetworkTemplates['values'] = ["NONE"] 
                    self.ComboBoxNetworkAllocationsAvailableNetworkTemplates.current(0)

                    
                    #------------------ TREE VIEW For NetworkProvisioning Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.NetworkProvisioningTreeviewDataColumns = ('Network','Slash','Description','Allocated','Provisioned','Country','Region','Facility','Type','Vlan No.','DHCP','Added By')
                    
                    self.NetworkProvisioningTreeview = ttk.Treeview(self.NetworkProvisioningFrame,columns=self.NetworkProvisioningTreeviewDataColumns, height=20) # <--- Make sure the frame is correct !!
                    self.NetworkProvisioningTreeviewysb = Scrollbar(self.NetworkProvisioningFrame,orient=VERTICAL, command=self.NetworkProvisioningTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.NetworkProvisioningTreeviewxsb = Scrollbar(self.NetworkProvisioningFrame,orient=HORIZONTAL, command=self.NetworkProvisioningTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.NetworkProvisioningTreeview['yscroll'] = self.NetworkProvisioningTreeviewysb.set
                    self.NetworkProvisioningTreeview['xscroll'] = self.NetworkProvisioningTreeviewxsb.set

                    # setup headings and column
                    self.NetworkProvisioningTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West
                    self.NetworkProvisioningTreeview.heading('#1', text='Network',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Network', False)) # E for East and W for West
                    self.NetworkProvisioningTreeview.heading('#2', text='Slash',         anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Slash', False)) # E for East and W for West
                    self.NetworkProvisioningTreeview.heading('#3', text='Description',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Description', False)) 
                    self.NetworkProvisioningTreeview.heading('#4', text='Allocated',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Allocated', False)) 
                    self.NetworkProvisioningTreeview.heading('#5', text='Provisioned',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Provisioned', False)) 
                    self.NetworkProvisioningTreeview.heading('#6', text='Country',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Country', False)) 
                    self.NetworkProvisioningTreeview.heading('#7', text='Region',        anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Region', False)) 
                    self.NetworkProvisioningTreeview.heading('#8', text='Facility',      anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Facility', False)) 
                    self.NetworkProvisioningTreeview.heading('#9', text='Type',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Type', False)) 
                    self.NetworkProvisioningTreeview.heading('#10', text='Vlan No.',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Vlan No.', False)) 
                    self.NetworkProvisioningTreeview.heading('#11',text='DHCP',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'DHCP', False)) 
                    self.NetworkProvisioningTreeview.heading('#12', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkAllocationsTreeview, 'Added By', False)) 
                    
                    self.NetworkProvisioningTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.NetworkProvisioningTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.NetworkProvisioningTreeview.column('#2', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#3', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#4', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#5', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#6', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#7', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#8', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#9', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#10', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#11', stretch=1, width=10)
                    self.NetworkProvisioningTreeview.column('#12', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.NetworkProvisioningTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.NetworkProvisioningTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.NetworkProvisioningTreeviewxsb.grid(row=2, column=0, sticky=EW)

                    # create fonts and tags
                    # Use later to mark Business Units per color.
                    self.NetworkProvisioningTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                    self.NetworkProvisioningTreeview.tag_configure('Summary', font=('Helvetica', 12), background='steelblue1')
                   
                    # Bind the double Click
                    self.NetworkProvisioningTreeview.bind('<ButtonRelease-1>', self.on_NetworkProvisioning_Tree_select_click) # When Select the Tree
                    #self.NetworkProvisioningTreeview.bind("<Double-1>", self.On_NetworkProvisioning_Tree_Refresh)         

                    #------------------ TREE VIEW For NetworkProvisioning Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.NetworkProvisioningFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.NetworkProvisioningFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 40)

                    self.ComboBoxFacilityID = ttk.Combobox(self.NetworkProvisioningFrame, state='disabled', width = 50)
                    self.ComboBoxFacilityID.bind("<<ComboboxSelected>>", self.on_facility_combo_changed)
                    self.ComboBoxFacilityID.place(x = 110, y = 70)

                    self.on_Country_Table_Refresh()
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.NetworkProvisioningWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.NetworkProvisioningFrame,title='NetworkProvisioning',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the NetworkProvisioning Class....:")
    NetworkProvisioning = Class_NetworkProvisioning("BV",Windows_Scaling)
    NetworkProvisioning.Display_NetworkProvisioning_Window()


if __name__ == '__main__':
    Main()

