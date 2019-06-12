#==========================================================================
# Program: Device.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-21-2018
# Date Last Modified: Feb-20-2018
# Summary: This is Class to for the Device
#==========================================================================

#***************************************************************
# ==================== Libraries Required <BEGIN> =============*
#***************************************************************

import os
from threading import Thread
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
    print ("*** NO Facility Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Facility_Available = False  

try:
    from Sites import *
    Is_Sites_Available = True
    Sites = Class_Sites(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Sites Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Sites_Available = False  

try:
    from Circuits import *
    Is_Circuit_Available = True
    Location = []
    Circuit = Class_Circuits(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Circuits Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Circuit_Available = False  

try:
    from ICMP import *
    Is_ICMP_Available = True
    Location = []
    ICMP = Class_ICMP(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO ICMP Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_ICMP_Available = False  

try:
    from LocalPointOfContacts import *
    Is_LocalPointOfContacts_Available = True
    Location = []
    LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Local Point Of Contacts Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_LocalPointOfContacts_Available = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Device','OPEN Window']    
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

class Class_Device:

    def __init__(self,DSN_Name,Windows_Scaling,Location):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.DeviceWindowExist = False
        self.DeviceCalendarInstalledDateExist = False
        self.DeviceCalendarActivatedDateExist = False
        self.DeviceCalendarDisconnectedDateExist = False
        self.DeviceCalendarExpirationDateExist = False
        self.Username = os.getlogin()
        self.date = ""
        self.Windows_Scaling = Windows_Scaling
        self.Selection = 'none'
        self.GetPasswordWindowsExists = False
        self.Go_To_Location = False
        if (len(Location) > 0):
            self.Init_Country = Location[0]
            self.Init_Region = Location[1]
            self.Init_Facility = Location[2]
            self.Init_Site = Location[3]
            self.Go_To_Location = True

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
        

    def IPFormatCheck(self,ip_str):
        if len(ip_str.split()) == 1:
               ipList = ip_str.split('.')
               if len(ipList) == 4:
                   for i, item in enumerate(ipList):
                       try:
                           ipList[i] = int(item)
                       except:
                           return False
                       if not isinstance(ipList[i], int):
                           return False
                   if max(ipList) < 256:
                       return True
                   else:
                       return False
               else:
                   return False
        else:
               return False

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************
        
    def Clean_Screen(self,option,option2):
        # Setup Buttons

        #self.DeviceBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False
        if (option == 'country'): ## The option are country,region and Device
            self.ComboBoxRegionID.set("")
            self.ComboBoxRegionID['state'] = DISABLED
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            self.ComboBoxSitesID.set("")
            self.ComboBoxSitesID['state'] = DISABLED
            self.ButtonDeviceOK['state'] = DISABLED # Due to Edit
            if (option2 != 'country-combo'):
                self.ComboBoxCoutryID.set("")
                self.ComboBoxRegionID['state'] = DISABLED
                self.ComboBoxFacilityID['state'] = DISABLED
                self.ComboBoxSitesID['state'] = DISABLED
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = DISABLED
                self.ButtonRegionRefresh['state'] = DISABLED
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = DISABLED
                self.ButtonFacilityRefresh['state'] = DISABLED
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = DISABLED
                self.ButtonSitesRefresh['state'] = DISABLED

        if (option == 'region'):
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            self.ComboBoxSitesID.set("")
            self.ComboBoxSitesID['state'] = DISABLED
            self.ButtonDeviceOK['state'] = DISABLED # Due to Edit
            if (option2 != 'region-combo'):
                self.ComboBoxRegionID.set("")
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = DISABLED
                self.ButtonFacilityRefresh['state'] = DISABLED
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = DISABLED
                self.ButtonSitesRefresh['state'] = DISABLED

        if (option == 'facility'):
            self.ComboBoxSitesID.set("")
            self.ComboBoxSitesID['state'] = DISABLED
            self.ButtonDeviceOK['state'] = DISABLED # Due to Edit
            if (option2 != 'facility-combo'):
                self.ComboBoxFacilityID.set("")
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = DISABLED
                self.ButtonSitesRefresh['state'] = DISABLED

        if (option2 == 'sites-combo'):
            if (self.Selection != 'edit'):
                self.ButtonDeviceAdd['state'] = ACTIVE
            if (self.Selection == 'edit'):
                self.ButtonDeviceOK['state'] = ACTIVE # Due to Edit
        else:
            self.ButtonDeviceAdd['state'] = DISABLED

        if (self.Selection != 'edit'):
            self.ButtonDeviceEdit['state'] = DISABLED
            self.ButtonDeviceRemove['state'] = DISABLED
            self.ButtonDeviceOK['state'] = DISABLED
            self.ButtonDeviceCancel['state'] = DISABLED
            self.ButtonDeviceCircuits['state'] = DISABLED
            self.ButtonDevicePing64['state'] = DISABLED
            self.ButtonDevicePing1500['state'] = DISABLED
            #self.ButtonDeviceContacts['state'] = DISABLED
            self.ButtonDeviceICMP['state'] = DISABLED
            self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED

            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.DeviceIDFrameEntry['state'] = 'normal'
            self.DeviceIDFrameEntry.delete(0,END)
            self.DeviceIDFrameEntry['state'] = 'readonly'

            self.DeviceDescriptionFrameEntry['state'] = 'normal'
            self.DeviceDescriptionFrameEntry.delete(0,END)
            self.DeviceDescriptionFrameEntry['state'] = 'readonly'

            # Setup Labels and Entry
            self.DeviceIDFrameEntry['state'] = 'normal'
            self.DeviceIDFrameEntry.delete(0,END)
            self.DeviceIDFrameEntry['state'] = 'readonly'
                    
            self.DeviceDescriptionFrameEntry['state'] = 'normal'
            self.DeviceDescriptionFrameEntry.delete(0,END)
            self.DeviceDescriptionFrameEntry['state'] = 'readonly'

            self.DeviceComboBoxTypeID.current(0)
            self.DeviceComboBoxTypeID['state'] = DISABLED

            self.DeviceComboBoxModelID.current(0)
            self.DeviceComboBoxModelID['state'] = DISABLED

            self.DeviceComboBoxStatus.current(0)
            self.DeviceComboBoxStatus['state'] = DISABLED

            self.DeviceButtonInstalledDate['state'] = DISABLED
            self.DeviceButtonActivatedDate['state'] = DISABLED
            self.DeviceButtonDisconnectedDate['state'] = DISABLED
            self.DeviceButtonExpirationDate['state'] = DISABLED


            self.DeviceIPAddressFrameEntry['state'] = 'normal'
            self.DeviceIPAddressFrameEntry.delete(0,END)
            self.DeviceIPAddressFrameEntry['state'] = 'readonly'

            # Setup Labels and Button Calendars Installed, Activated, Disconnected
            self.DeviceInstalledDateFrameEntry['state'] = 'normal'
            self.DeviceInstalledDateFrameEntry.delete(0,END)
            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'

            self.DeviceContractNoFrameEntry['state'] = 'normal'
            self.DeviceContractNoFrameEntry.delete(0,END)
            self.DeviceContractNoFrameEntry['state'] = 'readonly'

            self.DeviceExpirationDateFrameEntry['state'] = 'normal'
            self.DeviceExpirationDateFrameEntry.delete(0,END)
            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'

            self.DeviceActivatedDateFrameEntry['state'] = 'normal'
            self.DeviceActivatedDateFrameEntry.delete(0,END)
            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'

            self.DeviceSerialNoFrameEntry['state'] = 'normal'
            self.DeviceSerialNoFrameEntry.delete(0,END)
            self.DeviceSerialNoFrameEntry['state'] = 'readonly'

            self.DeviceMACAddressFrameEntry['state'] = 'normal'
            self.DeviceMACAddressFrameEntry.delete(0,END)
            self.DeviceMACAddressFrameEntry['state'] = 'readonly'

            self.DeviceDisconnectedDateFrameEntry['state'] = 'normal'
            self.DeviceDisconnectedDateFrameEntry.delete(0,END)
            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'

            self.DeviceOutSourceCostFrameEntry['state'] = 'normal'
            self.DeviceOutSourceCostFrameEntry.delete(0,END)
            self.DeviceOutSourceCostFrameEntry['state'] = 'readonly'

            self.DeviceMaintenanceCostFrameEntry['state'] = 'normal'
            self.DeviceMaintenanceCostFrameEntry.delete(0,END)
            self.DeviceMaintenanceCostFrameEntry['state'] = 'readonly'

            self.DeviceNotesFrameEntry['state'] = 'normal'
            self.DeviceNotesFrameEntry.delete(0,END)
            self.DeviceNotesFrameEntry['state'] = 'readonly'


            ##############################

            self.DeviceNATIPAddressFrameEntry['state'] = 'normal'
            self.DeviceNATIPAddressFrameEntry.delete(0,END)
            self.DeviceNATIPAddressFrameEntry['state'] = 'readonly'
            
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'normal'
            self.DeviceMgmtIPAddressFrameEntry.delete(0,END)
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'readonly'
                  
            self.DeviceComboBoxICMPCapable.current(0)
            self.DeviceComboBoxICMPCapable['state'] = DISABLED

            self.DeviceComboBoxICMPMonitor.current(0)
            self.DeviceComboBoxICMPMonitor['state'] = DISABLED

            self.DeviceLastICMPDateFrameEntry['state'] = 'normal'
            self.DeviceLastICMPDateFrameEntry.delete(0,END)
            self.DeviceLastICMPDateFrameEntry['state'] = 'readonly'

            self.DeviceLastICMPStatusFrameEntry['state'] = 'normal'
            self.DeviceLastICMPStatusFrameEntry.delete(0,END)
            self.DeviceLastICMPStatusFrameEntry['state'] = 'readonly'

            self.DeviceLastCMDBStatusFrameEntry['state'] = 'normal'
            self.DeviceLastCMDBStatusFrameEntry.delete(0,END)
            self.DeviceLastCMDBStatusFrameEntry['state'] = 'readonly'
            
            ##############################

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.DeviceTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.DeviceTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

    def Display_Screen(self,curItem):      
            # Create Progress Bar
            self.Get_Type_Model_and_Satus()
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            '''
                        self.DeviceTablePriaryKeyArray[i],
                        #self.DeviceTableDeviceDescriptionArray[i],                
                        #self.DeviceTableCountryIDArray[i],
                        #self.DeviceTableRegionIDArray[i],
                        #self.DeviceTableFacilityIDArray[i],
                        #self.DeviceTableSiteIDArray[i],
                        self.DeviceTableDeviceTypeIDArray[i],
                        self.DeviceTableIP4AddressArray[i],
                        #self.DeviceTableIP6AddressArray[i],
                        self.DeviceTableContractNoArray[i],
                        self.DeviceTableOutsourceCostArray[i],
                        self.DeviceTableMaintenanceCostArray[i],
                        self.DeviceTableStatusArray[i],
                        #self.DeviceTableMonthlyCostArray[i],
                        self.DeviceTableDateInstalledArray[i],
                        #self.DeviceTableDayInstalledArray[i],
                        #self.DeviceTableMonthInstalledArray[i],
                        #self.DeviceTableYearInstalledArray[i],Get_Type
                        self.DeviceTableDateActivatedArray[i],
                        #self.DeviceTableDayActivatedArray[i],
                        #self.DeviceTableMonthActivatedArray[i],
                        #self.DeviceTableYearActivatedArray[i],
                        self.DeviceTableDisconectedDateArray[i],
                        #self.DeviceTableDayDisconectedArray[i],
                        #self.DeviceTableMonthDisconectedArray[i],
                        #self.DeviceTableYearDisconectedArray[i],
                        self.DeviceTableExpirationDateArray[i],
                        #self.DeviceTableDayExpirationArray[i],
                        #self.DeviceTableMonthExpirationArray[i],
                        #self.DeviceTableYearExpirationArray[i],
                        self.DeviceTableSerilaNoArray[i],
                        self.DeviceTableExecutedByArray[i],
                        #self.DeviceTableNotesArray[i],
                        self.DeviceTableDeviceModelIDArray[i],
                        #self.DeviceTableMACAddressArray[i]
            '''
            # Setup Labels and Entry
            self.DeviceIDFrameEntry['state'] = 'normal'
            self.DeviceIDFrameEntry.delete(0,END)
            self.DeviceIDFrameEntry.insert(0,self.DeviceTablePriaryKeyArray[curItem])
            self.DeviceIDFrameEntry['state'] = 'readonly'
                
            self.DeviceDescriptionFrameEntry['state'] = 'normal'
            self.DeviceDescriptionFrameEntry.delete(0,END)
            self.DeviceDescriptionFrameEntry.insert(0,self.DeviceTableDeviceDescriptionArray[curItem])
            self.DeviceDescriptionFrameEntry['state'] = 'readonly'

            # Find Type in the Array
            i = 0
            self.DeviceComboBoxTypeID.current(i)
            while (i < len(self.DeviceTypeIDArray)):
                if (self.DeviceTableDeviceTypeIDArray[curItem]== self.DeviceTypeIDArray[i]):
                    self.DeviceComboBoxTypeID.current(i)
                    i = i + len(self.DeviceTypeIDArray) 
                else:
                    i = i + 1
            # find Model in the Array
            i = 0
            self.DeviceComboBoxModelID.current(i)
            while (i < len(self.DeviceModelIDArray)):
                if (self.DeviceTableDeviceModelIDArray[curItem]== self.DeviceModelIDArray[i]):
                    self.DeviceComboBoxModelID.current(i)
                    i = i + len(self.DeviceModelIDArray) 
                else:
                    i = i + 1            # find Status in the Array
            i = 0
            while (i < len(self.DeviceStatusValues)):
                if (self.DeviceTableStatusArray[curItem]== self.DeviceStatusValues[i]):
                    self.DeviceComboBoxStatus.current(i)
                    i = i + len(self.DeviceStatusValues) 
                else:
                    i = i + 1

            self.DeviceButtonInstalledDate['state'] = DISABLED
            self.DeviceButtonActivatedDate['state'] = DISABLED
            self.DeviceButtonDisconnectedDate['state'] = DISABLED
            self.DeviceButtonExpirationDate['state'] = DISABLED


            self.DeviceIPAddressFrameEntry['state'] = 'normal'
            self.DeviceIPAddressFrameEntry.delete(0,END)
            self.DeviceIPAddressFrameEntry.insert(0,self.DeviceTableIP4AddressArray[curItem])
            self.DeviceIPAddressFrameEntry['state'] = 'readonly'

            # Setup Labels and Button Calendars Installed, Activated, Disconnected
            self.DeviceInstalledDateFrameEntry['state'] = 'normal'
            self.DeviceInstalledDateFrameEntry.delete(0,END)
            if (self.DeviceTableDateInstalledArray[curItem] == None):
                self.DeviceInstalledDateFrameEntry.insert(0," ")
            else:
                self.DeviceInstalledDateFrameEntry.insert(0,self.DeviceTableDateInstalledArray[curItem])
            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'

            self.DeviceContractNoFrameEntry['state'] = 'normal'
            self.DeviceContractNoFrameEntry.delete(0,END)
            if (self.DeviceTableContractNoArray[curItem] == None):
                self.DeviceContractNoFrameEntry.insert(0," ")
            else:
                self.DeviceContractNoFrameEntry.insert(0,self.DeviceTableContractNoArray[curItem])
            self.DeviceContractNoFrameEntry['state'] = 'readonly'

            self.DeviceExpirationDateFrameEntry['state'] = 'normal'
            self.DeviceExpirationDateFrameEntry.delete(0,END)
            if (self.DeviceTableExpirationDateArray[curItem] == None):
                self.DeviceExpirationDateFrameEntry.insert(0," ")
            else:
                self.DeviceExpirationDateFrameEntry.insert(0,self.DeviceTableExpirationDateArray[curItem])
            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'

            self.DeviceActivatedDateFrameEntry['state'] = 'normal'
            self.DeviceActivatedDateFrameEntry.delete(0,END)
            if (self.DeviceTableDateActivatedArray[curItem] == None):
                self.DeviceActivatedDateFrameEntry.insert(0," ")
            else:
                self.DeviceActivatedDateFrameEntry.insert(0,self.DeviceTableDateActivatedArray[curItem])
            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'

            self.DeviceSerialNoFrameEntry['state'] = 'normal'
            self.DeviceSerialNoFrameEntry.delete(0,END)
            if (self.DeviceTableSerilaNoArray[curItem] == None):
                self.DeviceSerialNoFrameEntry.insert(0," ")
            else:
                self.DeviceSerialNoFrameEntry.insert(0,self.DeviceTableSerilaNoArray[curItem])
            self.DeviceSerialNoFrameEntry['state'] = 'readonly'

            self.DeviceMACAddressFrameEntry['state'] = 'normal'
            self.DeviceMACAddressFrameEntry.delete(0,END)
            if (self.DeviceTableMACAddressArray[curItem] == None):
                self.DeviceMACAddressFrameEntry.insert(0," ")
            else:
                self.DeviceMACAddressFrameEntry.insert(0,self.DeviceTableMACAddressArray[curItem])
            self.DeviceMACAddressFrameEntry['state'] = 'readonly'

            self.DeviceDisconnectedDateFrameEntry['state'] = 'normal'
            self.DeviceDisconnectedDateFrameEntry.delete(0,END)
            if (self.DeviceTableDisconectedDateArray[curItem] == None):
                self.DeviceTableDisconectedDateArray.insert(0," ")
            else:
                self.DeviceDisconnectedDateFrameEntry.insert(0,self.DeviceTableDisconectedDateArray[curItem])
            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'

            self.DeviceOutSourceCostFrameEntry['state'] = 'normal'
            self.DeviceOutSourceCostFrameEntry.delete(0,END)
            self.DeviceOutSourceCostFrameEntry.insert(0,str(self.DeviceTableOutsourceCostArray[curItem]))
            self.DeviceOutSourceCostFrameEntry['state'] = 'readonly'

            self.DeviceMaintenanceCostFrameEntry['state'] = 'normal'
            self.DeviceMaintenanceCostFrameEntry.delete(0,END)
            self.DeviceMaintenanceCostFrameEntry.insert(0,str(self.DeviceTableMaintenanceCostArray[curItem]))
            self.DeviceMaintenanceCostFrameEntry['state'] = 'readonly'

            self.DeviceNotesFrameEntry['state'] = 'normal'
            self.DeviceNotesFrameEntry.delete(0,END)
            self.DeviceNotesFrameEntry.insert(0,self.DeviceTableNotesArray[curItem])
            self.DeviceNotesFrameEntry['state'] = 'readonly'


            ##############################
            
            self.DeviceNATIPAddressFrameEntry['state'] = 'normal'
            self.DeviceNATIPAddressFrameEntry.delete(0,END)
            self.DeviceNATIPAddressFrameEntry.insert(0,self.DeviceTableNATIP4AddressArray[curItem])
            self.DeviceNATIPAddressFrameEntry['state'] = 'readonly'
            
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'normal'
            self.DeviceMgmtIPAddressFrameEntry.delete(0,END)
            self.DeviceMgmtIPAddressFrameEntry.insert(0,self.DeviceTableManagementIP4AddressArray[curItem])
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'readonly'
                  
            if (self.DeviceTableICMPCapableArray[curItem] == "YES"):
                self.DeviceComboBoxICMPCapable.current(1)
            else:
                self.DeviceComboBoxICMPCapable.current(0)
            self.DeviceComboBoxICMPCapable['state'] = DISABLED

            if (self.DeviceTableMonitorviaICMPArray[curItem] == "YES"):
                self.DeviceComboBoxICMPMonitor.current(1)
            else:
                self.DeviceComboBoxICMPMonitor.current(0)
            self.DeviceComboBoxICMPMonitor['state'] = DISABLED

            self.DeviceLastICMPDateFrameEntry['state'] = 'normal'
            self.DeviceLastICMPDateFrameEntry.delete(0,END)
            self.DeviceLastICMPDateFrameEntry.insert(0,self.DeviceTableLastSuccessICMPArray[curItem])
            self.DeviceLastICMPDateFrameEntry['state'] = 'readonly'

            self.DeviceLastICMPStatusFrameEntry['state'] = 'normal'
            self.DeviceLastICMPStatusFrameEntry.delete(0,END)
            self.DeviceLastICMPStatusFrameEntry.insert(0,self.DeviceTableLastICMPStatusArray[curItem])
            self.DeviceLastICMPStatusFrameEntry['state'] = 'readonly'

            self.DeviceLastCMDBStatusFrameEntry['state'] = 'normal'
            self.DeviceLastCMDBStatusFrameEntry.delete(0,END)
            self.DeviceLastCMDBStatusFrameEntry.insert(0,self.DeviceTableLastUpdatedCMDBDateArray[curItem])
            self.DeviceLastCMDBStatusFrameEntry['state'] = 'readonly'
            
            ##############################


    def Enable_Screen(self,option):
        # This function is used when the ADD button is selected

        #self.DeviceBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = DISABLED
            self.ButtonCountryRefresh['state'] = DISABLED
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = DISABLED
            self.ButtonRegionRefresh['state'] = DISABLED
        if (Is_Facility_Available):
            self.ButtonFacilityAdd['state'] = DISABLED
            self.ButtonFacilityRefresh['state'] = DISABLED
        if (Is_Sites_Available):
            self.ButtonSitesAdd['state'] = DISABLED
            self.ButtonSitesRefresh['state'] = DISABLED

            
        self.ButtonDeviceAdd['state'] = DISABLED
        self.ButtonDeviceEdit['state'] = DISABLED
        self.ButtonDeviceRemove['state'] = DISABLED
        self.ButtonDeviceOK['state'] = ACTIVE
        self.ButtonDeviceCancel['state'] = ACTIVE
        self.ButtonDeviceCircuits['state'] = DISABLED
        self.ButtonDevicePing64['state'] = DISABLED
        self.ButtonDevicePing1500['state'] = DISABLED
        #self.ButtonDeviceContacts['state'] = DISABLED
        self.ButtonDeviceICMP['state'] = DISABLED
        self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        if (option == 'add'): #<----------------------------------- ADD Button
            self.ComboBoxCoutryID['state'] = DISABLED
            self.ComboBoxRegionID['state'] = DISABLED
            self.ComboBoxFacilityID['state'] = DISABLED
            self.ComboBoxSitesID['state'] = DISABLED

            self.Get_Type_Model_and_Satus()
            self.DeviceIDFrameEntry['state'] = 'normal'
            self.DeviceIDFrameEntry.delete(0,END)

            self.DeviceDescriptionFrameEntry['state'] = 'normal'
            self.DeviceDescriptionFrameEntry.delete(0,END)

            # Calendars:
            self.DeviceButtonInstalledDate['state'] = ACTIVE
            self.DeviceButtonActivatedDate['state'] = ACTIVE
            self.DeviceButtonDisconnectedDate['state'] = ACTIVE
            self.DeviceButtonExpirationDate['state'] = ACTIVE

            self.DeviceComboBoxTypeID['state'] = DISABLED
            self.DeviceComboBoxModelID['state'] = DISABLED
            self.DeviceComboBoxStatus['state'] = DISABLED

            self.DeviceInstalledDateFrameEntry['state'] = 'normal'
            self.DeviceInstalledDateFrameEntry.delete(0,END)
            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'
            
            self.DeviceActivatedDateFrameEntry['state'] = 'normal'
            self.DeviceActivatedDateFrameEntry.delete(0,END)
            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
            
            self.DeviceDisconnectedDateFrameEntry['state'] = 'normal'
            self.DeviceDisconnectedDateFrameEntry.delete(0,END)
            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'

            self.DeviceExpirationDateFrameEntry['state'] = 'normal'
            self.DeviceExpirationDateFrameEntry.delete(0,END)
            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'

            self.DeviceComboBoxTypeID['state'] = 'readonly'
            self.DeviceComboBoxModelID['state'] = 'disabled'
            self.DeviceComboBoxStatus['state'] = 'readonly'

            self.DeviceIPAddressFrameEntry['state'] = 'normal'
            self.DeviceIPAddressFrameEntry.delete(0,END)
            self.DeviceIPAddressFrameEntry.insert(0,'0.0.0.0')

            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'
            self.DeviceInstalledDateFrameEntry.delete(0,END)

            self.DeviceContractNoFrameEntry['state'] = 'normal'
            self.DeviceContractNoFrameEntry.delete(0,END)

            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'
            self.DeviceExpirationDateFrameEntry.delete(0,END)

            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
            self.DeviceActivatedDateFrameEntry.delete(0,END)

            self.DeviceSerialNoFrameEntry['state'] = 'normal'
            self.DeviceSerialNoFrameEntry.delete(0,END)

            self.DeviceMACAddressFrameEntry['state'] = 'normal'
            self.DeviceMACAddressFrameEntry.delete(0,END)

            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'
            self.DeviceDisconnectedDateFrameEntry.delete(0,END)

            self.DeviceOutSourceCostFrameEntry['state'] = 'normal'
            self.DeviceOutSourceCostFrameEntry.delete(0,END)

            self.DeviceMaintenanceCostFrameEntry['state'] = 'normal'
            self.DeviceMaintenanceCostFrameEntry.delete(0,END)

            self.DeviceNotesFrameEntry['state'] = 'normal'
            self.DeviceNotesFrameEntry.delete(0,END)

            ##############################
            
            self.DeviceNATIPAddressFrameEntry['state'] = 'normal'
            self.DeviceNATIPAddressFrameEntry.delete(0,END)
            self.DeviceNATIPAddressFrameEntry.insert(0,'0.0.0.0')
            
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'normal'
            self.DeviceMgmtIPAddressFrameEntry.delete(0,END)
            self.DeviceMgmtIPAddressFrameEntry.insert(0,'0.0.0.0')
                  
            self.DeviceComboBoxICMPCapable.current(1)
            self.DeviceComboBoxICMPCapable['state'] = ACTIVE

            self.DeviceComboBoxICMPMonitor.current(0)
            self.DeviceComboBoxICMPMonitor['state'] = ACTIVE

            self.DeviceLastICMPDateFrameEntry['state'] = 'normal'
            self.DeviceLastICMPDateFrameEntry.delete(0,END)
            self.DeviceLastICMPDateFrameEntry.insert(0,'Never')
            self.DeviceLastICMPDateFrameEntry['state'] = 'readonly'

            self.DeviceLastICMPStatusFrameEntry['state'] = 'normal'
            self.DeviceLastICMPStatusFrameEntry.delete(0,END)
            self.DeviceLastICMPStatusFrameEntry.insert(0,'Never')
            self.DeviceLastICMPStatusFrameEntry['state'] = 'readonly'

            self.DeviceLastCMDBStatusFrameEntry['state'] = 'normal'
            self.DeviceLastCMDBStatusFrameEntry.delete(0,END)
            self.DeviceLastCMDBStatusFrameEntry.insert(0,'Never')
            self.DeviceLastCMDBStatusFrameEntry['state'] = 'readonly'
            
            ##############################


        if (option == 'edit'): #<----------------------------------- EDIT Button
            self.ComboBoxCoutryID['state'] = ACTIVE
            self.ComboBoxRegionID['state'] = ACTIVE
            self.ComboBoxFacilityID['state'] = ACTIVE
            self.ComboBoxSitesID['state'] = ACTIVE

            #self.Get_Type_Model_and_Satus() # <------------------ I might have to modified it
            self.DeviceIDFrameEntry['state'] = 'readonly'
            self.DeviceDescriptionFrameEntry['state'] = 'normal'
            # Calendars:
            self.DeviceButtonInstalledDate['state'] = ACTIVE
            self.DeviceButtonActivatedDate['state'] = ACTIVE
            self.DeviceButtonDisconnectedDate['state'] = ACTIVE
            self.DeviceButtonExpirationDate['state'] = ACTIVE

            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'            
            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'
            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'

            self.DeviceComboBoxTypeID['state'] = 'readonly'
            self.DeviceComboBoxModelID['state'] = 'disabled'
            self.DeviceComboBoxStatus['state'] = 'readonly'

            self.DeviceIPAddressFrameEntry['state'] = 'normal'
            self.DeviceInstalledDateFrameEntry['state'] = 'readonly'
            self.DeviceContractNoFrameEntry['state'] = 'normal'
            self.DeviceExpirationDateFrameEntry['state'] = 'readonly'
            self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
            self.DeviceSerialNoFrameEntry['state'] = 'normal'
            self.DeviceMACAddressFrameEntry['state'] = 'normal'
            self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'
            self.DeviceOutSourceCostFrameEntry['state'] = 'normal'
            self.DeviceMaintenanceCostFrameEntry['state'] = 'normal'
            self.DeviceNotesFrameEntry['state'] = 'normal'


            ##############################

            self.DeviceNATIPAddressFrameEntry['state'] = 'normal'            
            self.DeviceMgmtIPAddressFrameEntry['state'] = 'normal'
            self.DeviceComboBoxICMPCapable['state'] = ACTIVE
            self.DeviceComboBoxICMPMonitor['state'] = ACTIVE
            self.DeviceLastICMPDateFrameEntry['state'] = 'readonly'
            self.DeviceLastICMPStatusFrameEntry['state'] = 'readonly'
            self.DeviceLastCMDBStatusFrameEntry['state'] = 'readonly'
            
            ##############################
                
       
    def Disable_Screen(self):
        # This function is used when the entry was added.modified to the Database

        #self.DeviceBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

        self.ComboBoxCoutryID['state'] = 'readonly'
        self.ComboBoxRegionID['state'] = 'readonly'
        self.ComboBoxFacilityID['state'] = 'readonly'
        self.ComboBoxSitesID['state'] = 'readonly'
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = ACTIVE
            self.ButtonCountryRefresh['state'] = ACTIVE
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = ACTIVE
            self.ButtonRegionRefresh['state'] = ACTIVE
        if (Is_Facility_Available):
            self.ButtonFacilityAdd['state'] = ACTIVE
            self.ButtonFacilityRefresh['state'] = ACTIVE
        if (Is_Sites_Available):
            self.ButtonSitesAdd['state'] = ACTIVE
            self.ButtonSitesRefresh['state'] = ACTIVE
            
        self.ButtonDeviceAdd['state'] = ACTIVE
        self.ButtonDeviceEdit['state'] = DISABLED
        self.ButtonDeviceRemove['state'] = DISABLED
        self.ButtonDeviceOK['state'] = DISABLED
        self.ButtonDeviceCancel['state'] = DISABLED # ACTIVE
        self.ButtonDeviceCircuits['state'] = DISABLED
        self.ButtonDevicePing64['state'] = DISABLED
        self.ButtonDevicePing1500['state'] = DISABLED
        #self.ButtonDeviceContacts['state'] = DISABLED
        self.ButtonDeviceICMP['state'] = DISABLED
        self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.DeviceIDFrameEntry['state'] = 'readonly'                
        self.DeviceDescriptionFrameEntry['state'] = 'readonly'
        self.DeviceComboBoxTypeID['state'] = 'readonly'
        self.DeviceComboBoxModelID['state'] = 'readonly'
        self.DeviceComboBoxStatus['state'] = 'readonly'
        self.DeviceIPAddressFrameEntry['state'] = 'readonly'
        self.DeviceInstalledDateFrameEntry['state'] = 'readonly'
        self.DeviceContractNoFrameEntry['state'] = 'readonly'
        self.DeviceExpirationDateFrameEntry['state'] = 'readonly'
        self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
        self.DeviceSerialNoFrameEntry['state'] = 'readonly'
        self.DeviceMACAddressFrameEntry['state'] = 'readonly'
        self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'
        self.DeviceOutSourceCostFrameEntry['state'] = 'readonly'
        self.DeviceMaintenanceCostFrameEntry['state'] = 'readonly'
        self.DeviceNotesFrameEntry['state'] = 'readonly'
        # Calendars:
        self.DeviceButtonInstalledDate['state'] = DISABLED
        self.DeviceButtonActivatedDate['state'] = DISABLED
        self.DeviceButtonDisconnectedDate['state'] = DISABLED
        self.DeviceButtonExpirationDate['state'] = DISABLED

        self.DeviceComboBoxTypeID['state'] = DISABLED
        self.DeviceComboBoxModelID['state'] = DISABLED
        self.DeviceComboBoxStatus['state'] = DISABLED

        ##############################

        self.DeviceNATIPAddressFrameEntry['state'] = 'readonly'            
        self.DeviceMgmtIPAddressFrameEntry['state'] = 'readonly'
        self.DeviceComboBoxICMPCapable['state'] = DISABLED
        self.DeviceComboBoxICMPMonitor['state'] = DISABLED
        self.DeviceLastICMPDateFrameEntry['state'] = 'readonly'
        self.DeviceLastICMPStatusFrameEntry['state'] = 'readonly'
        self.DeviceLastCMDBStatusFrameEntry['state'] = 'readonly'
            
        ##############################

        

    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]
        
        self.DeviceID = self.DeviceIDFrameEntry.get()
        self.DeviceDescription = self.DeviceDescriptionFrameEntry.get()  

        # Calendars:
        self.DeviceButtonInstalledDate['state'] = ACTIVE
        self.DeviceButtonActivatedDate['state'] = ACTIVE
        self.DeviceButtonDisconnectedDate['state'] = ACTIVE
        self.DeviceButtonExpirationDate['state'] = ACTIVE

        self.DeviceInstalledDate = self.DeviceInstalledDateFrameEntry.get()        
        if (len(self.DeviceInstalledDate) > 0):
            self.DeviceInstalledDate = str(self.DeviceInstalledData['month_selected']) + '/' + str(self.DeviceInstalledData['day_selected']) + '/' + str(self.DeviceInstalledData['year_selected'])
            self.DeviceInstalledMonth = self.DeviceInstalledData['month_selected']
            self.DeviceInstalledDay = self.DeviceInstalledData['day_selected']
            self.DeviceInstalledYear = self.DeviceInstalledData['year_selected']
        else:
            self.DeviceInstalledDate = ""
            self.DeviceInstalledMonth = "0"
            self.DeviceInstalledDay = "0"
            self.DeviceInstalledYear = "0"

        self.DeviceActivatedDate = self.DeviceActivatedDateFrameEntry.get()        
        if (len(self.DeviceActivatedDate) > 0):
            self.DeviceActivatedDate = str(self.DeviceActivatedData['month_selected']) + '/' + str(self.DeviceActivatedData['day_selected']) + '/' + str(self.DeviceActivatedData['year_selected'])
            self.DeviceActivatedMonth = self.DeviceActivatedData['month_selected']
            self.DeviceActivatedDay = self.DeviceActivatedData['day_selected']
            self.DeviceActivatedYear = self.DeviceActivatedData['year_selected']
        else:
            self.DeviceActivatedDate = ""
            self.DeviceActivatedMonth = "0"
            self.DeviceActivatedDay = "0"
            self.DeviceActivatedYear = "0"

        self.DeviceDisconnectedDate = self.DeviceDisconnectedDateFrameEntry.get()        
        if (len(self.DeviceDisconnectedDate) > 0):
            self.DeviceDisconnectedDate = str(self.DeviceDisconnectedData['month_selected']) + '/' + str(self.DeviceDisconnectedData['day_selected']) + '/' + str(self.DeviceDisconnectedData['year_selected'])
            self.DeviceDisconnectedMonth = self.DeviceDisconnectedData['month_selected']
            self.DeviceDisconnectedDay = self.DeviceDisconnectedData['day_selected']
            self.DeviceDisconnectedYear = self.DeviceDisconnectedData['year_selected']
        else:
            self.DeviceDisconnectedDate = ""
            self.DeviceDisconnectedMonth = "0"
            self.DeviceDisconnectedDay = "0"
            self.DeviceDisconnectedYear = "0"

        self.DeviceExpirationDate = self.DeviceExpirationDateFrameEntry.get()        
        if (len(self.DeviceExpirationDate) > 0):
            self.DeviceExpirationDate = str(self.DeviceExpirationData['month_selected']) + '/' + str(self.DeviceExpirationData['day_selected']) + '/' + str(self.DeviceExpirationData['year_selected'])
            self.DeviceExpirationMonth = self.DeviceExpirationData['month_selected']
            self.DeviceExpirationDay = self.DeviceExpirationData['day_selected']
            self.DeviceExpirationYear = self.DeviceExpirationData['year_selected']
        else:
            self.DeviceExpirationDate = ""
            self.DeviceExpirationMonth = "0"
            self.DeviceExpirationDay = "0"
            self.DeviceExpirationYear = "0"
            
        self.DeviceTypeID = self.DeviceTypeIDArray[self.DeviceComboBoxTypeID.current()] 
        self.DeviceModelID = self.DeviceModelIDArray[self.DeviceComboBoxModelID.current()]
        self.DeviceStatus = self.DeviceStatusValues[self.DeviceComboBoxStatus.current()]
        self.DeviceIPAddress = self.DeviceIPAddressFrameEntry.get()            
        self.DeviceContract = self.DeviceContractNoFrameEntry.get()
        self.DeviceSerialNo = self.DeviceSerialNoFrameEntry.get()
        self.DeviceMACAddress = self.DeviceMACAddressFrameEntry.get()
        if (len(self.DeviceOutSourceCostFrameEntry.get()) > 0):
            self.DeviceOutSourceCost = float(self.DeviceOutSourceCostFrameEntry.get())
        else:
            self.DeviceOutSourceCost = 0
        if (len(self.DeviceMaintenanceCostFrameEntry.get()) > 0):    
            self.DeviceMaintenanceCost = float(self.DeviceMaintenanceCostFrameEntry.get())
        else:
            self.DeviceMaintenanceCost = 0
        self.DeviceNotes = self.DeviceNotesFrameEntry.get()


        self.DeviceTableNATIP4Address = self.DeviceNATIPAddressFrameEntry.get()           
        self.DeviceTableManagementIP4Address = self.DeviceMgmtIPAddressFrameEntry.get()
        self.DeviceTableICMPCapable = self.DeviceICMPValues[self.DeviceComboBoxICMPCapable.current()]
        self.DeviceTableMonitorviaICMP = self.DeviceICMPValues[self.DeviceComboBoxICMPMonitor.current()]
        self.DeviceTableLastSuccessICMP = self.DeviceLastICMPDateFrameEntry.get()
        self.DeviceTableLastICMPStatus = self.DeviceLastICMPStatusFrameEntry.get()
        self.DeviceTableLastUpdatedCMDBDate = self.DeviceLastCMDBStatusFrameEntry.get()        
        # self.DeviceTableLastUpdatedCMDBDay
        # self.DeviceTableLastUpdatedCMDBMonth
        # self.DeviceTableLastUpdatedCMDBYear


#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <END> ------------------------*
#****************************************************************************************


#****************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************

    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Country Window']    
            Logging.Log(Parameter)

        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        #print (event)
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # SQL Querry to the Device Table
            sql = """
                    SELECT * FROM Region
                       WHERE Country_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            #print (sql)
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = ACTIVE
                self.ButtonRegionRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                #print ("found it")
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
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Country Refresh']    
            Logging.Log(Parameter)

        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the Device Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Country
                        WHERE Country_ID = '%s'
                      """ % (self.CountryID_Pre)
            else:
                sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                self.sql_querry = True
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
                    #self.ComboBoxCoutryID['state'] = 'readonly'
                    #self.ComboBoxCoutryID.set("")
                    self.Clean_Screen('country','all')
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Country Records found')
                self.sql_querry = False
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
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
            Parameter = ['Device','Region Window']    
            Logging.Log(Parameter)

        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        self.Clean_Screen('region','region-combo')
        if self.db.Connect():
            # SQL Querry to the Device Table
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                    ORDER BY Facility_Name ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            '''
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
                        ORDER BY Facility_Name ASC
                      """ % (self.CountryID_Pre,self.RegionID_Pre)
            else:
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s'
                        ORDER BY Facility_Name ASC
                      """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            '''
            #print (sql)
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                #print ("found it")
                self.sql_querry = True
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
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
                self.sql_querry = False
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Region Refresh']    
            Logging.Log(Parameter)

        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the Device Table
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
                        ORDER BY Region_Name ASC
                      """ % (self.CountryID_Pre,self.RegionID_Pre)
            else:
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s'
                        ORDER BY Region_Name ASC
                      """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                self.sql_querry = True
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
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
                self.sql_querry = False
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- REGION SELECTION SECTION <END> ------------------------*
#*************************************************************************************


#***************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <BEGIN> ----------------------*
#***************************************************************************************
        
    def Display_Facility_Window(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Facility Windows']    
            Logging.Log(Parameter)

        Facility.Display_Facility_Window()

    def on_facility_combo_changed(self,event):
        self.Clean_Screen('facility','facility-combo')
        if self.db.Connect():
            # SQL Querry to the Device Table
            sql = """
                    SELECT * FROM Sites
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                             self.FacilityIDArray[self.ComboBoxFacilityID.current()])
            #print (sql)
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = ACTIVE
                self.ButtonSitesRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                self.SitesIDArray =  []
                self.SitesNameArray =  []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.SitesIDArray.append(self.db.results[i][4].strip())
                    self.SitesNameArray.append(self.db.results[i][5].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxSitesID['values'] = self.SitesNameArray
                if (len(self.SitesNameArray)== 0):
                    self.ComboBoxSitesID['state'] = DISABLED
                else:
                    self.ComboBoxSitesID['state'] = 'readonly'
                    self.ComboBoxSitesID.set("")
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Facility Refresh']    
            Logging.Log(Parameter)

        self.Clean_Screen('facility','all')
        if self.db.Connect():
            # SQL Querry to the Device Table
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                        ORDER BY Facility_Name ASC
                      """ % (self.CountryID_Pre,self.RegionID_Pre,self.FacilityID_Pre)
            else:
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
                        ORDER BY Facility_Name ASC
                      """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],
                             self.RegionIDArray[self.ComboBoxRegionID.current()])
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                self.sql_querry = True
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
                    self.ComboBoxSiteID['state'] = DISABLED
                else:
                    self.ComboBoxFacilityID['state'] = 'readonly'
                    self.ComboBoxFacilityID.set("")
                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd['state'] = 'active'
                        self.ButtonFacilityRefresh['state'] = 'active'                        
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
                self.sql_querry = False
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <END> ------------------------*
#*************************************************************************************


#***************************************************************************************
#---------------------------- SITES SELECTION SECTION <BEGIN> ------------------------*
#***************************************************************************************
        
    def Display_Sites_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Sites Window']    
            Logging.Log(Parameter)

        Sites.Display_Sites_Window()

    def on_sites_combo_changed(self,event):
        self.Clean_Screen('sites','sites-combo')
        if self.db.Connect():
            # SQL Querry to the Device Table
            sql = """
                    SELECT * FROM Devices
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' AND Site_ID = '%s'
                    ORDER BY Status desc, Device_Type_ID ASC, Device_Model_ID ASC, Device_ID ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                             self.FacilityIDArray[self.ComboBoxFacilityID.current()],self.SitesIDArray[self.ComboBoxSitesID.current()])
            #print (sql)
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = ACTIVE
                self.ButtonSitesRefresh['state'] = ACTIVE
            if ((self.db.Execute(sql)) and (self.Selection != 'edit')):
                #------------------------------- Deleting Tree View <BEGIN> --------
                #x = self.DeviceTreeview.get_children()
                #if x != '()': # checks if there is something in the first row
                #    for child in x:
                #       #print (child)
                #        self.DeviceTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                #-------------- Initializing Arrays <BEGIN> ----------------------
                self.DeviceTablePriaryKeyArray =  [] # Device ID
                self.DeviceTableDeviceDescriptionArray =  []                
                self.DeviceTableCountryIDArray =  [] 
                self.DeviceTableRegionIDArray =  []
                self.DeviceTableFacilityIDArray =  []
                self.DeviceTableSiteIDArray =  []
                self.DeviceTableDeviceTypeIDArray =  []
                self.DeviceTableIP4AddressArray = []
                self.DeviceTableIP6AddressArray = []
                self.DeviceTableContractNoArray = []
                self.DeviceTableOutsourceCostArray = []
                self.DeviceTableMaintenanceCostArray = []
                self.DeviceTableStatusArray = []
                self.DeviceTableMonthlyCostArray = []
                self.DeviceTableDateInstalledArray = []
                self.DeviceTableDayInstalledArray = []
                self.DeviceTableMonthInstalledArray = []
                self.DeviceTableYearInstalledArray = []
                self.DeviceTableDateActivatedArray = []
                self.DeviceTableDayActivatedArray = []
                self.DeviceTableMonthActivatedArray = []
                self.DeviceTableYearActivatedArray = []
                self.DeviceTableDisconectedDateArray = []
                self.DeviceTableDayDisconectedArray = []
                self.DeviceTableMonthDisconectedArray = []
                self.DeviceTableYearDisconectedArray = []
                self.DeviceTableExpirationDateArray = []
                self.DeviceTableDayExpirationArray = []
                self.DeviceTableMonthExpirationArray = []
                self.DeviceTableYearExpirationArray = []
                self.DeviceTableSerilaNoArray = []
                self.DeviceTableExecutedByArray = []
                self.DeviceTableNotesArray = []
                self.DeviceTableDeviceModelIDArray = []
                self.DeviceTableMACAddressArray = []
                self.DeviceTableNATIP4AddressArray = []    
                self.DeviceTableManagementIP4AddressArray = []
                self.DeviceTableLastSuccessICMPArray = []
                self.DeviceTableLastICMPStatusArray = []
                self.DeviceTableICMPCapableArray = []
                self.DeviceTableMonitorviaICMPArray = []
                self.DeviceTableLastUpdatedCMDBDateArray = []
                self.DeviceTableLastUpdatedCMDBDayArray = []
                self.DeviceTableLastUpdatedCMDBMonthArray = []
                self.DeviceTableLastUpdatedCMDBYearArray = []
                self.DeviceTableArrayColumns = []
                self.results = []
                self.DeviceTableArratRowsTemp = []


                self.DeviceTableArrayColumns = (
                    'Device ID',
                    'Device Description',
                    'Country ID',                    
                    'Region ID',
                    'Facility ID',
                    'Site ID',
                    'Device Type ID',
                    'IP4 Address',
                    'IP6 Address',
                    'Contract No',
                    'OutSource Cost',
                    'Maintenance Cost',
                    'Status',
                    'Monthly Cost',
                    'Date Installed',
                    'Day Installed',
                    'Month Installed',
                    'Year Installed',
                    'Date Activated',
                    'Day Activated',
                    'Month Activated',
                    'Year Activated',
                    'Disconnect Date',
                    'Day Disconnect',
                    'Month Disconnect',
                    'Year Disconnect',
                    'Expiration Date',
                    'Day Expiration',
                    'Month Expiration',
                    'Year Expiration',
                    'Serial No',
                    'Executed by UserID',
                    'Notes',
                    'Device Model ID',
                    'MAC Address',
                    'NAT IP4 Address',    
                    'Management IP4 Address',
                    'Last Success ICMP',
                    'Last ICMP Status',
                    'ICMP Capable',
                    'Monitor via ICMP',
                    'Last Updated CMDB Date',
                    'Last Updated CMDB Day',
                    'Last Updated CMDB Month',
                    'Last Updated CMDB Year'
                    )


                '''
                    0  Device_ID               CHAR(50) NOT NULL PRIMARY KEY,
                    1  Device_Description      CHAR(100) NOT NULL,
                    2  Country_ID              CHAR(20) NOT NULL,                    
                    3  Region_ID               CHAR(20) NOT NULL,
                    4  Facility_ID             CHAR(20) NOT NULL,
                    5  Site_ID                 CHAR(20) NOT NULL,
                    6  Device_Type_ID          CHAR(30) NOT NULL,
                    7  IP4_Address             CHAR(20),
                    8  IP6_Address             CHAR(200),
                    9  Contract_No             CHAR(20),
                    10 OutSource_Cost          FLOAT,
                    11 Maintenance_Cost        FLOAT,
                    12 Status                  CHAR(20),
                    13 Monthly_Cost            FLOAT,
                    14 Date_Installed          CHAR(20),
                    15 Day_Installed           INT,
                    16 Month_Installed         INT,
                    17 Year_Installed          INT,
                    18 Date_Activated          CHAR(20),
                    19 Day_Activated           INT,
                    20 Month_Activated         INT,
                    21 Year_Activated          INT,
                    22 Disconnect_Date         CHAR(20),
                    23 Day_Disconnect          INT,
                    24 Month_Disconnect        INT,
                    25 Year_Disconnect         INT,
                    26 Expiration_Date         CHAR(20),
                    27 Day_Expiration          INT,
                    28 Month_Expiration        INT,
                    29 Year_Expiration         INT,
                    30 Serial_No               CHAR(100),
                    31 Executed_by_UserID      CHAR(20),
                    32 Notes                   CHAR(200),
                    33 Device_Model_ID         CHAR(30),
                    34 MAC_Address             CHAR(20)
                    35 NAT_IP4_Address         CHAR(20), <----- NEW as March 5th 2018    
                    36 Management_IP4_Address  CHAR(20),
                    37 Last_Success_ICMP       CHAR(50),
                    38 Last_ICMP_Status        CHAR(50),
                    39 ICMP_Capable            CHAR(10),
                    40 Monitor_via_ICMP        CHAR(10),
                    41 Last_Updated_CMDB_Date  CHAR(50),
                    42 Last_Updated_CMDB_Day   INT,
                    43 Last_Updated_CMDB_Month INT,
                    44 Last_Updated_CMDB_Year  INT)"""
                '''
                #-------------- Initializing Arrays <END> ----------------------                        
                i = 0
                self.data_ready = True
                while (i < len(self.db.results)):
                    self.DeviceTablePriaryKeyArray.append(self.db.results[i][0].strip())
                    self.DeviceTableDeviceDescriptionArray.append(self.db.results[i][1].strip())
                    self.DeviceTableCountryIDArray.append(self.db.results[i][2].strip())
                    self.DeviceTableRegionIDArray.append(self.db.results[i][3].strip())
                    self.DeviceTableFacilityIDArray.append(self.db.results[i][4].strip())
                    self.DeviceTableSiteIDArray.append(self.db.results[i][5].strip())
                    self.DeviceTableDeviceTypeIDArray.append(self.db.results[i][6].strip())
                    if ((self.db.results[i][7]) == None):
                        self.DeviceTableIP4AddressArray.append("0.0.0.0")
                        #self.DeviceTableIP4AddressArray.append(self.db.results[i][7])
                    else:
                        self.DeviceTableIP4AddressArray.append(self.db.results[i][7].strip())
                    if ((self.db.results[i][8]) == None):
                        self.DeviceTableIP6AddressArray.append(self.db.results[i][8])
                    else:
                        self.DeviceTableIP6AddressArray.append(self.db.results[i][8].strip())
                    if ((self.db.results[i][9]) == None):
                        self.DeviceTableContractNoArray.append(self.db.results[i][9])
                    else:  
                        self.DeviceTableContractNoArray.append(self.db.results[i][9].strip())
                        
                    if (self.db.results[i][10] == None):
                        self.DeviceTableOutsourceCostArray.append(0.0)
                    else:
                        self.DeviceTableOutsourceCostArray.append(self.db.results[i][10])
                        
                    if (self.db.results[i][11] == None):
                        self.DeviceTableMaintenanceCostArray.append(0.0)
                    else:
                        self.DeviceTableMaintenanceCostArray.append(self.db.results[i][11])
                    self.DeviceTableStatusArray.append(self.db.results[i][12].strip())
                    if (self.db.results[i][13] == None):
                        self.DeviceTableMonthlyCostArray.append(0.0)
                    else:
                        self.DeviceTableMonthlyCostArray.append(self.db.results[i][13])
                    if (((self.db.results[i][14]) == None) or (self.db.results[i][15] == 0)):
                        self.DeviceTableDateInstalledArray.append("")
                        self.DeviceTableDayInstalledArray.append(0)
                        self.DeviceTableMonthInstalledArray.append(0)
                        self.DeviceTableYearInstalledArray.append(0)
                    else:
                        self.DeviceTableDateInstalledArray.append(self.db.results[i][14].strip())
                        self.DeviceTableDayInstalledArray.append(self.db.results[i][15])
                        self.DeviceTableMonthInstalledArray.append(self.db.results[i][16])
                        self.DeviceTableYearInstalledArray.append(self.db.results[i][17])
                    if (((self.db.results[i][18]) == None) or (self.db.results[i][19] == 0)):
                        self.DeviceTableDateActivatedArray.append("")
                        self.DeviceTableDayActivatedArray.append(0)
                        self.DeviceTableMonthActivatedArray.append(0)
                        self.DeviceTableYearActivatedArray.append(0)
                    else:
                        self.DeviceTableDateActivatedArray.append(self.db.results[i][18].strip())
                        self.DeviceTableDayActivatedArray.append(self.db.results[i][19])
                        self.DeviceTableMonthActivatedArray.append(self.db.results[i][20])
                        self.DeviceTableYearActivatedArray.append(self.db.results[i][21])
                    if (((self.db.results[i][22]) == None) or (self.db.results[i][23] == 0)):
                        self.DeviceTableDisconectedDateArray.append("")
                        self.DeviceTableDayDisconectedArray.append(0)
                        self.DeviceTableMonthDisconectedArray.append(0)
                        self.DeviceTableYearDisconectedArray.append(0)
                    else:
                        self.DeviceTableDisconectedDateArray.append(self.db.results[i][22].strip())
                        self.DeviceTableDayDisconectedArray.append(self.db.results[i][23])
                        self.DeviceTableMonthDisconectedArray.append(self.db.results[i][24])
                        self.DeviceTableYearDisconectedArray.append(self.db.results[i][25])
                    if (((self.db.results[i][26]) == None) or (self.db.results[i][27] == 0)):
                        self.DeviceTableExpirationDateArray.append("")
                        self.DeviceTableDayExpirationArray.append(0)
                        self.DeviceTableMonthExpirationArray.append(0)
                        self.DeviceTableYearExpirationArray.append(0)
                    else:
                        self.DeviceTableExpirationDateArray.append(self.db.results[i][26].strip())
                        self.DeviceTableDayExpirationArray.append(self.db.results[i][27])
                        self.DeviceTableMonthExpirationArray.append(self.db.results[i][28])
                        self.DeviceTableYearExpirationArray.append(self.db.results[i][29])
                    if ((self.db.results[i][30]) == None):
                        self.DeviceTableSerilaNoArray.append(self.db.results[i][30])
                    else:
                        self.DeviceTableSerilaNoArray.append(self.db.results[i][30].strip())
                    self.DeviceTableExecutedByArray.append(self.db.results[i][31].strip())
                    if ((self.db.results[i][32]) == None):
                        self.DeviceTableNotesArray.append(" ")
                        #self.DeviceTableNotesArray.append(self.db.results[i][32])
                    else:
                        self.DeviceTableNotesArray.append(self.db.results[i][32].strip())
                    if ((self.db.results[i][33]) == None):
                        self.DeviceTableDeviceModelIDArray.append(self.db.results[i][33])
                    else:
                        self.DeviceTableDeviceModelIDArray.append(self.db.results[i][33].strip())
                    if ((self.db.results[i][34]) == None):
                        self.DeviceTableMACAddressArray.append(self.db.results[i][34])
                    else:
                        self.DeviceTableMACAddressArray.append(self.db.results[i][34].strip())

                    ###################
                    if ((self.db.results[i][35]) == None):
                        self.DeviceTableNATIP4AddressArray.append("0.0.0.0")
                    else:
                        self.DeviceTableNATIP4AddressArray.append(self.db.results[i][35].strip())
                        
                    if ((self.db.results[i][36]) == None):
                        self.DeviceTableManagementIP4AddressArray.append("0.0.0.0")
                    else:
                        self.DeviceTableManagementIP4AddressArray.append(self.db.results[i][36].strip())

                    if ((self.db.results[i][37]) == None):
                        self.DeviceTableLastSuccessICMPArray.append("Never")
                    else:
                        self.DeviceTableLastSuccessICMPArray.append(self.db.results[i][37].strip())

                    if ((self.db.results[i][38]) == None):
                        self.DeviceTableLastICMPStatusArray.append("Never")
                    else:
                        self.DeviceTableLastICMPStatusArray.append(self.db.results[i][38].strip())

                    if ((self.db.results[i][39]) == None): # this field I toggle between YES / NO on Double Click on ICMP Menu
                        self.DeviceTableICMPCapableArray.append("YES")
                    else:
                        self.DeviceTableICMPCapableArray.append(self.db.results[i][39].strip())

                    if ((self.db.results[i][40]) == None): 
                        self.DeviceTableMonitorviaICMPArray.append("NO")
                    else:
                        self.DeviceTableMonitorviaICMPArray.append(self.db.results[i][40].strip())

                    if ((self.db.results[i][41]) == None): 
                        self.DeviceTableLastUpdatedCMDBDateArray.append("Never")
                    else:
                        self.DeviceTableLastUpdatedCMDBDateArray.append(self.db.results[i][41].strip())

                    if ((self.db.results[i][42]) == None): 
                        self.DeviceTableLastUpdatedCMDBDayArray.append(0)
                    else:
                        self.DeviceTableLastUpdatedCMDBDayArray.append(self.db.results[i][42])

                    if ((self.db.results[i][43]) == None): 
                        self.DeviceTableLastUpdatedCMDBMonthArray.append(0)
                    else:
                        self.DeviceTableLastUpdatedCMDBMonthArray.append(self.db.results[i][43])

                    if ((self.db.results[i][44]) == None): 
                        self.DeviceTableLastUpdatedCMDBYearArray.append(0)
                    else:
                        self.DeviceTableLastUpdatedCMDBYearArray.append(self.db.results[i][44])
                        
                    i = i + 1
                i = 0
                while (i < len(self.DeviceTablePriaryKeyArray)):
                    num = i + 1
                    tags = self.DeviceTableStatusArray[i] # To use in the futire
                    item = [
                            self.DeviceTablePriaryKeyArray[i],
                            self.DeviceTableDeviceDescriptionArray[i],                
                            self.DeviceTableDeviceTypeIDArray[i],
                            self.DeviceTableDeviceModelIDArray[i],
                            self.DeviceTableIP4AddressArray[i],
                            self.DeviceTableOutsourceCostArray[i],
                            self.DeviceTableMaintenanceCostArray[i],                            
                            self.DeviceTableDateInstalledArray[i],
                            self.DeviceTableDateActivatedArray[i],
                            self.DeviceTableDisconectedDateArray[i],
                            self.DeviceTableStatusArray[i],
                            self.DeviceTableContractNoArray[i],
                            self.DeviceTableExpirationDateArray[i],
                            self.DeviceTableSerilaNoArray[i],
                            self.DeviceTableExecutedByArray[i],
                            ]
                    self.DeviceTableArrayRowsTemp = [
                        self.DeviceTablePriaryKeyArray[i],
                        self.DeviceTableDeviceDescriptionArray[i],
                        self.DeviceTableCountryIDArray[i],
                        self.DeviceTableRegionIDArray[i],
                        self.DeviceTableFacilityIDArray[i],
                        self.DeviceTableSiteIDArray[i],
                        self.DeviceTableDeviceTypeIDArray[i],
                        self.DeviceTableIP4AddressArray[i],
                        self.DeviceTableIP6AddressArray[i],
                        self.DeviceTableContractNoArray[i],
                        self.DeviceTableOutsourceCostArray[i],
                        self.DeviceTableMaintenanceCostArray[i],
                        self.DeviceTableStatusArray[i],
                        self.DeviceTableMonthlyCostArray[i],
                        self.DeviceTableDateInstalledArray[i],
                        self.DeviceTableDayInstalledArray[i],
                        self.DeviceTableMonthInstalledArray[i],
                        self.DeviceTableYearInstalledArray[i],
                        self.DeviceTableDateActivatedArray[i],
                        self.DeviceTableDayActivatedArray[i],
                        self.DeviceTableMonthActivatedArray[i],
                        self.DeviceTableYearActivatedArray[i],
                        self.DeviceTableDisconectedDateArray[i],
                        self.DeviceTableDayDisconectedArray[i],
                        self.DeviceTableMonthDisconectedArray[i],
                        self.DeviceTableYearDisconectedArray[i],
                        self.DeviceTableExpirationDateArray[i],
                        self.DeviceTableDayExpirationArray[i],
                        self.DeviceTableMonthExpirationArray[i],
                        self.DeviceTableYearExpirationArray[i],
                        self.DeviceTableSerilaNoArray[i],
                        self.DeviceTableExecutedByArray[i],
                        self.DeviceTableNotesArray[i],
                        self.DeviceTableDeviceModelIDArray[i],
                        self.DeviceTableMACAddressArray[i],
                        self.DeviceTableNATIP4AddressArray[i],
                        self.DeviceTableManagementIP4AddressArray[i],
                        self.DeviceTableLastSuccessICMPArray[i],
                        self.DeviceTableLastICMPStatusArray[i],
                        self.DeviceTableICMPCapableArray[i],
                        self.DeviceTableMonitorviaICMPArray[i],
                        self.DeviceTableMonitorviaICMPArray[i],
                        self.DeviceTableLastUpdatedCMDBDayArray[i],
                        self.DeviceTableLastUpdatedCMDBMonthArray[i],
                        self.DeviceTableLastUpdatedCMDBYearArray[i]
                                            ]
                    self.results.append(self.DeviceTableArrayRowsTemp)

                    self.DeviceTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonDeviceAdd['state']    = ACTIVE
                self.ButtonDeviceEdit['state']   = DISABLED
                self.ButtonDeviceRemove['state'] = DISABLED
                self.ButtonDeviceOK['state']     = DISABLED
                self.ButtonDeviceCancel['state'] = DISABLED
                self.ButtonDeviceCircuits['state'] = DISABLED
                self.ButtonDevicePing64['state'] = DISABLED
                self.ButtonDevicePing1500['state'] = DISABLED
                #self.ButtonDeviceContacts['state'] = DISABLED
                self.ButtonDeviceICMP['state'] = DISABLED
                self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED
                
            else:
                if (self.Selection != 'edit'):
                    mbox.showerror(master=self.DeviceFrame,title='Device',
                    message = 'No Records found')
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Sites_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Sites Refresh']    
            Logging.Log(Parameter)

        self.Clean_Screen('sites','all')
        if self.db.Connect():
            # SQL Querry to the Device Table
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Sites
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' AND Site_ID = '%s'
                      """ % (self.CountryID_Pre,self.RegionID_Pre,self.FacilityID_Pre,self.SitesID_Pre)
            else:
                sql = """
                        SELECT * FROM Sites
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                      """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                             self.FacilityIDArray[self.ComboBoxFacilityID.current()])
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                self.sql_querry = True
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.SitesIDArray = []
                self.SitesNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.SitesIDArray.append(self.db.results[i][4].strip())
                    self.SitesNameArray.append(self.db.results[i][5].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxSitesID['values'] = self.SitesNameArray
                if (len(self.SitesNameArray)== 0):
                    self.ComboBoxSitesID['state'] = DISABLED
                else:
                    self.ComboBoxSitesID['state'] = 'readonly'
                    self.ComboBoxSitesID.set("")
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found')
                self.ComboBoxSitesID['state'] = DISABLED
                self.sql_querry = False
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- SITES SELECTION SECTION <END> ------------------------*
#*************************************************************************************
        
    def on_DeviceWindow_quit(self):
        if (self.DeviceWindowExist):
            self.DeviceWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Device','CLOSE Window']    
                Logging.Log(Parameter)
            self.DeviceWindow.destroy()

    def on_Device_Table_Refresh(self): # I need to do more research on this call.
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Device Refresh']    
            Logging.Log(Parameter)
        self.on_country_combo_changed("event")
       
    def Call_Button_Device_Add(self):
        #-- reset the progess bar --
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','ADD Button']    
            Logging.Log(Parameter)
        self.Enable_Screen('add')
        self.Selection = 'add'

    def Call_Button_Device_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        self.CountryID_Pre = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID_Pre = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID_Pre = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID_Pre = self.SitesIDArray[self.ComboBoxSitesID.current()]
        self.Enable_Screen('edit')
        # ----- Installed Date ---------------------
        self.DeviceInstalledData = {}
        if (self.Selection == 'edit'):
            self.DeviceInstalledDateName = self.DeviceInstalledDateFrameEntry.get()
            curItem = self.DeviceTreeview.focus()  
            dic = self.DeviceTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.DeviceInstalledDateName) > 0):
                if (self.DeviceTableDateInstalledArray[curItem] != 0):
                    self.DeviceInstalledData['day_selected'] = self.DeviceTableDayInstalledArray[curItem]
                    self.DeviceInstalledData['month_selected'] = self.DeviceTableMonthInstalledArray[curItem]
                    self.DeviceInstalledData['year_selected'] = self.DeviceTableYearInstalledArray[curItem]
        # ----- Activated Date ---------------------
        self.DeviceActivatedDateName = self.DeviceActivatedDateFrameEntry.get()
        curItem = self.DeviceTreeview.focus()  
        dic = self.DeviceTreeview.item(curItem)
        curItem = int(dic.get('text')) - 1
        self.DeviceActivatedData = {}
        if (len(self.DeviceActivatedDateName) > 0):
            if (self.DeviceTableDateActivatedArray[curItem] != 0):
                self.DeviceActivatedData['day_selected'] = self.DeviceTableDayActivatedArray[curItem]
                self.DeviceActivatedData['month_selected'] = self.DeviceTableMonthActivatedArray[curItem]
                self.DeviceActivatedData['year_selected'] = self.DeviceTableYearActivatedArray[curItem]
                #print ("Day, Month, Year")
                #print (self.DeviceActivatedData['day_selected'])
                #print (self.DeviceActivatedData['month_selected'])
                #print (self.DeviceActivatedData['year_selected'])
        # ----- Disconnected Date ---------------------
        self.DeviceDisconnectedData = {}
        if (self.Selection == 'edit'):
            self.DeviceDisconnectedDateName = self.DeviceDisconnectedDateFrameEntry.get()
            curItem = self.DeviceTreeview.focus()  
            dic = self.DeviceTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.DeviceDisconnectedDateName) > 0):
                if (self.DeviceTableDisconectedDateArray[curItem] != 0):
                    self.DeviceDisconnectedData['day_selected'] = self.DeviceTableDayDisconectedArray[curItem]
                    self.DeviceDisconnectedData['month_selected'] = self.DeviceTableMonthDisconectedArray[curItem]
                    self.DeviceDisconnectedData['year_selected'] = self.DeviceTableYearDisconectedArray[curItem]
        # ----- Expiration Date ---------------------
        self.DeviceExpirationData = {}
        if (self.Selection == 'edit'):
            self.DeviceExpirationDateName = self.DeviceExpirationDateFrameEntry.get()
            curItem = self.DeviceTreeview.focus()  
            dic = self.DeviceTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.DeviceExpirationDateName) > 0):
                if (self.DeviceTableExpirationDateArray[curItem] != 0):
                    self.DeviceExpirationData['day_selected'] = self.DeviceTableDayExpirationArray[curItem]
                    self.DeviceExpirationData['month_selected'] = self.DeviceTableMonthExpirationArray[curItem]
                    self.DeviceExpirationData['year_selected'] = self.DeviceTableYearExpirationArray[curItem]

    #-------------- Using a Password Question to make sure it was the intent to be deleted <Begin> ---------------
    def Remove_Device_From_DB(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            #self.DeviceID = self.DeviceIDFrameEntry.get()
            PrimaryKey = (self.DeviceID)
            if (mbox.askyesnocancel(master=self.DeviceFrame,title='Device',message = 'Are you Sure you want to Remove it?')):
                #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.DeviceID)
                #print (PrimaryKey)
                sql = """
                        SELECT * FROM Devices
                        WHERE Device_ID = '%s'
                       """ % (PrimaryKey)                
                if (self.db.Execute(sql)):
                    sql = "DELETE FROM Devices WHERE Device_ID = '%s'" % (PrimaryKey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        #self.db.Disconnect()
                        mbox.showwarning(master=self.DeviceFrame,title='Device',
                        message = '*** The Device ID you entered was Removed ***')
                    else:
                        #self.db.Disconnect()
                        mbox.showerror(master=self.DeviceFrame,title='Device',
                        message = '*** The Device ID you entered was NOT Removed ***')               
                    self.on_sites_combo_changed("event")
                    self.Disable_Screen()
                else:
                    mbox.showerror(master=self.DeviceFrame,title='Device',
                    message = '*** The Device ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.GetPasswordWindow.destroy()
            self.Remove_Device_From_DB()
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
            
    #-------------- Using a Password Question to make sure it was the intent to be deleted <End> ---------------

    def Call_Button_Device_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','REMOVE Button']    
            Logging.Log(Parameter)
        self.DeviceID = self.DeviceIDFrameEntry.get()
        self.Get_Usernanme_and_Password()

    def Call_Button_Device_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ButtonDeviceAdd['state']    = DISABLED
            self.ButtonDeviceEdit['state']   = DISABLED
            self.ButtonDeviceRemove['state'] = DISABLED
            self.ButtonDeviceOK['state']     = ACTIVE
            self.ButtonDeviceCancel['state'] = ACTIVE
            self.ButtonDeviceCircuits['state'] = DISABLED
            self.ButtonDevicePing64['state'] = DISABLED
            self.ButtonDevicePing1500['state'] = DISABLED
            #self.ButtonDeviceContacts['state'] = DISABLED
            self.ButtonDeviceICMP['state'] = DISABLED
            self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED
            
            self.Collect_Screen() # <-------------------- Collect all Data on Screen
            PrimaryKey = (self.DeviceID)
            #self.Selection = 'notyet'
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(self.DeviceID) > 0) and (len(self.DeviceDescription) > 0)):
                    if (self.IPFormatCheck(self.DeviceIPAddress) and self.IPFormatCheck(self.DeviceTableNATIP4Address) and
                        self.IPFormatCheck(self.DeviceTableManagementIP4Address)):
                        sql = """
                                SELECT * FROM Devices
                                WHERE Device_ID = '%s'
                              """ % (PrimaryKey)
                        if (self.db.Execute(sql)):
                            mbox.showerror(master=self.DeviceFrame,title='Device',
                            message = '*** The Device ID you entered already exist ***')
                        else:
                            if (len(self.DeviceDescription) > 100):
                                self.DeviceDescription = self.DeviceDescription[:100]
                            sql = "INSERT INTO Devices(Device_ID, Device_Description, Country_ID, Region_ID, Facility_ID, Site_ID, Device_Type_ID, IP4_Address, \
                                       IP6_Address, Contract_No, OutSource_Cost, Maintenance_Cost, Status, Monthly_Cost, \
                                       Date_Installed, Day_Installed, Month_Installed, Year_Installed, \
                                       Date_Activated, Day_Activated, Month_Activated, Year_Activated, \
                                       Disconnect_Date, Day_Disconnect, Month_Disconnect, Year_Disconnect, \
                                       Expiration_Date, Day_Expiration, Month_Expiration, Year_Expiration, Serial_No, Executed_by_UserID, \
                                       Notes, Device_Model_ID, MAC_Address, NAT_IP4_Address, Management_IP4_Address, Last_Success_ICMP, \
                                       Last_ICMP_Status, ICMP_Capable, Monitor_via_ICMP, Last_Updated_CMDB_Date, Last_Updated_CMDB_Day, \
                                       Last_Updated_CMDB_Month, Last_Updated_CMDB_Year) \
                                       VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%f','%f','%s','%f','%s','%d','%d','%d','%s',\
                                       '%d','%d','%d','%s','%d','%d','%d','%s','%d','%d','%d','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s',\
                                       '%s','%s','%d','%d','%d')" %(PrimaryKey,self.DeviceDescription, self.CountryID, self.RegionID, self.FacilityID, self.SitesID,
                                        self.DeviceTypeID, self.DeviceIPAddress, "::", self.DeviceContract, self.DeviceOutSourceCost, self.DeviceMaintenanceCost,
                                        self.DeviceStatus, 0.0, self.DeviceInstalledDate, int(self.DeviceInstalledDay), int(self.DeviceInstalledMonth), int(self.DeviceInstalledYear),
                                        self.DeviceActivatedDate, int(self.DeviceActivatedDay), int(self.DeviceActivatedMonth), int(self.DeviceActivatedYear),
                                        self.DeviceDisconnectedDate, int(self.DeviceDisconnectedDay), int(self.DeviceDisconnectedMonth), int(self.DeviceDisconnectedYear),
                                        self.DeviceExpirationDate, int(self.DeviceExpirationDay), int(self.DeviceExpirationMonth), int(self.DeviceExpirationYear),
                                        self.DeviceSerialNo,self.Username, self.DeviceNotes, self.DeviceModelID, self.DeviceMACAddress, self.DeviceTableNATIP4Address,
                                        self.DeviceTableManagementIP4Address, self.DeviceTableLastSuccessICMP, self.DeviceTableLastICMPStatus,
                                        self.DeviceTableICMPCapable,self.DeviceTableMonitorviaICMP,self.DeviceTableLastUpdatedCMDBDate,0,0,0)
                            #print (sql)
                            '''
                            self.DeviceTableNATIP4Address = self.DeviceNATIPAddressFrameEntry.get()           
                            self.DeviceTableManagementIP4Address = self.DeviceMgmtIPAddressFrameEntry.get()
                            self.DeviceTableICMPCapable = self.DeviceComboBoxICMPCapable.current()
                            self.DeviceTableMonitorviaICMP = self.DeviceComboBoxICMPMonitor.current()
                            self.DeviceTableLastSuccessICMP = self.DeviceLastICMPDateFrameEntry.get()
                            self.DeviceTableLastICMPStatus = self.DeviceLastICMPStatusFrameEntry.get()
                            self.DeviceTableLastUpdatedCMDBDate = self.DeviceLastCMDBStatusFrameEntry.get()
                            # self.DeviceTableLastUpdatedCMDBDay
                            # self.DeviceTableLastUpdatedCMDBMonth
                            # self.DeviceTableLastUpdatedCMDBYear

                            35 NAT_IP4_Address         CHAR(20), '%s'
                            36 Management_IP4_Address  CHAR(20), '%s'
                            37 Last_Success_ICMP       CHAR(50), '%s'
                            38 Last_ICMP_Status        CHAR(50), '%s'
                            39 ICMP_Capable            CHAR(10), '%s'
                            40 Monitor_via_ICMP        CHAR(10), '%s'
                            41 Last_Updated_CMDB_Date  CHAR(50), '%s'
                            42 Last_Updated_CMDB_Day   INT, '%d'
                            43 Last_Updated_CMDB_Month INT, '%d'
                            44 Last_Updated_CMDB_Year  INT)""" '%d'

                            '''
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.DeviceFrame,title='Device',
                                    message = '*** The Device ID you entered was Added ***')
                                self.on_sites_combo_changed("event")
                                self.Disable_Screen()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.DeviceFrame,title='Device',
                                    message = '*** The Device ID you entered was NOT Added ***')
                                self.on_sites_combo_changed("event")
                                self.Disable_Screen()
                    else:
                     mbox.showerror(master=self.DeviceFrame,title='Device',
                        message = '*** The Device IP Address is invalid ***')                         
                else:
                     mbox.showerror(master=self.DeviceFrame,title='Device',
                        message = '*** The Device ID and Device Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                #print (PrimaryKey+"....")
                if (len(self.DeviceDescription) == 0):
                    mbox.showerror(master=self.DeviceFrame,title='Device',
                    message = '*** The Device Name Cannot be BLANK ***')
                else:
                    sql = """
                            SELECT * FROM Devices
                            WHERE Device_ID = '%s'
                          """ % (PrimaryKey)                   
                    if (self.db.Execute(sql)):   
                        #sql = "UPDATE Device SET Device_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID_Device_ID = '%s'" % (DeviceName,self.Username,PrimaryKey)
                        self.DeviceIPv6 = "::"
                        if (len(self.DeviceDescription) > 100):
                            self.DeviceDescription = self.DeviceDescription[:100]
                        #sql = "UPDATE Device SET Site_Name = '%s', Description = '%s', Executed_by_UserID = '%s' \
                        #       WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'" %(self.DeviceName,self.DeviceDescription,self.Username,PrimaryKey)

                        sql = "UPDATE Devices SET Device_Description = '%s', Country_ID = '%s', Region_ID = '%s', Facility_ID = '%s', \
                                       Site_ID = '%s', Device_Type_ID = '%s', IP4_Address = '%s', \
                                       IP6_Address = '%s', Contract_No = '%s', OutSource_Cost = '%f', Maintenance_Cost = '%f', Status = '%s', Monthly_Cost = '%f', \
                                       Date_Installed = '%s', Day_Installed = '%d', Month_Installed = '%d', Year_Installed = '%d', \
                                       Date_Activated = '%s', Day_Activated = '%d', Month_Activated = '%d', Year_Activated = '%d', \
                                       Disconnect_Date = '%s', Day_Disconnect = '%d', Month_Disconnect = '%d', Year_Disconnect = '%d', \
                                       Expiration_Date = '%s', Day_Expiration = '%d', Month_Expiration = '%d', Year_Expiration = '%d', Serial_No = '%s', Executed_by_UserID = '%s', \
                                       Notes = '%s', Device_Model_ID = '%s', MAC_Address = '%s', NAT_IP4_Address = '%s', Management_IP4_Address = '%s', \
                                       ICMP_Capable = '%s', Monitor_via_ICMP = '%s' \
                                WHERE Device_ID  = '%s'" %(self.DeviceDescription, self.CountryID, self.RegionID, self.FacilityID, self.SitesID,
                                        self.DeviceTypeID, self.DeviceIPAddress, "::", self.DeviceContract, self.DeviceOutSourceCost, self.DeviceMaintenanceCost,
                                        self.DeviceStatus, 0.0, self.DeviceInstalledDate, int(self.DeviceInstalledDay), int(self.DeviceInstalledMonth), int(self.DeviceInstalledYear),
                                        self.DeviceActivatedDate, int(self.DeviceActivatedDay), int(self.DeviceActivatedMonth), int(self.DeviceActivatedYear),
                                        self.DeviceDisconnectedDate, int(self.DeviceDisconnectedDay), int(self.DeviceDisconnectedMonth), int(self.DeviceDisconnectedYear),
                                        self.DeviceExpirationDate, int(self.DeviceExpirationDay), int(self.DeviceExpirationMonth), int(self.DeviceExpirationYear),
                                        self.DeviceSerialNo,self.Username, self.DeviceNotes, self.DeviceModelID, self.DeviceMACAddress, self.DeviceTableNATIP4Address,
                                        self.DeviceTableManagementIP4Address,self.DeviceTableICMPCapable,self.DeviceTableMonitorviaICMP,PrimaryKey)

                        '''
                            self.DeviceTableNATIP4Address = self.DeviceNATIPAddressFrameEntry.get()           
                            self.DeviceTableManagementIP4Address = self.DeviceMgmtIPAddressFrameEntry.get()
                            self.DeviceTableICMPCapable = self.DeviceComboBoxICMPCapable.current()
                            self.DeviceTableMonitorviaICMP = self.DeviceComboBoxICMPMonitor.current()
                            self.DeviceTableLastSuccessICMP = self.DeviceLastICMPDateFrameEntry.get()
                            self.DeviceTableLastICMPStatus = self.DeviceLastICMPStatusFrameEntry.get()
                            self.DeviceTableLastUpdatedCMDBDate = self.DeviceLastCMDBStatusFrameEntry.get()
                            # self.DeviceTableLastUpdatedCMDBDay
                            # self.DeviceTableLastUpdatedCMDBMonth
                            # self.DeviceTableLastUpdatedCMDBYear

                            35 NAT_IP4_Address         CHAR(20), '%s'
                            36 Management_IP4_Address  CHAR(20), '%s'
                            37 Last_Success_ICMP       CHAR(50), '%s'
                            38 Last_ICMP_Status        CHAR(50), '%s'
                            39 ICMP_Capable            CHAR(10), '%s'
                            40 Monitor_via_ICMP        CHAR(10), '%s'
                            41 Last_Updated_CMDB_Date  CHAR(50), '%s'
                            42 Last_Updated_CMDB_Day   INT, '%d'
                            43 Last_Updated_CMDB_Month INT, '%d'
                            44 Last_Updated_CMDB_Year  INT)""" '%d'

                        '''

                        
                        '''
                            self.CountryID
                            self.RegionID
                            self.FacilityID
                            self.SitesID
                            self.DeviceID
                            self.DeviceDescription
                            self.DeviceInstalledDate, self.DeviceInstalledDay, self.DeviceInstalledMonth, self.DeviceInstalledYear
                            self.DeviceActivatedDate, self.DeviceActivatedMonth, self.DeviceActivatedDay, self.DeviceActivatedYear
                            self.DeviceDisconnectedDate, self.DeviceDisconnectedDay, self.DeviceDisconnectedMonth, self.DeviceDisconnectedYear
                            self.DeviceExpirationDate, self.DeviceExpirationDay, self.DeviceExpirationMonth, self.DeviceExpirationYear
                            self.DeviceTypeID 
                            self.DeviceModelID
                            self.DeviceStatus 
                            self.DeviceIPAddress 
                            self.DeviceContract
                            self.DeviceSerialNo 
                            self.DeviceMACAddress
                            self.DeviceOutSourceCost 
                            self.DeviceMaintenanceCost 
                            self.DeviceNotes 

                            Device_ID               CHAR(50) NOT NULL PRIMARY KEY,
                            Device_Description      CHAR(100) NOT NULL,
                            Country_ID              CHAR(20) NOT NULL,                    
                            Region_ID               CHAR(20) NOT NULL,
                            Facility_ID             CHAR(20) NOT NULL,
                            Site_ID                 CHAR(20) NOT NULL,
                            Device_Type_ID          CHAR(30) NOT NULL,
                            IP4_Address             CHAR(20),
                            IP6_Address             CHAR(200),
                            Contract_No             CHAR(20),
                            OutSource_Cost          FLOAT,
                            Maintenance_Cost        FLOAT,
                            Status                  CHAR(20),
                            Monthly_Cost            FLOAT,
                            Date_Installed          CHAR(20),
                            Day_Installed           INT,
                            Month_Installed         INT,
                            Year_Installed          INT,
                            Date_Activated          CHAR(20),
                            Day_Activated           INT,
                            Month_Activated         INT,
                            Year_Activated          INT,
                            Disconnect_Date         CHAR(20),
                            Day_Disconnect          INT,
                            Month_Disconnect        INT,
                            Year_Disconnect         INT,
                            Expiration_Date         CHAR(20),
                            Day_Expiration          INT,
                            Month_Expiration        INT,
                            Year_Expiration         INT,
                            Serial_No               CHAR(100),
                            Executed_by_UserID      CHAR(20),
                            Notes                   CHAR(200),
                            Device_Model_ID         CHAR(30),
                            MAC_Address             CHAR(20))"""

                    NAT_IP4_Address         CHAR(20),    
                    Management_IP4_Address  CHAR(20),
                    Last_Success_ICMP       CHAR(50),
                    Last_ICMP_Status        CHAR(50),
                    ICMP_Capable            CHAR(10),
                    Monitor_via_ICMP        CHAR(10),
                    Last_Updated_CMDB_Date  CHAR(50),
                    Last_Updated_CMDB_Day   INT,
                    Last_Updated_CMDB_Month INT,
                    Last_Updated_CMDB_Year  INT)"""

                            
                        '''
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.DeviceFrame,title='Device',
                                message = '*** The Device ID you entered was Updated ***')
                            if ((self.CountryID_Pre != self.CountryIDArray[self.ComboBoxCoutryID.current()]) or
                                (self.RegionID_Pre != self.RegionIDArray[self.ComboBoxRegionID.current()]) or
                                (self.FacilityID_Pre != self.FacilityIDArray[self.ComboBoxFacilityID.current()]) or
                                (self.SitesID_Pre != self.SitesIDArray[self.ComboBoxSitesID.current()])):
                                # ------ The Location Change so we need to move to a new location !!!!!
                                self.CountryID_Pre = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                                self.RegionID_Pre = self.RegionIDArray[self.ComboBoxRegionID.current()]
                                self.FacilityID_Pre = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
                                self.SitesID_Pre = self.SitesIDArray[self.ComboBoxSitesID.current()]
                                self.Selection = 'cancel_edit'
                                self.on_Country_Table_Refresh()
                                self.ComboBoxCoutryID.current(0)            
                                self.on_Region_Table_Refresh()
                                self.ComboBoxRegionID.current(0)
                                self.on_Facility_Table_Refresh()
                                self.ComboBoxFacilityID.current(0)
                                self.on_Sites_Table_Refresh()
                                self.ComboBoxSitesID.current(0)
                                self.on_sites_combo_changed("event")
                            self.Selection = 'edit_ok'
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.DeviceFrame,title='Device',
                                message = '*** The Device ID you entered was NOT Upadted ***')
                        self.on_sites_combo_changed("event")
                        self.Disable_Screen()
                    else:
                        mbox.showerror(master=self.DeviceFrame,title='Device',
                        message = '*** The Device ID you try to Edit Does not exist Anymore ***')
                        self.on_sites_combo_changed("event")
                        self.Disable_Screen()                        
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Device_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','CANCEL Button']    
            Logging.Log(Parameter)
        if (self.Selection == 'edit'):
            if ((self.CountryID_Pre != self.CountryIDArray[self.ComboBoxCoutryID.current()]) or
                (self.RegionID_Pre != self.RegionIDArray[self.ComboBoxRegionID.current()]) or
                (self.FacilityID_Pre != self.FacilityIDArray[self.ComboBoxFacilityID.current()]) or
                (self.SitesID_Pre != self.SitesIDArray[self.ComboBoxSitesID.current()])):
                # ------ The Location Change so we need to move to a new location !!!!!
                self.CountryID_Pre = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                self.RegionID_Pre = self.RegionIDArray[self.ComboBoxRegionID.current()]
                self.FacilityID_Pre = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
                self.SitesID_Pre = self.SitesIDArray[self.ComboBoxSitesID.current()]
                self.Selection = 'cancel_edit'
                self.on_Country_Table_Refresh()
                self.ComboBoxCoutryID.current(0)            
                self.on_Region_Table_Refresh()
                self.ComboBoxRegionID.current(0)
                self.on_Facility_Table_Refresh()
                self.ComboBoxFacilityID.current(0)
                self.on_Sites_Table_Refresh()
                self.ComboBoxSitesID.current(0)
                self.on_sites_combo_changed("event")
        self.Selection = 'cancel'
        self.Clean_Screen('Device','all')
        self.on_sites_combo_changed("test")
        self.ComboBoxCoutryID['state'] = 'readonly'
        self.ComboBoxRegionID['state'] = 'readonly'
        self.ComboBoxFacilityID['state'] = 'readonly'
        self.ComboBoxSitesID['state'] = 'readonly'
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = ACTIVE
            self.ButtonCountryRefresh['state'] = ACTIVE
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = ACTIVE
            self.ButtonRegionRefresh['state'] = ACTIVE
        if (Is_Facility_Available):
            self.ButtonFacilityAdd['state'] = ACTIVE
            self.ButtonFacilityRefresh['state'] = ACTIVE
        if (Is_Sites_Available):
            self.ButtonSitesAdd['state'] = ACTIVE
            self.ButtonSitesRefresh['state'] = ACTIVE
        #self.Selection = 'cancel'


    def On_Run_Ping(self,ip,size):
        #print ("PING.....")
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','PING Window']    
            Logging.Log(Parameter)
        if (len(ip) > 6):
            '''
            import os
            if os.system("ping -c 1 google.com") == 0:
            ...     print "host appears to be up"
            '''
            cmd = ("ping " + ip + " -l "+size+" -w 2 -t")
            os.system(cmd)

    def On_Ping(self,ip,size):
        try:
            self.thread = Thread(target=self.On_Run_Ping, args=(ip,size,))
            self.thread.daemon=True
            self.thread.start()
        except(KeyboardInterrupt, SystemExit):
            sys.exit("Interrupted by ctrl+c\n")

            
    def Call_Button_Ping64(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','PING 64 Button']    
            Logging.Log(Parameter)
        ip = self.DeviceIPAddressFrameEntry.get()
        self.On_Ping(ip,'64')

    def Call_Button_Ping1500(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','PING 1500 Button']    
            Logging.Log(Parameter)
        ip = self.DeviceIPAddressFrameEntry.get()
        self.On_Ping(ip,'1500')

    def Call_Button_Circuits(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Circuits Button']    
            Logging.Log(Parameter)
        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]        
        Location = []
        Location = [self.CountryID, self.RegionID, self.FacilityID, self.SitesID]
        Circuit = Class_Circuits(ODBC_DSN_name,Windows_Scaling,Location)
        Circuit.Display_Circuits_Window()

    def Call_Button_ICMP(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','ICMP Button']    
            Logging.Log(Parameter)
        Location = []
        ip = self.DeviceIPAddressFrameEntry.get()
        Location = ['IP Address',ip]
        ICMP = Class_ICMP(ODBC_DSN_name,Windows_Scaling,Location)
        ICMP.Display_ICMP_Window()

    def Call_Button_LocalPointOfContacts(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Local POC Button']    
            Logging.Log(Parameter)
        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]        
        Location = []
        Location = [self.CountryID, self.RegionID, self.FacilityID, self.SitesID]
        LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,Location)
        LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def On_Run_Putty(self,ip):
        #print ("PING.....")
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Putty Run']    
            Logging.Log(Parameter)
        if (len(ip) > 6):
            cmd = ("putty.exe -ssh " + ip)
            os.system(cmd)

    def On_Putty(self,ip):
        try:
            self.thread = Thread(target=self.On_Run_Putty, args=(ip,))
            self.thread.daemon=True
            self.thread.start()
        except(KeyboardInterrupt, SystemExit):
            sys.exit("Interrupted by ctrl+c\n")

            
    def Get_IPAddress(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Putty IP Address']    
            Logging.Log(Parameter)
        ip = self.DeviceIPAddressFrameEntry.get()
        self.On_Putty(ip)

    def Get_NAT_IPAddress(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Putty IP NAT Address']    
            Logging.Log(Parameter)
        ip = self.DeviceNATIPAddressFrameEntry.get()
        self.On_Putty(ip)

    def Get_Mgmt_IPAddress(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Device','Putty IP Management Address']    
            Logging.Log(Parameter)
        ip = self.DeviceMgmtIPAddressFrameEntry.get()
        self.On_Putty(ip)


    def Call_Button_Contacts(self):
        mbox.showerror(master=self.DeviceFrame,title='Under Construction',
                         message = '... Comming Soon...')               
        

    def on_Device_Tree_select_click(self,event):
        #print ("Select")
        if (self.Selection == 'edit'):
            if ((self.CountryID_Pre != self.CountryIDArray[self.ComboBoxCoutryID.current()]) or
                (self.RegionID_Pre != self.RegionIDArray[self.ComboBoxRegionID.current()]) or
                (self.FacilityID_Pre != self.FacilityIDArray[self.ComboBoxFacilityID.current()]) or
                (self.SitesID_Pre != self.SitesIDArray[self.ComboBoxSitesID.current()])):
                # ------ The Location Change so we need to move to a new location !!!!!
                self.CountryID_Pre = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                self.RegionID_Pre = self.RegionIDArray[self.ComboBoxRegionID.current()]
                self.FacilityID_Pre = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
                self.SitesID_Pre = self.SitesIDArray[self.ComboBoxSitesID.current()]
                self.Selection = 'cancel_edit'
                self.on_Country_Table_Refresh()
                self.ComboBoxCoutryID.current(0)            
                self.on_Region_Table_Refresh()
                self.ComboBoxRegionID.current(0)
                self.on_Facility_Table_Refresh()
                self.ComboBoxFacilityID.current(0)
                self.on_Sites_Table_Refresh()
                self.ComboBoxSitesID.current(0)
                self.on_sites_combo_changed("event")
        self.Selection = 'select'
        curItem = self.DeviceTreeview.focus()  
        dic = self.DeviceTreeview.item(curItem)
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
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = ACTIVE
                self.ButtonSitesRefresh['state'] = ACTIVE            
            curItem = int(dic.get('text')) - 1
            self.ButtonDeviceAdd['state']    = ACTIVE
            self.ButtonDeviceEdit['state']   = ACTIVE
            self.ButtonDeviceRemove['state'] = ACTIVE
            self.ButtonDeviceOK['state']     = DISABLED
            self.ButtonDeviceCancel['state'] = DISABLED
            self.ButtonDeviceCircuits['state'] = ACTIVE
            self.ButtonDevicePing64['state'] = ACTIVE
            self.ButtonDevicePing1500['state'] = ACTIVE
            #self.ButtonDeviceContacts['state'] = ACTIVE
            self.ButtonDeviceICMP['state'] = ACTIVE
            self.ButtonDeviceLocalPointOfContacts['state'] = ACTIVE
            self.Display_Screen(curItem)

    def On_Device_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_Device_Table_Refresh()

    #--------------------- Installed Date Calendar  <BEGIN> ------------------------------
    def on_DeviceInstalledDateWindow_quit(self):
        self.DeviceInstalledData = self.cal_Installed.kill_and_save()
        self.DeviceCalendarInstalledDateExist = False
        self.DeviceCalendarInstalledDateWindow.destroy()
        if (len(self.DeviceInstalledData) > 0):
            self.DeviceInstalledDate = str(self.DeviceInstalledData['month_selected']) + '/' + str(self.DeviceInstalledData['day_selected']) + '/' + str(self.DeviceInstalledData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.DeviceInstalledDate = ""
        self.DeviceInstalledDateFrameEntry['state'] = 'normal'
        self.DeviceInstalledDateFrameEntry.delete(0,END)
        self.DeviceInstalledDateFrameEntry.insert(0,self.DeviceInstalledDate)
        self.DeviceInstalledDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Installed_Date(self):
        if not (self.DeviceCalendarInstalledDateExist):
            self.DeviceCalendarInstalledDateExist = True
            self.DeviceCalendarInstalledDateWindow = Tk()
            self.DeviceCalendarInstalledDateWindow.title("Installed")
            self.DeviceCalendarInstalledDateWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceInstalledDateWindow_quit)
            self.DeviceCalendarInstalledDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.DeviceInstalledData = {}
            if (self.Selection == 'edit'):
                self.DeviceInstalledDateName = self.DeviceInstalledDateFrameEntry.get()
                curItem = self.DeviceTreeview.focus()  
                dic = self.DeviceTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.DeviceInstalledDateName) > 0):
                    if (self.DeviceTableDateInstalledArray[curItem] != 0):
                        self.DeviceInstalledData['day_selected'] = self.DeviceTableDayInstalledArray[curItem]
                        self.DeviceInstalledData['month_selected'] = self.DeviceTableMonthInstalledArray[curItem]
                        self.DeviceInstalledData['year_selected'] = self.DeviceTableYearInstalledArray[curItem]
            self.cal_Installed = Class_Calendar(self.DeviceCalendarInstalledDateWindow, self.DeviceInstalledData)
            self.cal_Installed.setup()
    #--------------------- Installed Date Calendar  <END> ------------------------------

    #--------------------- Activated Date Calendar  <BEGIN> ------------------------------
    def on_DeviceActivatedDateWindow_quit(self):
        self.DeviceActivatedData = self.cal_Activated.kill_and_save()
        self.DeviceCalendarActivatedDateExist = False
        self.DeviceCalendarActivatedDateWindow.destroy()
        if (len(self.DeviceActivatedData) > 0):
            self.DeviceActivatedDate = str(self.DeviceActivatedData['month_selected']) + '/' + str(self.DeviceActivatedData['day_selected']) + '/' + str(self.DeviceActivatedData['year_selected'])
            #print ("=> Day, Month, Year")
            #print (self.DeviceActivatedData['day_selected'])
            #print (self.DeviceActivatedData['month_selected'])
            #print (self.DeviceActivatedData['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.DeviceActivatedDate = ""
        self.DeviceActivatedDateFrameEntry['state'] = 'normal'
        self.DeviceActivatedDateFrameEntry.delete(0,END)
        self.DeviceActivatedDateFrameEntry.insert(0,self.DeviceActivatedDate)
        self.DeviceActivatedDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Activated_Date(self):
        if not (self.DeviceCalendarActivatedDateExist):
            self.DeviceCalendarActivatedDateExist = True
            self.DeviceCalendarActivatedDateWindow = Tk()
            self.DeviceCalendarActivatedDateWindow.title("Activated")
            self.DeviceCalendarActivatedDateWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceActivatedDateWindow_quit)
            self.DeviceCalendarActivatedDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.DeviceActivatedData = {}
            if (self.Selection == 'edit'):
                self.DeviceActivatedDateName = self.DeviceActivatedDateFrameEntry.get()
                curItem = self.DeviceTreeview.focus()  
                dic = self.DeviceTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.DeviceActivatedDateName) > 0):
                    if (self.DeviceTableDateActivatedArray[curItem] != 0):
                        self.DeviceActivatedData['day_selected'] = self.DeviceTableDayActivatedArray[curItem]
                        self.DeviceActivatedData['month_selected'] = self.DeviceTableMonthActivatedArray[curItem]
                        self.DeviceActivatedData['year_selected'] = self.DeviceTableYearActivatedArray[curItem]
                        #print ("Day, Month, Year")
                        #print (self.DeviceActivatedData['day_selected'])
                        #print (self.DeviceActivatedData['month_selected'])
                        #print (self.DeviceActivatedData['year_selected'])
 
            self.cal_Activated = Class_Calendar(self.DeviceCalendarActivatedDateWindow, self.DeviceActivatedData)
            self.cal_Activated.setup()
    #--------------------- Activated Date Calendar  <END> ------------------------------

    #--------------------- Disconnected Date Calendar  <BEGIN> ------------------------------
    def on_DeviceDisconnectedDateWindow_quit(self):
        self.DeviceDisconnectedData = self.cal_Disconnected.kill_and_save()
        self.DeviceCalendarDisconnectedDateExist = False
        self.DeviceCalendarDisconnectedDateWindow.destroy()
        if (len(self.DeviceDisconnectedData) > 0):
            self.DeviceDisconnectedDate = str(self.DeviceDisconnectedData['month_selected']) + '/' + str(self.DeviceDisconnectedData['day_selected']) + '/' + str(self.DeviceDisconnectedData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.DeviceDisconnectedDate = ""
        self.DeviceDisconnectedDateFrameEntry['state'] = 'normal'
        self.DeviceDisconnectedDateFrameEntry.delete(0,END)
        self.DeviceDisconnectedDateFrameEntry.insert(0,self.DeviceDisconnectedDate)
        self.DeviceDisconnectedDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Disconnected_Date(self):
        if not (self.DeviceCalendarDisconnectedDateExist):
            self.DeviceCalendarDisconnectedDateExist = True
            self.DeviceCalendarDisconnectedDateWindow = Tk()
            self.DeviceCalendarDisconnectedDateWindow.title("Disconnected")
            self.DeviceCalendarDisconnectedDateWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceDisconnectedDateWindow_quit)
            self.DeviceCalendarDisconnectedDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.DeviceDisconnectedData = {}
            if (self.Selection == 'edit'):
                self.DeviceDisconnectedDateName = self.DeviceDisconnectedDateFrameEntry.get()
                curItem = self.DeviceTreeview.focus()  
                dic = self.DeviceTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.DeviceDisconnectedDateName) > 0):
                    if (self.DeviceTableDisconectedDateArray[curItem] != 0):
                        self.DeviceDisconnectedData['day_selected'] = self.DeviceTableDayDisconectedArray[curItem]
                        self.DeviceDisconnectedData['month_selected'] = self.DeviceTableMonthDisconectedArray[curItem]
                        self.DeviceDisconnectedData['year_selected'] = self.DeviceTableYearDisconectedArray[curItem]
            self.cal_Disconnected = Class_Calendar(self.DeviceCalendarDisconnectedDateWindow, self.DeviceDisconnectedData)
            self.cal_Disconnected.setup()
    #--------------------- Disconnected Date Calendar  <END> ------------------------------

    #--------------------- Expiration Date Calendar  <BEGIN> ------------------------------
    def on_DeviceExpirationDateWindow_quit(self):
        self.DeviceExpirationData = self.cal_Expiration.kill_and_save()
        self.DeviceCalendarExpirationDateExist = False
        self.DeviceCalendarExpirationDateWindow.destroy()
        if (len(self.DeviceExpirationData) > 0):
            self.DeviceExpirationDate = str(self.DeviceExpirationData['month_selected']) + '/' + str(self.DeviceExpirationData['day_selected']) + '/' + str(self.DeviceExpirationData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.DeviceExpirationDate = ""
        self.DeviceExpirationDateFrameEntry['state'] = 'normal'
        self.DeviceExpirationDateFrameEntry.delete(0,END)
        self.DeviceExpirationDateFrameEntry.insert(0,self.DeviceExpirationDate)
        self.DeviceExpirationDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Expiration_Date(self):
        if not (self.DeviceCalendarExpirationDateExist):
            self.DeviceCalendarExpirationDateExist = True
            self.DeviceCalendarExpirationDateWindow = Tk()
            self.DeviceCalendarExpirationDateWindow.title("Expiration")
            self.DeviceCalendarExpirationDateWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceExpirationDateWindow_quit)
            self.DeviceCalendarExpirationDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.DeviceExpirationData = {}
            if (self.Selection == 'edit'):
                self.DeviceExpirationDateName = self.DeviceExpirationDateFrameEntry.get()
                curItem = self.DeviceTreeview.focus()  
                dic = self.DeviceTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.DeviceExpirationDateName) > 0):
                    if (self.DeviceTableExpirationDateArray[curItem] != 0):
                        self.DeviceExpirationData['day_selected'] = self.DeviceTableDayExpirationArray[curItem]
                        self.DeviceExpirationData['month_selected'] = self.DeviceTableMonthExpirationArray[curItem]
                        self.DeviceExpirationData['year_selected'] = self.DeviceTableYearExpirationArray[curItem]
            self.cal_Expiration = Class_Calendar(self.DeviceCalendarExpirationDateWindow, self.DeviceExpirationData)
            self.cal_Expiration.setup()
    #--------------------- Expiration Date Calendar  <END> ------------------------------


    def on_device_type_combo_changed(self,event):
        #print (".")
        if self.db.Connect():
            # SQL Querry to the Device Table
            #self.DeviceTypeVendorIDArray
            #self.DeviceComboBoxTypeID
            sql = """
                    SELECT * FROM Device_Model
                    WHERE Vendor_ID = '%s' AND Device_Type_ID = '%s'
                    ORDER BY Device_Model_Name ASC
                  """ % (self.DeviceTypeVendorIDArray[self.DeviceComboBoxTypeID.current()],
                         self.DeviceTypeIDArray[self.DeviceComboBoxTypeID.current()])
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.DeviceModelIDArray = []
                self.DeviceModelNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceModelIDArray.append(self.db.results[i][0].strip())
                    self.DeviceModelNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.DeviceComboBoxModelID['values'] = self.DeviceModelNameArray
                if (len(self.DeviceModelNameArray)== 0):
                    self.DeviceComboBoxModelID['state'] = DISABLED ######
                    self.Is_Get_Type_and_Model = False
                else:
                    self.DeviceComboBoxModelID['state'] = 'readonly'
                    self.DeviceComboBoxModelID.current(0)
                    #self.DeviceComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                sql = """
                        SELECT * FROM Device_Model
                        WHERE Vendor_ID = '%s'
                      """ % ('UNKNOWN')
                #print (sql)
                if (self.db.Execute(sql)):
                    #print ("found it")
                    i = 0
                    self.progress['maximum'] = len(self.db.results)
                    self.DeviceModelIDArray = []
                    self.DeviceModelNameArray = []
                    while (i < len(self.db.results)):
                        num = i + 1
                        self.DeviceModelIDArray.append(self.db.results[i][0].strip())
                        self.DeviceModelNameArray.append(self.db.results[i][1].strip())
                        i = i + 1
                        self.progress['value'] = i
                    self.DeviceComboBoxModelID['values'] = self.DeviceModelNameArray
                    if (len(self.DeviceModelNameArray)== 0):
                        self.DeviceComboBoxModelID['state'] = DISABLED ########
                        self.Is_Get_Type_and_Model = False
                    else:
                        self.DeviceComboBoxModelID['state'] = 'readonly'
                        self.DeviceComboBoxModelID.current(0)
                        #self.DeviceComboBoxTypeID.set("")
                        self.Is_Get_Type_and_Model = True
                else:
                    mbox.showerror(master=self.DeviceFrame,title='Device',
                    message = 'No Records found for Device Model')                
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def Get_Type_Model_and_Satus(self):
        #print ("")
        # self.DeviceComboBoxTypeID
        # self.DeviceComboBoxModelID
        # self.DeviceComboBoxStatus
        self.Is_Get_Type_and_Model = False
        if self.db.Connect():
            # SQL Querry to the Device Type
            sql = """
                    SELECT * FROM Device_Type
                    ORDER BY Device_Type_Name ASC
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.DeviceTypeIDArray = []
                self.DeviceTypeNameArray = []
                self.DeviceTypeVendorIDArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceTypeIDArray.append(self.db.results[i][0].strip())
                    self.DeviceTypeNameArray.append(self.db.results[i][1].strip())
                    self.DeviceTypeVendorIDArray.append(self.db.results[i][2].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.DeviceComboBoxTypeID['values'] = self.DeviceTypeNameArray
                if (len(self.DeviceTypeNameArray)== 0):
                    self.DeviceComboBoxTypeID['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.DeviceComboBoxTypeID['state'] = DISABLED #'readonly'
                    self.DeviceComboBoxTypeID.current(0)
                    #self.DeviceComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found for Device Type')

            # SQL Querry to the Device Model
            sql = """
                    SELECT * FROM Device_Model
                    ORDER BY Device_Model_ID ASC
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.DeviceModelIDArray = []
                self.DeviceModelNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.DeviceModelIDArray.append(self.db.results[i][0].strip())
                    self.DeviceModelNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.DeviceComboBoxModelID['values'] = self.DeviceModelNameArray
                if (len(self.DeviceModelNameArray)== 0):
                    self.DeviceComboBoxModelID['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.DeviceComboBoxModelID['state'] = DISABLED #'readonly'
                    self.DeviceComboBoxModelID.current(0)
                    #self.DeviceComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                message = 'No Records found for Device Model')
        else:
            mbox.showerror(master=self.DeviceFrame,title='Device',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        # ---- Status ----
        self.DeviceStatusValues = ['production','implementation','pre-production','unmanaged','obsolete']
        self.DeviceComboBoxStatus['values'] = self.DeviceStatusValues
        self.DeviceComboBoxStatus['state'] = DISABLED
        self.DeviceComboBoxStatus.current(0)

        self.DeviceICMPValues = ['NO','YES']
        self.DeviceComboBoxICMPCapable['values'] = self.DeviceICMPValues
        self.DeviceComboBoxICMPCapable['state'] = DISABLED
        self.DeviceComboBoxICMPCapable.current(0)

        self.DeviceComboBoxICMPMonitor['values'] = self.DeviceICMPValues
        self.DeviceComboBoxICMPMonitor['state'] = DISABLED
        self.DeviceComboBoxICMPMonitor.current(0)
        
    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Device List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.DeviceTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.DeviceFrame,title='Device List',
                            message = '!!! The File was saved !!!')
                    self.DeviceWindow.title("Device                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Device','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.DeviceFrame,title='Device List',
                            message = '*** The File was not saved, Perhaps It is already open ***')



    def Display_Device_Window(self): 
        if not self.DeviceWindowExist:
            # Set up the Window
            self.DeviceWindowExist = True
            self.DeviceWindow = Tk()
            self.DeviceWindow.geometry('1350x700+350+70')
            self.DeviceWindow.title("Device")
            self.DeviceWindow.protocol("WM_DELETE_WINDOW", self.on_DeviceWindow_quit)
            self.DeviceWindow.call('tk', 'scaling', Windows_Scaling)
            DeviceLabel = Label(self.DeviceWindow,text="Helvetica", font=("Helvetica", 19))
            DeviceLabel["text"] = "Device / Equipment"
            DeviceLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.DeviceWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_DeviceWindow_quit)
            self.DeviceWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------

                    
            # Setup Frame
            self.DeviceFrame = Frame(self.DeviceWindow)
            self.DeviceFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.DeviceFrame.rowconfigure(0, weight=1)
            self.DeviceFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.DeviceFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 8, width=75, height=24)

                        self.ButtonCountryRefresh = Button(self.DeviceFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 8, width=75, height=24)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.DeviceFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 38, width=75, height=24)

                        self.ButtonRegionRefresh = Button(self.DeviceFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 38, width=75, height=24)

                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd = Button(self.DeviceFrame, text = '+ Facility', command = self.Display_Facility_Window, state=DISABLED)
                        self.ButtonFacilityAdd.place(x = 450, y = 68, width=75, height=24)

                        self.ButtonFacilityRefresh = Button(self.DeviceFrame, text = 'Refresh', command = self.on_Facility_Table_Refresh, state=DISABLED)
                        self.ButtonFacilityRefresh.place(x = 550, y = 68, width=75, height=24)

                    if (Is_Sites_Available):
                        self.ButtonSitesAdd = Button(self.DeviceFrame, text = '+ Site', command = self.Display_Sites_Window, state=DISABLED)
                        self.ButtonSitesAdd.place(x = 450, y = 98, width=75, height=24)

                        self.ButtonSitesRefresh = Button(self.DeviceFrame, text = 'Refresh', command = self.on_Sites_Table_Refresh, state=DISABLED)
                        self.ButtonSitesRefresh.place(x = 550, y = 98, width=75, height=24)


                    self.ButtonDeviceAdd = Button(self.DeviceFrame, text = 'Add', command = self.Call_Button_Device_Add, state=DISABLED)
                    self.ButtonDeviceAdd.place(x = 450, y = 128, width=75, height=25)

                    self.ButtonDeviceEdit = Button(self.DeviceFrame, text = 'Edit', command = self.Call_Button_Device_Edit, state=DISABLED)
                    self.ButtonDeviceEdit.place(x = 550, y = 128, width=75, height=25)

                    self.ButtonDeviceRemove = Button(self.DeviceFrame, text = 'Remove', command = self.Call_Button_Device_Remove, state=DISABLED)
                    self.ButtonDeviceRemove.place(x = 650, y = 128, width=75, height=25)

                    self.ButtonDeviceOK = Button(self.DeviceFrame, text = 'OK / UPDATE', command = self.Call_Button_Device_OK, state=DISABLED)
                    self.ButtonDeviceOK.place(x = 750, y = 128, width=100, height=25)

                    self.ButtonDeviceCancel = Button(self.DeviceFrame, text = 'Cancel', command = self.Call_Button_Device_Cancel, state=DISABLED)
                    self.ButtonDeviceCancel.place(x = 875, y = 128, width=75, height=25)

                    # Utilities Buttons
                    self.ButtonDeviceICMP = Button(self.DeviceFrame, text = 'ICMP', command = self.Call_Button_ICMP, state=DISABLED)
                    self.ButtonDeviceICMP.place(x = 650, y = 8, width=75, height=25)

                    self.ButtonDeviceCircuits = Button(self.DeviceFrame, text = 'Circuits', command = self.Call_Button_Circuits, state=DISABLED)
                    self.ButtonDeviceCircuits.place(x = 750, y = 8, width=100, height=25)

                    self.ButtonDeviceLocalPointOfContacts = Button(self.DeviceFrame, text = 'Local POC', command = self.Call_Button_LocalPointOfContacts, state=DISABLED)
                    self.ButtonDeviceLocalPointOfContacts.place(x = 650, y = 38, width=75, height=25)


                    self.ButtonDevicePing64 = Button(self.DeviceFrame, text = 'PING 64 Bytes', command = self.Call_Button_Ping64, state=DISABLED)
                    self.ButtonDevicePing64.place(x = 750, y = 38, width=100, height=25)

                    self.ButtonDevicePing1500 = Button(self.DeviceFrame, text = 'PING 1500 Bytes', command = self.Call_Button_Ping1500, state=DISABLED)
                    self.ButtonDevicePing1500.place(x = 750, y = 68, width=100, height=25)

                    #self.ButtonDeviceContacts = Button(self.DeviceFrame, text = 'Local Contacts', command = self.Call_Button_Contacts, state=DISABLED)
                    #self.ButtonDeviceContacts.place(x = 750, y = 98, width=100, height=25)


                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.DeviceFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=158)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=40)

                    FacilityIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityIDFrameLabel["text"] = "Facility Name:"
                    FacilityIDFrameLabel.place(x=10, y=70)

                    SitesIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    SitesIDFrameLabel["text"] = "Site Name:"
                    SitesIDFrameLabel.place(x=10, y=100)

                    # Setup Labels and Entry
                    self.DeviceIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceIDFrameLabel["text"] = "Device Name:"
                    self.DeviceIDFrameLabel.place(x=10, y=130)
                    self.DeviceIDFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceIDFrameEntry['width']=50
                    self.DeviceIDFrameEntry.place(x=110, y=130)
                    self.DeviceIDFrameEntry['state'] = DISABLED

                    
                    self.DeviceDescriptionFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceDescriptionFrameLabel["text"] = "Description:"
                    self.DeviceDescriptionFrameLabel.place(x=10, y=160)
                    self.DeviceDescriptionFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceDescriptionFrameEntry['width']=50
                    self.DeviceDescriptionFrameEntry.place(x=110, y=160)
                    self.DeviceDescriptionFrameEntry['state'] = DISABLED               

                    # ComboBox for Type, Model, Status
                    
                    TypeIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    TypeIDFrameLabel["text"] = "Type:"
                    TypeIDFrameLabel.place(x=10, y=190)
                    self.DeviceComboBoxTypeID = ttk.Combobox(self.DeviceFrame, state='disabled', width = 20)
                    self.DeviceComboBoxTypeID.bind("<<ComboboxSelected>>", self.on_device_type_combo_changed)
                    self.DeviceComboBoxTypeID.place(x = 60, y = 190)

                    self.ModelIDFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.ModelIDFrameLabel["text"] = "Model:"
                    self.ModelIDFrameLabel.place(x=220, y=190)
                    self.DeviceComboBoxModelID = ttk.Combobox(self.DeviceFrame, state='disabled', width = 20)
                    self.DeviceComboBoxModelID.place(x = 280, y = 190)
                    #self.DeviceComboBoxModelID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)

                    self.DeviceIPAddressFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceIPAddressFrameLabel["text"] = "IP Address:"
                    self.DeviceIPAddressFrameLabel.place(x=430, y=190)
                    self.DeviceIPAddressFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceIPAddressFrameEntry['width']=17
                    self.DeviceIPAddressFrameEntry.place(x=510, y=190)
                    self.DeviceIPAddressFrameEntry['state'] = DISABLED
                    self.DeviceIPAddressFrameEntry.bind("<Double-1>", self.Get_IPAddress)

                    self.DeviceStatusFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceStatusFrameLabel["text"] = "Status:"
                    self.DeviceStatusFrameLabel.place(x=625, y=190)
                    self.DeviceComboBoxStatus = ttk.Combobox(self.DeviceFrame, state='disabled', width = 20)
                    #self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.DeviceComboBoxStatus.place(x = 675, y = 190)


                    ##########################################

                    self.DeviceNATIPAddressFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceNATIPAddressFrameLabel["text"] = "NAT IP Address:"
                    self.DeviceNATIPAddressFrameLabel.place(x=925, y=190)
                    self.DeviceNATIPAddressFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceNATIPAddressFrameEntry['width']=17
                    self.DeviceNATIPAddressFrameEntry.place(x=1040, y=190)
                    self.DeviceNATIPAddressFrameEntry['state'] = DISABLED
                    self.DeviceNATIPAddressFrameEntry.bind("<Double-1>", self.Get_NAT_IPAddress)

                    self.DeviceMgmtIPAddressFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceMgmtIPAddressFrameLabel["text"] = "Mgmt. IP Address:"
                    self.DeviceMgmtIPAddressFrameLabel.place(x=925, y=220)
                    self.DeviceMgmtIPAddressFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceMgmtIPAddressFrameEntry['width']=17
                    self.DeviceMgmtIPAddressFrameEntry.place(x=1040, y=220)
                    self.DeviceMgmtIPAddressFrameEntry['state'] = DISABLED
                    self.DeviceMgmtIPAddressFrameEntry.bind("<Double-1>", self.Get_Mgmt_IPAddress)
                  
                    self.DeviceICMPCapableFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceICMPCapableFrameLabel["text"] = "ICMP Capable:"
                    self.DeviceICMPCapableFrameLabel.place(x=925, y=250)
                    self.DeviceComboBoxICMPCapable = ttk.Combobox(self.DeviceFrame, state='disabled', width = 14)
                    self.DeviceComboBoxICMPCapable.place(x = 1040, y = 250)

                    self.DeviceICMPMonitorFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceICMPMonitorFrameLabel["text"] = "Monitor Via ICMP:"
                    self.DeviceICMPMonitorFrameLabel.place(x=925, y=280)
                    self.DeviceComboBoxICMPMonitor = ttk.Combobox(self.DeviceFrame, state='disabled', width = 14)
                    self.DeviceComboBoxICMPMonitor.place(x = 1040, y = 280)

                    self.DeviceLastICMPDateFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceLastICMPDateFrameLabel["text"] = "Last ICMP on:"
                    self.DeviceLastICMPDateFrameLabel.place(x=925, y=310)
                    self.DeviceLastICMPDateFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceLastICMPDateFrameEntry['width']=17
                    self.DeviceLastICMPDateFrameEntry.place(x=1040, y=310)
                    self.DeviceLastICMPDateFrameEntry['state'] = DISABLED

                    self.DeviceLastICMPStatusFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceLastICMPStatusFrameLabel["text"] = "Was:"
                    self.DeviceLastICMPStatusFrameLabel.place(x=1150, y=310)
                    self.DeviceLastICMPStatusFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceLastICMPStatusFrameEntry['width']=22
                    self.DeviceLastICMPStatusFrameEntry.place(x=1190, y=310)
                    self.DeviceLastICMPStatusFrameEntry['state'] = DISABLED

                    self.DeviceLastCMDBStatusFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceLastCMDBStatusFrameLabel["text"] = "CMDB Last Update:"
                    self.DeviceLastCMDBStatusFrameLabel.place(x=625, y=310)
                    self.DeviceLastCMDBStatusFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceLastCMDBStatusFrameEntry['width']=17
                    self.DeviceLastCMDBStatusFrameEntry.place(x=750, y=310)
                    self.DeviceLastCMDBStatusFrameEntry['state'] = DISABLED
                    
                    ##########################################

                    # Setup Labels and Button Calendars Installed, Activated, Disconnected
                    self.DeviceInstalledDateFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceInstalledDateFrameLabel["text"] = "Install or Buy Date:" #"Installed Date:"
                    self.DeviceInstalledDateFrameLabel.place(x=10, y=220)
                    self.DeviceInstalledDateFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceInstalledDateFrameEntry['width']=20
                    self.DeviceInstalledDateFrameEntry.place(x=130, y=220)
                    self.DeviceInstalledDateFrameEntry['state'] = DISABLED
                    self.DeviceButtonInstalledDate = Button(self.DeviceFrame, text = 'Calendar', command = self.Call_Button_Installed_Date, state=DISABLED)
                    self.DeviceButtonInstalledDate.place(x = 260 , y = 217, width=75, height=25)

                    self.DeviceContractNoFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceContractNoFrameLabel["text"] = "Contract No:"
                    self.DeviceContractNoFrameLabel.place(x=430, y=220)
                    self.DeviceContractNoFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceContractNoFrameEntry['width']=17
                    self.DeviceContractNoFrameEntry.place(x=510, y=220)
                    self.DeviceContractNoFrameEntry['state'] = DISABLED

                    self.DeviceExpirationDateFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceExpirationDateFrameLabel["text"] = "Expiration Date:"
                    self.DeviceExpirationDateFrameLabel.place(x=625, y=220)
                    self.DeviceExpirationDateFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceExpirationDateFrameEntry['width']=15
                    self.DeviceExpirationDateFrameEntry.place(x=725, y=220)
                    self.DeviceExpirationDateFrameEntry['state'] = DISABLED
                    self.DeviceButtonExpirationDate = Button(self.DeviceFrame, text = 'Calendar', command = self.Call_Button_Expiration_Date, state=DISABLED)
                    self.DeviceButtonExpirationDate.place(x = 825 , y = 217, width=75, height=25)

                    self.DeviceActivatedDateFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceActivatedDateFrameLabel["text"] = "Activated Date:"
                    self.DeviceActivatedDateFrameLabel.place(x=10, y=250)
                    self.DeviceActivatedDateFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceActivatedDateFrameEntry['width']=20
                    self.DeviceActivatedDateFrameEntry.place(x=130, y=250)
                    self.DeviceActivatedDateFrameEntry['state'] = DISABLED
                    self.DeviceButtonActivatedDate = Button(self.DeviceFrame, text = 'Calendar', command = self.Call_Button_Activated_Date, state=DISABLED)
                    self.DeviceButtonActivatedDate.place(x = 260 , y = 247, width=75, height=25)

                    self.DeviceSerialNoFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceSerialNoFrameLabel["text"] = "Serial No:"
                    self.DeviceSerialNoFrameLabel.place(x=430, y=250)
                    self.DeviceSerialNoFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceSerialNoFrameEntry['width']=17
                    self.DeviceSerialNoFrameEntry.place(x=510, y=250)
                    self.DeviceSerialNoFrameEntry['state'] = DISABLED

                    self.DeviceMACAddressFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceMACAddressFrameLabel["text"] = "MAC Address:"
                    self.DeviceMACAddressFrameLabel.place(x=625, y=250)
                    self.DeviceMACAddressFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceMACAddressFrameEntry['width']=20
                    self.DeviceMACAddressFrameEntry.place(x=725, y=250)
                    self.DeviceMACAddressFrameEntry['state'] = DISABLED

                    self.DeviceDisconnectedDateFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceDisconnectedDateFrameLabel["text"] = "Disconnected Date:"
                    self.DeviceDisconnectedDateFrameLabel.place(x=10, y=280)
                    self.DeviceDisconnectedDateFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceDisconnectedDateFrameEntry['width']=20
                    self.DeviceDisconnectedDateFrameEntry.place(x=130, y=280)
                    self.DeviceDisconnectedDateFrameEntry['state'] = DISABLED
                    self.DeviceButtonDisconnectedDate = Button(self.DeviceFrame, text = 'Calendar', command = self.Call_Button_Disconnected_Date, state=DISABLED)
                    self.DeviceButtonDisconnectedDate.place(x = 260 , y = 277, width=75, height=25)

                    self.DeviceOutSourceCostFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceOutSourceCostFrameLabel["text"] = "Outsource $:"
                    self.DeviceOutSourceCostFrameLabel.place(x=430, y=280)
                    self.DeviceOutSourceCostFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceOutSourceCostFrameEntry['width']=17
                    self.DeviceOutSourceCostFrameEntry.place(x=510, y=280)
                    self.DeviceOutSourceCostFrameEntry['state'] = DISABLED

                    self.DeviceMaintenanceCostFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceMaintenanceCostFrameLabel["text"] = "Maintenance $:"
                    self.DeviceMaintenanceCostFrameLabel.place(x=625, y=280)
                    self.DeviceMaintenanceCostFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceMaintenanceCostFrameEntry['width']=20
                    self.DeviceMaintenanceCostFrameEntry.place(x=725, y=280)
                    self.DeviceMaintenanceCostFrameEntry['state'] = DISABLED

                    self.DeviceNotesFrameLabel = Label(self.DeviceFrame,text="Helvetica", font=("Helvetica", 10))
                    self.DeviceNotesFrameLabel["text"] = "NOTES:"
                    self.DeviceNotesFrameLabel.place(x=10, y=310)
                    self.DeviceNotesFrameEntry = Entry(self.DeviceFrame)
                    self.DeviceNotesFrameEntry['width']=80
                    self.DeviceNotesFrameEntry.place(x=130, y=310)
                    self.DeviceNotesFrameEntry['state']=DISABLED

                    
                    #------------------ TREE VIEW For Device Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars                            
                    self.DeviceTreeviewDataColumns = ('Device Name','Description','Type','Model','IP Address',
                                                       'OutSrc. Cost','Mnt. Cost','Installed',
                                                       'Activated','Disconected','Status','Contract No','Expiration',
                                                       'Serial No.','Added By')
                    
                    self.DeviceTreeview = ttk.Treeview(self.DeviceFrame,columns=self.DeviceTreeviewDataColumns, height=14) # <--- Make sure the frame is correct !!
                    self.DeviceTreeviewysb = Scrollbar(self.DeviceFrame,orient=VERTICAL, command=self.DeviceTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.DeviceTreeviewxsb = Scrollbar(self.DeviceFrame,orient=HORIZONTAL, command=self.DeviceTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.DeviceTreeview['yscroll'] = self.DeviceTreeviewysb.set
                    self.DeviceTreeview['xscroll'] = self.DeviceTreeviewxsb.set

                    # setup headings and column 
                    self.DeviceTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                    self.DeviceTreeview.heading('#1', text='Device Name',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Device Name', False)) # E for East and W for West
                    self.DeviceTreeview.heading('#2', text='Description',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Description', False)) # E for East and W for West
                    self.DeviceTreeview.heading('#3', text='Type', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Type', False)) # E for East and W for West
                    self.DeviceTreeview.heading('#4', text='Model', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Model', False)) # E for East and W for West
                    self.DeviceTreeview.heading('#5', text='IP Address', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'IP Address', False))
                    self.DeviceTreeview.heading('#6', text='OutSrc. Cost', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'OutSrc. Cost', False))
                    self.DeviceTreeview.heading('#7', text='Mnt. Cost', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Mnt. Cost', False))
                    self.DeviceTreeview.heading('#8', text='Installed', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Installed', False))
                    self.DeviceTreeview.heading('#9', text='Activated', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Activated', False))
                    self.DeviceTreeview.heading('#10', text='Disconected', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Disconected', False))
                    self.DeviceTreeview.heading('#11', text='Status', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Status', False))
                    self.DeviceTreeview.heading('#12', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Contract No', False))
                    self.DeviceTreeview.heading('#13', text='Expiration', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Expiration', False))
                    self.DeviceTreeview.heading('#14', text='Serial No.', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Serial No.', False))
                    self.DeviceTreeview.heading('#15', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Added By', False)) 
                    
                    self.DeviceTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.DeviceTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.DeviceTreeview.column('#2', stretch=1, width=10)
                    self.DeviceTreeview.column('#3', stretch=1, width=10)
                    self.DeviceTreeview.column('#4', stretch=1, width=10)
                    self.DeviceTreeview.column('#5', stretch=1, width=10)
                    self.DeviceTreeview.column('#6', stretch=1, width=10)
                    self.DeviceTreeview.column('#7', stretch=1, width=10)
                    self.DeviceTreeview.column('#8', stretch=1, width=10)
                    self.DeviceTreeview.column('#9', stretch=1, width=10)
                    self.DeviceTreeview.column('#10', stretch=1, width=10)
                    self.DeviceTreeview.column('#11', stretch=1, width=10)
                    self.DeviceTreeview.column('#12', stretch=1, width=10)
                    self.DeviceTreeview.column('#13', stretch=1, width=10)
                    self.DeviceTreeview.column('#14', stretch=1, width=10)
                    self.DeviceTreeview.column('#15', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.DeviceTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.DeviceTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.DeviceTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    # Use later to mark Business Units per color.
                    self.DeviceTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                    #self.DeviceTreeview.tag_configure('production', font=('Helvetica', 8), background='green2')
                    self.DeviceTreeview.tag_configure('obsoleted', font=('Helvetica', 8), background='ivory3') # 'red'
                    self.DeviceTreeview.tag_configure('obsolete', font=('Helvetica', 8), background='ivory3')  # 'red'
                    self.DeviceTreeview.tag_configure('implementation', font=('Helvetica', 8), background='green3')
                    self.DeviceTreeview.tag_configure('pre-production', font=('Helvetica', 8), background='yellow2')
                    self.DeviceTreeview.tag_configure('up', font=('Helvetica', 8), background='green2')
                    self.DeviceTreeview.tag_configure('down', font=('Helvetica', 8), background='pink')

                   
                    # Bind the double Click
                    self.DeviceTreeview.bind('<ButtonRelease-1>', self.on_Device_Tree_select_click) # When Select the Tree
                    #self.DeviceTreeview.bind("<Double-1>", self.On_Device_Tree_Refresh)         

                    #------------------ TREE VIEW For Device Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.DeviceFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.DeviceFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 40)

                    self.ComboBoxFacilityID = ttk.Combobox(self.DeviceFrame, state='disabled', width = 50)
                    self.ComboBoxFacilityID.bind("<<ComboboxSelected>>", self.on_facility_combo_changed)
                    self.ComboBoxFacilityID.place(x = 110, y = 70)

                    self.ComboBoxSitesID = ttk.Combobox(self.DeviceFrame, state='disabled', width = 50)
                    self.ComboBoxSitesID.bind("<<ComboboxSelected>>", self.on_sites_combo_changed)
                    self.ComboBoxSitesID.place(x = 110, y = 100)

                    # Get the Type and Models
                    self.Get_Type_Model_and_Satus()
                    if (self.Is_Get_Type_and_Model):
                        if (self.Go_To_Location):
                            self.CountryID_Pre = self.Init_Country
                            self.RegionID_Pre = self.Init_Region
                            self.FacilityID_Pre = self.Init_Facility
                            self.SitesID_Pre = self.Init_Site
                            self.Selection = 'cancel_edit'
                            self.on_Country_Table_Refresh()
                            if (self.sql_querry):
                                self.ComboBoxCoutryID.current(0)            
                                self.on_Region_Table_Refresh()
                                if (self.sql_querry):
                                    self.ComboBoxRegionID.current(0)
                                    self.on_Facility_Table_Refresh()
                                    if (self.sql_querry):
                                        self.ComboBoxFacilityID.current(0)
                                        self.on_Sites_Table_Refresh()
                                        if (self.sql_querry):
                                            self.ComboBoxSitesID.current(0)
                                            self.on_sites_combo_changed("event")
                            self.Selection = 'edit_ok'
                        else:                        
                            self.on_Country_Table_Refresh()
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.DeviceWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.DeviceFrame,title='Device',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Device Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    Device = Class_Device("BV",Windows_Scaling,location)
    Device.Display_Device_Window()


if __name__ == '__main__':
    Main()

