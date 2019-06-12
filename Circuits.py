#==========================================================================
# Program: Circuits.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-21-2018
# Date Last Modified: Jan-21-2018
# Summary: This is Class to for the Circuits
#==========================================================================

#***************************************************************
# ==================== Libraries Required <BEGIN> =============*
#***************************************************************

#************************ For PING ************************
import re  # Required for teh Class
import subprocess
from time import time, sleep
try:
    import socket
    import threading
#    fromthreading import *
except:
    print ("NO Sockets is available")
#************************ For PING ************************

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
    from Carriers import *
    Is_Carriers_Available = True
    Carrier = Class_Carrier(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************** \n")
    print ("*** NO Carriers Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************** \n")
    Is_Carriers_Available = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Circuits','OPEN Window']    
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

class Class_Circuits:

    def __init__(self,DSN_Name,Windows_Scaling,Location):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.CircuitsWindowExist = False
        self.CircuitsCalendarInstalledDateExist = False
        self.CircuitsCalendarActivatedDateExist = False
        self.CircuitsCalendarDisconnectedDateExist = False
        self.CircuitsCalendarExpirationDateExist = False
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

        #self.CircuitsBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False
        if (option == 'country'): ## The option are country,region and Circuits
            self.ComboBoxRegionID.set("")
            self.ComboBoxRegionID['state'] = DISABLED
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            self.ComboBoxSitesID.set("")
            self.ComboBoxSitesID['state'] = DISABLED
            self.ButtonCircuitsOK['state'] = DISABLED # Due to Edit
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
            self.ButtonCircuitsOK['state'] = DISABLED # Due to Edit
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
            self.ButtonCircuitsOK['state'] = DISABLED # Due to Edit
            if (option2 != 'facility-combo'):
                self.ComboBoxFacilityID.set("")
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = DISABLED
                self.ButtonSitesRefresh['state'] = DISABLED

        if (option2 == 'sites-combo'):
            if (self.Selection != 'edit'):
                self.ButtonCircuitsAdd['state'] = ACTIVE
            if (self.Selection == 'edit'):
                self.ButtonCircuitsOK['state'] = ACTIVE # Due to Edit
        else:
            self.ButtonCircuitsAdd['state'] = DISABLED

        if (self.Selection != 'edit'):
            self.ButtonCircuitsEdit['state'] = DISABLED
            self.ButtonCircuitsRemove['state'] = DISABLED
            self.ButtonCircuitsOK['state'] = DISABLED
            self.ButtonCircuitsCancel['state'] = DISABLED
            self.ButtonDevicePingPE64['state'] = DISABLED
            self.ButtonDevicePingPE1500['state'] = DISABLED
            self.ButtonDevicePingCE64['state'] = DISABLED
            self.ButtonDevicePingCE1500['state'] = DISABLED
            self.ButtonDeviceCircuits['state'] = DISABLED
            self.ButtonDeviceTraceroutePE['state'] = DISABLED
            self.ButtonDeviceTracerouteCE['state'] = DISABLED
            self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED


            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.CircuitIDFrameEntry['state'] = 'normal'
            self.CircuitIDFrameEntry.delete(0,END)
            self.CircuitIDFrameEntry['state'] = 'readonly'
                
            self.CircuitsDescriptionFrameEntry['state'] = 'normal'
            self.CircuitsDescriptionFrameEntry.delete(0,END)
            self.CircuitsDescriptionFrameEntry['state'] = 'readonly'

            self.CircuitComboBoxTypeID.current(0)
            self.CircuitComboBoxTypeID['state'] = DISABLED

            self.CircuitsComboBoxPortSpeed.current(0)
            self.CircuitsComboBoxPortSpeed['state'] = DISABLED
            
            self.CircuitsComboBoxStatus.current(0)
            self.CircuitsComboBoxStatus['state'] = DISABLED

            self.CircuitComboBoxTerm.current(0)
            self.CircuitComboBoxTerm['state'] = DISABLED

            self.CircuitComboBoxCarrier.current(0)
            self.CircuitComboBoxCarrier['state'] = DISABLED

            self.CircuitsBandwidthFrameEntry['state'] = 'normal'
            self.CircuitsBandwidthFrameEntry.delete(0,END)
            self.CircuitsBandwidthFrameEntry['state'] = 'readonly'

            self.CircuitTermFrameEntry['state'] = 'normal'
            self.CircuitTermFrameEntry.delete(0,END)
            self.CircuitTermFrameEntry['state'] = 'readonly'

            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry.delete(0,END)
            self.CircuitsContractNoFrameEntry['state'] = 'readonly'
            
            self.CircuitAccountNoFrameEntry['state'] = 'normal'
            self.CircuitAccountNoFrameEntry.delete(0,END)
            self.CircuitAccountNoFrameEntry['state'] = 'readonly'

            self.CircuitOrderNoFrameEntry['state'] = 'normal'
            self.CircuitOrderNoFrameEntry.delete(0,END)
            self.CircuitOrderNoFrameEntry['state'] = 'readonly'

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)
            self.CircuitCEASNFrameEntry['state'] = 'readonly'         

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)
            self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitVLANNoFrameEntry['state'] = 'normal'
            self.CircuitVLANNoFrameEntry.delete(0,END)
            self.CircuitVLANNoFrameEntry['state'] = 'readonly'

            self.CircuitNPANXXFrameEntry['state'] = 'normal'
            self.CircuitNPANXXFrameEntry.delete(0,END)
            self.CircuitNPANXXFrameEntry['state'] = 'readonly'

            self.CircuitPEASNFrameEntry['state'] = 'normal'
            self.CircuitPEASNFrameEntry.delete(0,END)
            self.CircuitPEASNFrameEntry['state'] = 'readonly'

            self.CircuitPEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitPEIPAddressFrameEntry.delete(0,END)
            self.CircuitPEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitPESwitchFrameEntry['state'] = 'normal'
            self.CircuitPESwitchFrameEntry.delete(0,END)
            self.CircuitPESwitchFrameEntry['state'] = 'readonly'

            self.CircuitPELocationFrameEntry['state'] = 'normal'
            self.CircuitPELocationFrameEntry.delete(0,END)
            self.CircuitPELocationFrameEntry['state'] = 'readonly'

            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)
            self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)
            self.CircuitETFFrameEntry['state'] = 'readonly'

            self.CircuitLEC1FrameEntry['state'] = 'normal'
            self.CircuitLEC1FrameEntry.delete(0,END)
            self.CircuitLEC1FrameEntry['state'] = 'readonly'

            self.CircuitLEC2FrameEntry['state'] = 'normal'
            self.CircuitLEC2FrameEntry.delete(0,END)
            self.CircuitLEC2FrameEntry['state'] = 'readonly'

            self.CircuitLEC3FrameEntry['state'] = 'normal'
            self.CircuitLEC3FrameEntry.delete(0,END)
            self.CircuitLEC3FrameEntry['state'] = 'readonly'

            self.CircuitLEC4FrameEntry['state'] = 'normal'
            self.CircuitLEC4FrameEntry.delete(0,END)
            self.CircuitLEC4FrameEntry['state'] = 'readonly'

            self.CircuitLEC5FrameEntry['state'] = 'normal'
            self.CircuitLEC5FrameEntry.delete(0,END)
            self.CircuitLEC5FrameEntry['state'] = 'readonly'

            self.CircuitDMARK1FrameEntry['state'] = 'normal'
            self.CircuitDMARK1FrameEntry.delete(0,END)
            self.CircuitDMARK1FrameEntry['state'] = 'readonly'

            self.CircuitDMARK2FrameEntry['state'] = 'normal'
            self.CircuitDMARK2FrameEntry.delete(0,END)
            self.CircuitDMARK2FrameEntry['state'] = 'readonly'


            # Setup Labels and Button Calendars Installed, Activated, Disconnected

            self.CircuitsButtonInstalledDate['state'] = DISABLED
            self.CircuitsButtonActivatedDate['state'] = DISABLED
            self.CircuitsButtonDisconnectedDate['state'] = DISABLED
            self.CircuitsButtonExpirationDate['state'] = DISABLED

            self.CircuitsButtonInstalledDateClear['state'] = DISABLED
            self.CircuitsButtonActivatedDateClear['state'] = DISABLED
            self.CircuitsButtonExpirationDateClear['state'] = DISABLED
            self.CircuitsButtonDisconnectedDateClear['state'] = DISABLED

            self.CircuitsInstalledDateFrameEntry['state'] = 'normal'
            self.CircuitsInstalledDateFrameEntry.delete(0,END)
            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'

            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry.delete(0,END)
            self.CircuitsContractNoFrameEntry['state'] = 'readonly'

            self.CircuitsExpirationDateFrameEntry['state'] = 'normal'
            self.CircuitsExpirationDateFrameEntry.delete(0,END)
            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'

            self.CircuitsActivatedDateFrameEntry['state'] = 'normal'
            self.CircuitsActivatedDateFrameEntry.delete(0,END)
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)
            self.CircuitCEASNFrameEntry['state'] = 'readonly'

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)
            self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitsDisconnectedDateFrameEntry['state'] = 'normal'
            self.CircuitsDisconnectedDateFrameEntry.delete(0,END)
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'

            # COST Entries (Float)
            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)
            self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)
            self.CircuitETFFrameEntry['state'] = 'readonly'

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.CircuitsTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.CircuitsTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

    def Display_Screen(self,curItem):      
            # Create Progress Bar
            self.Get_Type_PortSpeed_and_Satus()
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            '''
                self.CircuitsTablePriaryKeyArray =  [] # Circuits ID
                self.CircuitsTableDescriptionArray =  []                
                self.CircuitsTableCountryIDArray =  [] 
                self.CircuitsTableRegionIDArray =  []
                self.CircuitsTableFacilityIDArray =  []
                self.CircuitsTableSiteIDArray =  []
                self.CircuitsTableCarrierIDArray =  []
                self.CircuitsTableCircuitTypeArray = []
                self.CircuitsTablePortSpeedArray = []
                self.CircuitsTableBandwidthArray = []
                self.CircuitsTableStatusArray = []
                self.CircuitsTableDmarc_Info_1Array = []
                self.CircuitsTableDmarc_Info_2Array = []
                self.CircuitsTableLEC1Array = []
                self.CircuitsTableLEC2Array = []
                self.CircuitsTableLEC3Array = []
                self.CircuitsTableLEC4Array = []
                self.CircuitsTableLEC5Array = []
                self.CircuitsTableCE_ASNArray = []
                self.CircuitsTableCE_IP_AddressArray = []
                self.CircuitsTablePE_ASNArray = []
                self.CircuitsTablePE_IP_AddressArray = []
                self.CircuitsTableVLAN_IDArray = []
                self.CircuitsTablePE_SwitchArray = []
                self.CircuitsTablePE_LocationArray = []
                self.CircuitsTableNPA_NXXArray = []
                self.CircuitsTableMonthlyCostArray = []
                self.CircuitsTableOrderNumberArray = []
                self.CircuitsTableDateInstalledArray = []
                self.CircuitsTableDayInstalledArray = []
                self.CircuitsTableMonthInstalledArray = []
                self.CircuitsTableYearInstalledArray = []
                self.CircuitsTableDateActivatedArray = []
                self.CircuitsTableDayActivatedArray = []
                self.CircuitsTableMonthActivatedArray = []
                self.CircuitsTableYearActivatedArray = []
                self.CircuitsTableDisconectedDateArray = []
                self.CircuitsTableDayDisconectedArray = []
                self.CircuitsTableMonthDisconectedArray = []
                self.CircuitsTableYearDisconectedArray = []
                self.CircuitsTableExpirationDateArray = []
                self.CircuitsTableDayExpirationArray = []
                self.CircuitsTableMonthExpirationArray = []
                self.CircuitsTableYearExpirationArray = []
                self.CircuitsTableTerm_DayArray = []
                self.CircuitsTableTerm_TimeArray = []
                self.CircuitsTableETFArray = []
                self.CircuitsTableContract_NoArray = []
                self.CircuitsTableAccount_NoArray = []
                self.CircuitsTableExecutedByArray = []

            '''
            # Setup Labels and Entry
            self.CircuitIDFrameEntry['state'] = 'normal'
            self.CircuitIDFrameEntry.delete(0,END)
            self.CircuitIDFrameEntry.insert(0,self.CircuitsTablePriaryKeyArray[curItem])
            self.CircuitIDFrameEntry['state'] = 'readonly'
                
            self.CircuitsDescriptionFrameEntry['state'] = 'normal'
            self.CircuitsDescriptionFrameEntry.delete(0,END)
            self.CircuitsDescriptionFrameEntry.insert(0,self.CircuitsTableDescriptionArray[curItem])
            self.CircuitsDescriptionFrameEntry['state'] = 'readonly'

            # Find Circuit Type in the Array
            i = 0
            self.CircuitComboBoxTypeID.current(i)
            while (i < len(self.CircuitTypeIDArray)):
                if (self.CircuitsTableCircuitTypeArray[curItem] == self.CircuitTypeIDArray[i]):
                    self.CircuitComboBoxTypeID.current(i)
                    i = i + len(self.CircuitTypeIDArray) 
                else:
                    i = i + 1
            # find Port Speed in the Array
            i = 0
            self.CircuitsComboBoxPortSpeed.current(i)
            while (i < len(self.CircuitPortSpeedIDArray)):
                if (self.CircuitsTablePortSpeedArray[curItem] == self.CircuitPortSpeedIDArray[i]):
                    self.CircuitsComboBoxPortSpeed.current(i)
                    i = i + len(self.CircuitPortSpeedIDArray) 
                else:
                    i = i + 1
            # find Status in the Array
            i = 0
            while (i < len(self.CircuitstatusValues)):
                if (self.CircuitsTableStatusArray[curItem] == self.CircuitstatusValues[i]):
                    self.CircuitsComboBoxStatus.current(i)
                    i = i + len(self.CircuitstatusValues) 
                else:
                    i = i + 1

            # find Term in the Array
            i = 0
            while (i < len(self.CircuitTermValues)):
                if (self.CircuitsTableTerm_TimeArray[curItem] == self.CircuitTermValues[i]):
                    self.CircuitComboBoxTerm.current(i)
                    i = i + len(self.CircuitTermValues) 
                else:
                    i = i + 1
                    
            # find Carrier in the Array
            i = 0
            while (i < len(self.CircuitCarrierNameArray)):
                if (self.CircuitsTableCarrierIDArray[curItem] == self.CircuitCarrierNameArray[i]):
                    self.CircuitComboBoxCarrier.current(i)
                    i = i + len(self.CircuitCarrierNameArray) 
                else:
                    i = i + 1

            self.CircuitsBandwidthFrameEntry['state'] = 'normal'
            self.CircuitsBandwidthFrameEntry.delete(0,END)
            self.CircuitsBandwidthFrameEntry.insert(0,self.CircuitsTableBandwidthArray[curItem])
            self.CircuitsBandwidthFrameEntry['state'] = 'readonly'

            self.CircuitTermFrameEntry['state'] = 'normal'
            self.CircuitTermFrameEntry.delete(0,END)
            self.CircuitTermFrameEntry.insert(0,self.CircuitsTableTerm_DayArray[curItem])
            self.CircuitTermFrameEntry['state'] = 'readonly'

            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry.delete(0,END)
            self.CircuitsContractNoFrameEntry.insert(0,self.CircuitsTableContract_NoArray[curItem])
            self.CircuitsContractNoFrameEntry['state'] = 'readonly'
            
            self.CircuitAccountNoFrameEntry['state'] = 'normal'
            self.CircuitAccountNoFrameEntry.delete(0,END)
            self.CircuitAccountNoFrameEntry.insert(0,self.CircuitsTableAccount_NoArray[curItem])
            self.CircuitAccountNoFrameEntry['state'] = 'readonly'

            self.CircuitOrderNoFrameEntry['state'] = 'normal'
            self.CircuitOrderNoFrameEntry.delete(0,END)
            self.CircuitOrderNoFrameEntry.insert(0,self.CircuitsTableOrderNumberArray[curItem])
            self.CircuitOrderNoFrameEntry['state'] = 'readonly'

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)
            self.CircuitCEASNFrameEntry.insert(0,self.CircuitsTableCE_ASNArray[curItem])
            self.CircuitCEASNFrameEntry['state'] = 'readonly'         

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)
            self.CircuitCEIPAddressFrameEntry.insert(0,self.CircuitsTableCE_IP_AddressArray[curItem])
            self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitVLANNoFrameEntry['state'] = 'normal'
            self.CircuitVLANNoFrameEntry.delete(0,END)
            self.CircuitVLANNoFrameEntry.insert(0,self.CircuitsTableVLAN_IDArray[curItem])
            self.CircuitVLANNoFrameEntry['state'] = 'readonly'

            self.CircuitNPANXXFrameEntry['state'] = 'normal'
            self.CircuitNPANXXFrameEntry.delete(0,END)
            self.CircuitNPANXXFrameEntry.insert(0,self.CircuitsTableNPA_NXXArray[curItem])
            self.CircuitNPANXXFrameEntry['state'] = 'readonly'

            self.CircuitPEASNFrameEntry['state'] = 'normal'
            self.CircuitPEASNFrameEntry.delete(0,END)
            self.CircuitPEASNFrameEntry.insert(0,self.CircuitsTablePE_ASNArray[curItem])
            self.CircuitPEASNFrameEntry['state'] = 'readonly'

            self.CircuitPEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitPEIPAddressFrameEntry.delete(0,END)
            self.CircuitPEIPAddressFrameEntry.insert(0,self.CircuitsTablePE_IP_AddressArray[curItem])
            self.CircuitPEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitPESwitchFrameEntry['state'] = 'normal'
            self.CircuitPESwitchFrameEntry.delete(0,END)
            self.CircuitPESwitchFrameEntry.insert(0,self.CircuitsTablePE_SwitchArray[curItem])
            self.CircuitPESwitchFrameEntry['state'] = 'readonly'

            self.CircuitPELocationFrameEntry['state'] = 'normal'
            self.CircuitPELocationFrameEntry.delete(0,END)
            self.CircuitPELocationFrameEntry.insert(0,self.CircuitsTablePE_LocationArray[curItem])
            self.CircuitPELocationFrameEntry['state'] = 'readonly'

            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)
            self.CircuitMonthlyCostFrameEntry.insert(0,self.CircuitsTableMonthlyCostArray[curItem])
            self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)
            self.CircuitETFFrameEntry.insert(0,self.CircuitsTableETFArray[curItem])
            self.CircuitETFFrameEntry['state'] = 'readonly'

            self.CircuitLEC1FrameEntry['state'] = 'normal'
            self.CircuitLEC1FrameEntry.delete(0,END)
            self.CircuitLEC1FrameEntry.insert(0,self.CircuitsTableLEC1Array[curItem])
            self.CircuitLEC1FrameEntry['state'] = 'readonly'

            self.CircuitLEC2FrameEntry['state'] = 'normal'
            self.CircuitLEC2FrameEntry.delete(0,END)
            self.CircuitLEC2FrameEntry.insert(0,self.CircuitsTableLEC2Array[curItem])
            self.CircuitLEC2FrameEntry['state'] = 'readonly'

            self.CircuitLEC3FrameEntry['state'] = 'normal'
            self.CircuitLEC3FrameEntry.delete(0,END)
            self.CircuitLEC3FrameEntry.insert(0,self.CircuitsTableLEC3Array[curItem])
            self.CircuitLEC3FrameEntry['state'] = 'readonly'

            self.CircuitLEC4FrameEntry['state'] = 'normal'
            self.CircuitLEC4FrameEntry.delete(0,END)
            self.CircuitLEC4FrameEntry.insert(0,self.CircuitsTableLEC4Array[curItem])
            self.CircuitLEC4FrameEntry['state'] = 'readonly'

            self.CircuitLEC5FrameEntry['state'] = 'normal'
            self.CircuitLEC5FrameEntry.delete(0,END)
            self.CircuitLEC5FrameEntry.insert(0,self.CircuitsTableLEC5Array[curItem])
            self.CircuitLEC5FrameEntry['state'] = 'readonly'

            self.CircuitDMARK1FrameEntry['state'] = 'normal'
            self.CircuitDMARK1FrameEntry.delete(0,END)
            self.CircuitDMARK1FrameEntry.insert(0,self.CircuitsTableDmarc_Info_1Array[curItem])
            self.CircuitDMARK1FrameEntry['state'] = 'readonly'

            self.CircuitDMARK2FrameEntry['state'] = 'normal'
            self.CircuitDMARK2FrameEntry.delete(0,END)
            self.CircuitDMARK2FrameEntry.insert(0,self.CircuitsTableDmarc_Info_2Array[curItem])
            self.CircuitDMARK2FrameEntry['state'] = 'readonly'

            # Setup Labels and Button Calendars Installed, Activated, Disconnected

            self.CircuitsButtonInstalledDate['state'] = DISABLED
            self.CircuitsButtonActivatedDate['state'] = DISABLED
            self.CircuitsButtonDisconnectedDate['state'] = DISABLED
            self.CircuitsButtonExpirationDate['state'] = DISABLED

            self.CircuitsButtonInstalledDateClear['state'] = DISABLED
            self.CircuitsButtonActivatedDateClear['state'] = DISABLED
            self.CircuitsButtonExpirationDateClear['state'] = DISABLED
            self.CircuitsButtonDisconnectedDateClear['state'] = DISABLED
            

            self.CircuitsInstalledDateFrameEntry['state'] = 'normal'
            self.CircuitsInstalledDateFrameEntry.delete(0,END)
            if (self.CircuitsTableDateInstalledArray[curItem] == None):
                self.CircuitsInstalledDateFrameEntry.insert(0," ")
            else:
                self.CircuitsInstalledDateFrameEntry.insert(0,self.CircuitsTableDateInstalledArray[curItem])
            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'

            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry.delete(0,END)
            if (self.CircuitsTableContract_NoArray[curItem] == None):
                self.CircuitsContractNoFrameEntry.insert(0," ")
            else:
                self.CircuitsContractNoFrameEntry.insert(0,self.CircuitsTableContract_NoArray[curItem])
            self.CircuitsContractNoFrameEntry['state'] = 'readonly'

            self.CircuitsExpirationDateFrameEntry['state'] = 'normal'
            self.CircuitsExpirationDateFrameEntry.delete(0,END)
            if (self.CircuitsTableExpirationDateArray[curItem] == None):
                self.CircuitsExpirationDateFrameEntry.insert(0," ")
            else:
                self.CircuitsExpirationDateFrameEntry.insert(0,self.CircuitsTableExpirationDateArray[curItem])
            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'

            self.CircuitsActivatedDateFrameEntry['state'] = 'normal'
            self.CircuitsActivatedDateFrameEntry.delete(0,END)
            if (self.CircuitsTableDateActivatedArray[curItem] == None):
                self.CircuitsActivatedDateFrameEntry.insert(0," ")
            else:
                self.CircuitsActivatedDateFrameEntry.insert(0,self.CircuitsTableDateActivatedArray[curItem])
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)
            if (self.CircuitsTableCE_ASNArray[curItem] == None):
                self.CircuitCEASNFrameEntry.insert(0," ")
            else:
                self.CircuitCEASNFrameEntry.insert(0,self.CircuitsTableCE_ASNArray[curItem])
            self.CircuitCEASNFrameEntry['state'] = 'readonly'

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)
            if (self.CircuitsTableCE_IP_AddressArray[curItem] == None):
                self.CircuitCEIPAddressFrameEntry.insert(0," ")
            else:
                self.CircuitCEIPAddressFrameEntry.insert(0,self.CircuitsTableCE_IP_AddressArray[curItem])
            self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'

            self.CircuitsDisconnectedDateFrameEntry['state'] = 'normal'
            self.CircuitsDisconnectedDateFrameEntry.delete(0,END)
            if (self.CircuitsTableDisconectedDateArray[curItem] == None):
                self.CircuitsTableDisconectedDateArray.insert(0," ")
            else:
                self.CircuitsDisconnectedDateFrameEntry.insert(0,self.CircuitsTableDisconectedDateArray[curItem])
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'

            # COST Entries (Float)
            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)
            self.CircuitMonthlyCostFrameEntry.insert(0,str(self.CircuitsTableMonthlyCostArray[curItem]))
            self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)
            self.CircuitETFFrameEntry.insert(0,str(self.CircuitsTableETFArray[curItem]))
            self.CircuitETFFrameEntry['state'] = 'readonly'

    def Enable_Screen(self,option):
        # This function is used when the ADD button is selected

        #self.CircuitsBusinessUnitPowerCheckbutton.select()
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

            
        self.ButtonCircuitsAdd['state'] = DISABLED
        self.ButtonCircuitsEdit['state'] = DISABLED
        self.ButtonCircuitsRemove['state'] = DISABLED
        self.ButtonCircuitsOK['state'] = ACTIVE
        self.ButtonCircuitsCancel['state'] = ACTIVE
        self.ButtonDevicePingPE64['state'] = DISABLED
        self.ButtonDevicePingPE1500['state'] = DISABLED
        self.ButtonDevicePingCE64['state'] = DISABLED
        self.ButtonDevicePingCE1500['state'] = DISABLED
        self.ButtonDeviceCircuits['state'] = DISABLED
        self.ButtonDeviceTraceroutePE['state'] = DISABLED
        self.ButtonDeviceTracerouteCE['state'] = DISABLED
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

            self.Get_Type_PortSpeed_and_Satus()

            self.CircuitIDFrameEntry['state'] = 'normal'
            self.CircuitIDFrameEntry.delete(0,END)

            self.CircuitsDescriptionFrameEntry['state'] = 'normal'
            self.CircuitsDescriptionFrameEntry.delete(0,END)

            # Calendars:
            self.CircuitsButtonInstalledDate['state'] = ACTIVE
            self.CircuitsButtonActivatedDate['state'] = ACTIVE
            self.CircuitsButtonDisconnectedDate['state'] = ACTIVE
            self.CircuitsButtonExpirationDate['state'] = ACTIVE            

            self.CircuitsButtonInstalledDateClear['state'] = ACTIVE
            self.CircuitsButtonActivatedDateClear['state'] = ACTIVE
            self.CircuitsButtonExpirationDateClear['state'] = ACTIVE
            self.CircuitsButtonDisconnectedDateClear['state'] = ACTIVE

            self.CircuitsInstalledDateFrameEntry['state'] = 'normal'
            self.CircuitsInstalledDateFrameEntry.delete(0,END)
            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'
            
            self.CircuitsActivatedDateFrameEntry['state'] = 'normal'
            self.CircuitsActivatedDateFrameEntry.delete(0,END)
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
            
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'normal'
            self.CircuitsDisconnectedDateFrameEntry.delete(0,END)
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'

            self.CircuitsExpirationDateFrameEntry['state'] = 'normal'
            self.CircuitsExpirationDateFrameEntry.delete(0,END)
            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'

            # ComboBox

            self.CircuitComboBoxTypeID['state'] = DISABLED
            self.CircuitsComboBoxPortSpeed['state'] = DISABLED
            self.CircuitsComboBoxStatus['state'] = DISABLED
            self.CircuitComboBoxTerm['state'] = DISABLED
            self.CircuitComboBoxCarrier['state'] = DISABLED
            
            self.CircuitComboBoxTypeID['state'] = 'readonly'
            self.CircuitsComboBoxPortSpeed['state'] = 'readonly'
            self.CircuitsComboBoxStatus['state'] = 'readonly'
            self.CircuitComboBoxTerm['state'] = 'readonly'
            self.CircuitComboBoxCarrier['state'] = 'readonly'

            # Calendars
            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'
            self.CircuitsInstalledDateFrameEntry.delete(0,END)
            
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
            self.CircuitsActivatedDateFrameEntry.delete(0,END)

            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'
            self.CircuitsDisconnectedDateFrameEntry.delete(0,END)

            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'
            self.CircuitsExpirationDateFrameEntry.delete(0,END)

            self.CircuitsBandwidthFrameEntry['state'] = 'normal'
            self.CircuitsBandwidthFrameEntry.delete(0,END)

            # Setup Labels and Entry

            self.CircuitsBandwidthFrameEntry['state'] = 'normal'
            self.CircuitsBandwidthFrameEntry.delete(0,END)

            self.CircuitTermFrameEntry['state'] = 'normal'
            self.CircuitTermFrameEntry.delete(0,END)

            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry.delete(0,END)
            
            self.CircuitAccountNoFrameEntry['state'] = 'normal'
            self.CircuitAccountNoFrameEntry.delete(0,END)

            self.CircuitOrderNoFrameEntry['state'] = 'normal'
            self.CircuitOrderNoFrameEntry.delete(0,END)

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)

            self.CircuitVLANNoFrameEntry['state'] = 'normal'
            self.CircuitVLANNoFrameEntry.delete(0,END)

            self.CircuitNPANXXFrameEntry['state'] = 'normal'
            self.CircuitNPANXXFrameEntry.delete(0,END)

            self.CircuitPEASNFrameEntry['state'] = 'normal'
            self.CircuitPEASNFrameEntry.delete(0,END)

            self.CircuitPEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitPEIPAddressFrameEntry.delete(0,END)
            self.CircuitPEIPAddressFrameEntry.insert(0,"0.0.0.0")

            self.CircuitPESwitchFrameEntry['state'] = 'normal'
            self.CircuitPESwitchFrameEntry.delete(0,END)

            self.CircuitPELocationFrameEntry['state'] = 'normal'
            self.CircuitPELocationFrameEntry.delete(0,END)

            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)

            self.CircuitLEC1FrameEntry['state'] = 'normal'
            self.CircuitLEC1FrameEntry.delete(0,END)

            self.CircuitLEC2FrameEntry['state'] = 'normal'
            self.CircuitLEC2FrameEntry.delete(0,END)

            self.CircuitLEC3FrameEntry['state'] = 'normal'
            self.CircuitLEC3FrameEntry.delete(0,END)

            self.CircuitLEC4FrameEntry['state'] = 'normal'
            self.CircuitLEC4FrameEntry.delete(0,END)

            self.CircuitLEC5FrameEntry['state'] = 'normal'
            self.CircuitLEC5FrameEntry.delete(0,END)

            self.CircuitDMARK1FrameEntry['state'] = 'normal'
            self.CircuitDMARK1FrameEntry.delete(0,END)

            self.CircuitDMARK2FrameEntry['state'] = 'normal'
            self.CircuitDMARK2FrameEntry.delete(0,END)

            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry.delete(0,END)

            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry.delete(0,END)
            self.CircuitCEIPAddressFrameEntry.insert(0,"0.0.0.0")

            # COST Entries (Float)
            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry.delete(0,END)

            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry.delete(0,END)

        if (option == 'edit'): #<----------------------------------- EDIT Button
            self.ComboBoxCoutryID['state'] = ACTIVE
            self.ComboBoxRegionID['state'] = ACTIVE
            self.ComboBoxFacilityID['state'] = ACTIVE
            self.ComboBoxSitesID['state'] = ACTIVE

            #self.Get_Type_PortSpeed_and_Satus() # <------------------ I might have to modified it
            self.CircuitIDFrameEntry['state'] = 'readonly'
            self.CircuitsDescriptionFrameEntry['state'] = 'normal'
            # Calendars:
            self.CircuitsButtonInstalledDate['state'] = ACTIVE
            self.CircuitsButtonActivatedDate['state'] = ACTIVE
            self.CircuitsButtonDisconnectedDate['state'] = ACTIVE
            self.CircuitsButtonExpirationDate['state'] = ACTIVE

            self.CircuitsButtonInstalledDateClear['state'] = ACTIVE
            self.CircuitsButtonActivatedDateClear['state'] = ACTIVE
            self.CircuitsButtonExpirationDateClear['state'] = ACTIVE
            self.CircuitsButtonDisconnectedDateClear['state'] = ACTIVE

            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'            
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'
            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'

            # ComboBox:
            self.CircuitComboBoxTypeID['state'] = 'readonly'
            self.CircuitsComboBoxPortSpeed['state'] = 'readonly'
            self.CircuitsComboBoxStatus['state'] = 'readonly'
            self.CircuitComboBoxTerm['state'] = 'readonly'
            self.CircuitComboBoxCarrier['state'] = 'readonly'

            # Setup Labels and Entry
            self.CircuitsBandwidthFrameEntry['state'] = 'normal'
            self.CircuitTermFrameEntry['state'] = 'normal'
            self.CircuitsContractNoFrameEntry['state'] = 'normal'
            self.CircuitAccountNoFrameEntry['state'] = 'normal'
            self.CircuitOrderNoFrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitVLANNoFrameEntry['state'] = 'normal'
            self.CircuitNPANXXFrameEntry['state'] = 'normal'
            self.CircuitPEASNFrameEntry['state'] = 'normal'
            self.CircuitPEIPAddressFrameEntry['state'] = 'normal'
            self.CircuitPESwitchFrameEntry['state'] = 'normal'
            self.CircuitPELocationFrameEntry['state'] = 'normal'
            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry['state'] = 'normal'
            self.CircuitLEC1FrameEntry['state'] = 'normal'
            self.CircuitLEC2FrameEntry['state'] = 'normal'
            self.CircuitLEC3FrameEntry['state'] = 'normal'
            self.CircuitLEC4FrameEntry['state'] = 'normal'
            self.CircuitLEC5FrameEntry['state'] = 'normal'
            self.CircuitDMARK1FrameEntry['state'] = 'normal'
            self.CircuitDMARK2FrameEntry['state'] = 'normal'
            self.CircuitCEASNFrameEntry['state'] = 'normal'
            self.CircuitCEIPAddressFrameEntry['state'] = 'normal'

            # COST Entries (Float)
            self.CircuitMonthlyCostFrameEntry['state'] = 'normal'
            self.CircuitETFFrameEntry['state'] = 'normal'
                
       
    def Disable_Screen(self):
        # This function is used when the entry was added.modified to the Database

        #self.CircuitsBusinessUnitPowerCheckbutton.select()
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
            
        self.ButtonCircuitsAdd['state'] = ACTIVE
        self.ButtonCircuitsEdit['state'] = DISABLED
        self.ButtonCircuitsRemove['state'] = DISABLED
        self.ButtonCircuitsOK['state'] = DISABLED
        self.ButtonCircuitsCancel['state'] = DISABLED 
        self.ButtonDevicePingPE64['state'] = DISABLED
        self.ButtonDevicePingPE1500['state'] = DISABLED
        self.ButtonDevicePingCE64['state'] = DISABLED
        self.ButtonDevicePingCE1500['state'] = DISABLED
        self.ButtonDeviceCircuits['state'] = DISABLED
        self.ButtonDeviceTraceroutePE['state'] = DISABLED
        self.ButtonDeviceTracerouteCE['state'] = DISABLED
        self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED

        '''
        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.CircuitIDFrameEntry['state'] = 'readonly'                
        self.CircuitsDescriptionFrameEntry['state'] = 'readonly'
        self.CircuitComboBoxTypeID['state'] = 'readonly'
        self.CircuitsComboBoxPortSpeed['state'] = 'readonly'
        self.CircuitsComboBoxStatus['state'] = 'readonly'
        self.CircuitsBandwidthFrameEntry['state'] = 'readonly'
        self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'
        self.CircuitsContractNoFrameEntry['state'] = 'readonly'
        self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'
        self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
        self.CircuitCEASNFrameEntry['state'] = 'readonly'
        self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'
        self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'
        self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'
        self.CircuitETFFrameEntry['state'] = 'readonly'
        self.CircuitDMARK1FrameEntry['state'] = 'readonly'
        # Calendars:
        self.CircuitsButtonInstalledDate['state'] = DISABLED
        self.CircuitsButtonActivatedDate['state'] = DISABLED
        self.CircuitsButtonDisconnectedDate['state'] = DISABLED
        self.CircuitsButtonExpirationDate['state'] = DISABLED
        # Combobox:
        self.CircuitComboBoxTypeID['state'] = DISABLED
        self.CircuitsComboBoxPortSpeed['state'] = DISABLED
        self.CircuitsComboBoxStatus['state'] = DISABLED
        '''
#################3

        #self.Get_Type_PortSpeed_and_Satus() # <------------------ I might have to modified it
        self.CircuitIDFrameEntry['state'] = 'readonly'
        self.CircuitsDescriptionFrameEntry['state'] = 'readonly'
        # Calendars:
        self.CircuitsButtonInstalledDate['state'] = DISABLED
        self.CircuitsButtonActivatedDate['state'] = DISABLED
        self.CircuitsButtonDisconnectedDate['state'] = DISABLED
        self.CircuitsButtonExpirationDate['state'] = DISABLED

        self.CircuitsButtonInstalledDateClear['state'] = DISABLED
        self.CircuitsButtonActivatedDateClear['state'] = DISABLED
        self.CircuitsButtonExpirationDateClear['state'] = DISABLED
        self.CircuitsButtonDisconnectedDateClear['state'] = DISABLED

        self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'            
        self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
        self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'
        self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'

        # ComboBox:
        self.CircuitComboBoxTypeID['state'] = 'disabled'
        self.CircuitsComboBoxPortSpeed['state'] = 'disabled'
        self.CircuitsComboBoxStatus['state'] = 'disabled'
        self.CircuitComboBoxTerm['state'] = 'disabled'
        self.CircuitComboBoxCarrier['state'] = 'disabled'

        # Setup Labels and Entry
        self.CircuitsBandwidthFrameEntry['state'] = 'readonly'
        self.CircuitTermFrameEntry['state'] = 'readonly'
        self.CircuitsContractNoFrameEntry['state'] = 'readonly'
        self.CircuitAccountNoFrameEntry['state'] = 'readonly'
        self.CircuitOrderNoFrameEntry['state'] = 'readonly'
        self.CircuitCEASNFrameEntry['state'] = 'readonly'
        self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'
        self.CircuitVLANNoFrameEntry['state'] = 'readonly'
        self.CircuitNPANXXFrameEntry['state'] = 'readonly'
        self.CircuitPEASNFrameEntry['state'] = 'readonly'
        self.CircuitPEIPAddressFrameEntry['state'] = 'readonly'
        self.CircuitPESwitchFrameEntry['state'] = 'readonly'
        self.CircuitPELocationFrameEntry['state'] = 'readonly'
        self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'
        self.CircuitETFFrameEntry['state'] = 'readonly'
        self.CircuitLEC1FrameEntry['state'] = 'readonly'
        self.CircuitLEC2FrameEntry['state'] = 'readonly'
        self.CircuitLEC3FrameEntry['state'] = 'readonly'
        self.CircuitLEC4FrameEntry['state'] = 'readonly'
        self.CircuitLEC5FrameEntry['state'] = 'readonly'
        self.CircuitDMARK1FrameEntry['state'] = 'readonly'
        self.CircuitDMARK2FrameEntry['state'] = 'readonly'
        self.CircuitCEASNFrameEntry['state'] = 'readonly'
        self.CircuitCEIPAddressFrameEntry['state'] = 'readonly'

        # COST Entries (Float)
        self.CircuitMonthlyCostFrameEntry['state'] = 'readonly'
        self.CircuitETFFrameEntry['state'] = 'readonly'


##################
        

    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]
        
        self.CircuitsID = self.CircuitIDFrameEntry.get()
        self.Description = self.CircuitsDescriptionFrameEntry.get()  

        # Calendars:
        self.CircuitsButtonInstalledDate['state'] = ACTIVE
        self.CircuitsButtonActivatedDate['state'] = ACTIVE
        self.CircuitsButtonDisconnectedDate['state'] = ACTIVE
        self.CircuitsButtonExpirationDate['state'] = ACTIVE

        self.CircuitsButtonInstalledDateClear['state'] = ACTIVE
        self.CircuitsButtonActivatedDateClear['state'] = ACTIVE
        self.CircuitsButtonExpirationDateClear['state'] = ACTIVE
        self.CircuitsButtonDisconnectedDateClear['state'] = ACTIVE

        self.CircuitsInstalledDate = self.CircuitsInstalledDateFrameEntry.get()        
        if (len(self.CircuitsInstalledDate) > 0):
            self.CircuitsInstalledDate = str(self.CircuitsInstalledData['month_selected']) + '/' + str(self.CircuitsInstalledData['day_selected']) + '/' + str(self.CircuitsInstalledData['year_selected'])
            self.CircuitsInstalledMonth = self.CircuitsInstalledData['month_selected']
            self.CircuitsInstalledDay = self.CircuitsInstalledData['day_selected']
            self.CircuitsInstalledYear = self.CircuitsInstalledData['year_selected']
        else:
            self.CircuitsInstalledDate = ""
            self.CircuitsInstalledMonth = "0"
            self.CircuitsInstalledDay = "0"
            self.CircuitsInstalledYear = "0"

        self.CircuitsActivatedDate = self.CircuitsActivatedDateFrameEntry.get()        
        if (len(self.CircuitsActivatedDate) > 0):
            self.CircuitsActivatedDate = str(self.CircuitsActivatedData['month_selected']) + '/' + str(self.CircuitsActivatedData['day_selected']) + '/' + str(self.CircuitsActivatedData['year_selected'])
            self.CircuitsActivatedMonth = self.CircuitsActivatedData['month_selected']
            self.CircuitsActivatedDay = self.CircuitsActivatedData['day_selected']
            self.CircuitsActivatedYear = self.CircuitsActivatedData['year_selected']
        else:
            self.CircuitsActivatedDate = ""
            self.CircuitsActivatedMonth = "0"
            self.CircuitsActivatedDay = "0"
            self.CircuitsActivatedYear = "0"

        self.CircuitsDisconnectedDate = self.CircuitsDisconnectedDateFrameEntry.get()        
        if (len(self.CircuitsDisconnectedDate) > 0):
            self.CircuitsDisconnectedDate = str(self.CircuitsDisconnectedData['month_selected']) + '/' + str(self.CircuitsDisconnectedData['day_selected']) + '/' + str(self.CircuitsDisconnectedData['year_selected'])
            self.CircuitsDisconnectedMonth = self.CircuitsDisconnectedData['month_selected']
            self.CircuitsDisconnectedDay = self.CircuitsDisconnectedData['day_selected']
            self.CircuitsDisconnectedYear = self.CircuitsDisconnectedData['year_selected']
        else:
            self.CircuitsDisconnectedDate = ""
            self.CircuitsDisconnectedMonth = "0"
            self.CircuitsDisconnectedDay = "0"
            self.CircuitsDisconnectedYear = "0"
            
        self.CircuitsExpirationDate = self.CircuitsExpirationDateFrameEntry.get()        
        if (len(self.CircuitsExpirationDate) > 0):
            self.CircuitsExpirationDate = str(self.CircuitsExpirationData['month_selected']) + '/' + str(self.CircuitsExpirationData['day_selected']) + '/' + str(self.CircuitsExpirationData['year_selected'])
            self.CircuitsExpirationMonth = self.CircuitsExpirationData['month_selected']
            self.CircuitsExpirationDay = self.CircuitsExpirationData['day_selected']
            self.CircuitsExpirationYear = self.CircuitsExpirationData['year_selected']
        else:
            self.CircuitsExpirationDate = ""
            self.CircuitsExpirationMonth = "0"
            self.CircuitsExpirationDay = "0"
            self.CircuitsExpirationYear = "0"

        # Comboboxes:
        self.CircuitsTypeID = self.CircuitTypeIDArray[self.CircuitComboBoxTypeID.current()] 
        self.CircuitsPortID = self.CircuitPortSpeedIDArray[self.CircuitsComboBoxPortSpeed.current()]
        self.Circuitstatus = self.CircuitstatusValues[self.CircuitsComboBoxStatus.current()]
        self.CircuitsCarrier = self.CircuitCarrierIDArray[self.CircuitComboBoxCarrier.current()]
        self.CircuitsTermTime = self.CircuitTermValues[self.CircuitComboBoxTerm.current()]

        # Setup Labels and Entry        
        self.CircuitsBandwidth = self.CircuitsBandwidthFrameEntry.get()
        self.CircuitTerm = self.CircuitTermFrameEntry.get()
        self.CircuitsContractNo = self.CircuitsContractNoFrameEntry.get()
        self.CircuitAccountNo = self.CircuitAccountNoFrameEntry.get()
        self.CircuitOrderNo = self.CircuitOrderNoFrameEntry.get()
        self.CircuitCEASN = self.CircuitCEASNFrameEntry.get()
        self.CircuitCEIPAddress = self.CircuitCEIPAddressFrameEntry.get()
        self.CircuitVLANNo = self.CircuitVLANNoFrameEntry.get()
        self.CircuitNPANXX = self.CircuitNPANXXFrameEntry.get()
        self.CircuitPEASN = self.CircuitPEASNFrameEntry.get()
        self.CircuitPEIPAddress = self.CircuitPEIPAddressFrameEntry.get()
        self.CircuitPESwitch = self.CircuitPESwitchFrameEntry.get()
        self.CircuitPELocation = self.CircuitPELocationFrameEntry.get()
        self.CircuitMonthlyCost = self.CircuitMonthlyCostFrameEntry.get()
        self.CircuitLEC1 = self.CircuitLEC1FrameEntry.get()
        self.CircuitLEC2 = self.CircuitLEC2FrameEntry.get()
        self.CircuitLEC3 = self.CircuitLEC3FrameEntry.get()
        self.CircuitLEC4 = self.CircuitLEC4FrameEntry.get()
        self.CircuitLEC5 = self.CircuitLEC5FrameEntry.get()
        self.CircuitDMARK1 = self.CircuitDMARK1FrameEntry.get()
        self.CircuitDMARK2 = self.CircuitDMARK2FrameEntry.get()
        self.CircuitCEASN = self.CircuitCEASNFrameEntry.get()
        self.CircuitCEIPAddress = self.CircuitCEIPAddressFrameEntry.get()

        # COST Entries (Float)
        if (len(self.CircuitMonthlyCostFrameEntry.get()) > 0):    
            self.CircuitMonthlyCost = float(self.CircuitMonthlyCostFrameEntry.get())
        else:
            self.CircuitMonthlyCost = 0
        
        if (len(self.CircuitETFFrameEntry.get()) > 0):    
            self.CircuitETF = float(self.CircuitETFFrameEntry.get())
        else:
            self.CircuitETF = 0

        '''
        self.CircuitsIPAddress = self.CircuitsBandwidthFrameEntry.get()            
        self.CircuitsContract = self.CircuitsContractNoFrameEntry.get()
        self.CircuitserialNo = self.CircuitCEASNFrameEntry.get()
        self.CircuitsMACAddress = self.CircuitCEIPAddressFrameEntry.get()
        if (len(self.CircuitMonthlyCostFrameEntry.get()) > 0):
            self.CircuitsOutSourceCost = float(self.CircuitMonthlyCostFrameEntry.get())
        else:
            self.CircuitsOutSourceCost = 0
        if (len(self.CircuitETFFrameEntry.get()) > 0):    
            self.CircuitETF = float(self.CircuitETFFrameEntry.get())
        else:
            self.CircuitETF = 0
        self.CircuitDMARK1 = self.CircuitDMARK1FrameEntry.get()
        '''

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <END> ------------------------*
#****************************************************************************************


#****************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************

    def Display_Country_Window(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','OPEN Country Window']    
            Logging.Log(Parameter)        
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        #print (event)
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REFRESH Country Window']    
            Logging.Log(Parameter)        
        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the Circuits Table
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Country Records found')
                self.sql_querry = False
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
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
            Parameter = ['Circuits','OPEN Region Window']    
            Logging.Log(Parameter)        
        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        self.Clean_Screen('region','region-combo')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                    ORDER BY Facility_Name ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            #print (sql)
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                #print ("found it")
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REFRESH Region Window']    
            Logging.Log(Parameter)        
        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
                      """ % (self.CountryID_Pre,self.RegionID_Pre)
            else:
                sql = """
                        SELECT * FROM Region
                        WHERE Country_ID = '%s'
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
                self.sql_querry = False
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
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
            Parameter = ['Circuits','OPEN Facility Window']    
            Logging.Log(Parameter)        
        Facility.Display_Facility_Window()

    def on_facility_combo_changed(self,event):
        self.Clean_Screen('facility','facility-combo')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REFRESH Facility Window']    
            Logging.Log(Parameter)        
        self.Clean_Screen('facility','all')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                      """ % (self.CountryID_Pre,self.RegionID_Pre,self.FacilityID_Pre)
            else:
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
                self.sql_querry = False
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
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
            Parameter = ['Circuits','OPEN Sites Window']    
            Logging.Log(Parameter)        
        Sites.Display_Sites_Window()

    def on_sites_combo_changed(self,event):
        self.Clean_Screen('sites','sites-combo')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
            sql = """
                    SELECT * FROM Circuits
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' AND Site_ID = '%s'
                    ORDER BY Circuit_ID ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                             self.FacilityIDArray[self.ComboBoxFacilityID.current()],self.SitesIDArray[self.ComboBoxSitesID.current()])
            #print (sql)
            if (Is_Sites_Available):
                self.ButtonSitesAdd['state'] = ACTIVE
                self.ButtonSitesRefresh['state'] = ACTIVE
            if ((self.db.Execute(sql)) and (self.Selection != 'edit')):
                #------------------------------- Deleting Tree View <BEGIN> --------
                #x = self.CircuitsTreeview.get_children()
                #if x != '()': # checks if there is something in the first row
                #    for child in x:
                #       #print (child)
                #        self.CircuitsTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                #-------------- Initializing Arrays <BEGIN> ----------------------
                self.CircuitsTablePriaryKeyArray =  [] # Circuits ID
                self.CircuitsTableDescriptionArray =  []                
                self.CircuitsTableCountryIDArray =  [] 
                self.CircuitsTableRegionIDArray =  []
                self.CircuitsTableFacilityIDArray =  []
                self.CircuitsTableSiteIDArray =  []
                self.CircuitsTableCarrierIDArray =  []
                self.CircuitsTableCircuitTypeArray = []
                self.CircuitsTablePortSpeedArray = []
                self.CircuitsTableBandwidthArray = []
                self.CircuitsTableStatusArray = []
                self.CircuitsTableDmarc_Info_1Array = []
                self.CircuitsTableDmarc_Info_2Array = []
                self.CircuitsTableLEC1Array = []
                self.CircuitsTableLEC2Array = []
                self.CircuitsTableLEC3Array = []
                self.CircuitsTableLEC4Array = []
                self.CircuitsTableLEC5Array = []
                self.CircuitsTableCE_ASNArray = []
                self.CircuitsTableCE_IP_AddressArray = []
                self.CircuitsTablePE_ASNArray = []
                self.CircuitsTablePE_IP_AddressArray = []
                self.CircuitsTableVLAN_IDArray = []
                self.CircuitsTablePE_SwitchArray = []
                self.CircuitsTablePE_LocationArray = []
                self.CircuitsTableNPA_NXXArray = []
                self.CircuitsTableMonthlyCostArray = []
                self.CircuitsTableOrderNumberArray = []
                self.CircuitsTableDateInstalledArray = []
                self.CircuitsTableDayInstalledArray = []
                self.CircuitsTableMonthInstalledArray = []
                self.CircuitsTableYearInstalledArray = []
                self.CircuitsTableDateActivatedArray = []
                self.CircuitsTableDayActivatedArray = []
                self.CircuitsTableMonthActivatedArray = []
                self.CircuitsTableYearActivatedArray = []
                self.CircuitsTableDisconectedDateArray = []
                self.CircuitsTableDayDisconectedArray = []
                self.CircuitsTableMonthDisconectedArray = []
                self.CircuitsTableYearDisconectedArray = []
                self.CircuitsTableExpirationDateArray = []
                self.CircuitsTableDayExpirationArray = []
                self.CircuitsTableMonthExpirationArray = []
                self.CircuitsTableYearExpirationArray = []
                self.CircuitsTableTerm_DayArray = []
                self.CircuitsTableTerm_TimeArray = []
                self.CircuitsTableETFArray = []
                self.CircuitsTableContract_NoArray = []
                self.CircuitsTableAccount_NoArray = []
                self.CircuitsTableExecutedByArray = []
                self.results = []
                self.CircuitsTableArrayTemp = []
                self.circuitsTableColumns = (
                    'Circuit ID',
                    'Description',
                    'Country ID',                    
                    'Region ID',
                    'Facility ID',
                    'Site ID',
                    'Carrier ID',
                    'Circuit Type',
                    'Port Speed',
                    'Bandwidth',
                    'Status',
                    'Dmarc Info 1',
                    'Dmarc Info 2',
                    'LEC1',
                    'LEC2',
                    'LEC3',
                    'LEC4',
                    'LEC5',
                    'CE ASN',
                    'CE IP Address',
                    'PE ASN',
                    'PE IP Address',
                    'VLAN ID',
                    'PE Switch',
                    'PE Location',
                    'NPA NXX',
                    'Monthly Cost',
                    'Order Number',
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
                    'Term Day',
                    'Term Time',
                    'ETF',
                    'Contract No',
                    'Account No',
                    'Executed by UserID')
                
                '''
                    0 - Circuit_ID              CHAR(100) NOT NULL PRIMARY KEY,
                    1 - Description             CHAR(200),
                    2 - Country_ID              CHAR(20),                    
                    3 - Region_ID               CHAR(20),
                    4 - Facility_ID             CHAR(20),
                    5 - Site_ID                 CHAR(20),
                    6 - Carrier_ID              CHAR(20),
                    7 - Circuit_Type            CHAR(40),
                    8 - Port_Speed              CHAR(20),
                    9 - Bandwidth               CHAR(20),
                    10- Status                  CHAR(20),
                    11- Dmarc_Info_1            CHAR(200),
                    12- Dmarc_Info_2            CHAR(200),
                    13- LEC1                    CHAR(50),
                    14- LEC2                    CHAR(50),
                    15- LEC3                    CHAR(50),
                    16- LEC4                    CHAR(50),
                    17- LEC5                    CHAR(50),
                    18- CE_ASN                  CHAR(20),
                    19- CE_IP_Address           CHAR(50),
                    20- PE_ASN                  CHAR(20),
                    21- PE_IP_Address           CHAR(50),
                    22- VLAN_ID                 CHAR(10),
                    23- PE_Switch               CHAR(100),
                    24- PE_Location             CHAR(100),
                    25- NPA_NXX                 CHAR(20),
                    26- Monthly_Cost            FLOAT,
                    27- Order_Number            CHAR(40),
                    28- Date_Installed          CHAR(20),
                    29- Day_Installed           INT,
                    30- Month_Installed         INT,
                    31- Year_Installed          INT,
                    32- Date_Activated          CHAR(20),
                    33- Day_Activated           INT,
                    34- Month_Activated         INT,
                    35- Year_Activated          INT,
                    36- Disconnect_Date         CHAR(20),
                    37- Day_Disconnect          INT,
                    38- Month_Disconnect        INT,
                    39- Year_Disconnect         INT,
                    40- Expiration_Date         CHAR(20),
                    41- Day_Expiration          INT,
                    42- Month_Expiration        INT,
                    43- Year_Expiration         INT,
                    44- Term_Day                CHAR(10),
                    45- Term_Time               CHAR(10),
                    46- ETF                     FLOAT,
                    47- Contract_No             CHAR(50),
                    48- Account_No              CAHR(50),
                    49- Executed_by_UserID      CHAR(20))"""

                '''
                #-------------- Initializing Arrays <END> ----------------------
                self.data_ready = True
                i = 0
                while (i < len(self.db.results)):                    
                    self.CircuitsTablePriaryKeyArray.append(self.db.results[i][0].strip())
                    self.CircuitsTableDescriptionArray.append(self.db.results[i][1].strip())
                    self.CircuitsTableCountryIDArray.append(self.db.results[i][2].strip())
                    self.CircuitsTableRegionIDArray.append(self.db.results[i][3].strip())
                    self.CircuitsTableFacilityIDArray.append(self.db.results[i][4].strip())
                    self.CircuitsTableSiteIDArray.append(self.db.results[i][5].strip())
                    self.CircuitsTableCarrierIDArray.append(self.db.results[i][6].strip())
                    self.CircuitsTableCircuitTypeArray.append(self.db.results[i][7].strip())
                    self.CircuitsTablePortSpeedArray.append(self.db.results[i][8].strip())
                    self.CircuitsTableBandwidthArray.append(self.db.results[i][9].strip())
                    self.CircuitsTableStatusArray.append(self.db.results[i][10].strip())
                    if ((self.db.results[i][11]) == None):
                        self.CircuitsTableDmarc_Info_1Array.append("")
                    else:
                        self.CircuitsTableDmarc_Info_1Array.append(self.db.results[i][11].strip())
                    if ((self.db.results[i][12]) == None):
                        self.CircuitsTableDmarc_Info_2Array.append("")
                    else:
                        self.CircuitsTableDmarc_Info_2Array.append(self.db.results[i][12].strip())
                    if ((self.db.results[i][13]) == None):
                        self.CircuitsTableLEC1Array.append("")
                    else:
                        self.CircuitsTableLEC1Array.append(self.db.results[i][13].strip())
                    if ((self.db.results[i][14]) == None):    
                        self.CircuitsTableLEC2Array.append("")
                    else:
                        self.CircuitsTableLEC2Array.append(self.db.results[i][14].strip())

                    if ((self.db.results[i][15]) == None):
                        self.CircuitsTableLEC3Array.append("")
                    else:
                        self.CircuitsTableLEC3Array.append(self.db.results[i][15].strip())
                    if ((self.db.results[i][16]) == None):
                        self.CircuitsTableLEC4Array.append("")
                    else:
                        self.CircuitsTableLEC4Array.append(self.db.results[i][16].strip())
                    if ((self.db.results[i][17]) == None):
                        self.CircuitsTableLEC5Array.append("")
                    else:
                        self.CircuitsTableLEC5Array.append(self.db.results[i][17].strip())
                    if ((self.db.results[i][18]) == None):
                        self.CircuitsTableCE_ASNArray.append("")
                    else:
                        self.CircuitsTableCE_ASNArray.append(self.db.results[i][18].strip())
                    if ((self.db.results[i][19]) == None):
                        self.CircuitsTableCE_IP_AddressArray.append("0.0.0.0")
                    else:
                        self.CircuitsTableCE_IP_AddressArray.append(self.db.results[i][19].strip())
                    if ((self.db.results[i][20]) == None):
                        self.CircuitsTablePE_ASNArray.append("")
                    else:
                        self.CircuitsTablePE_ASNArray.append(self.db.results[i][20].strip())
                    if ((self.db.results[i][21]) == None):
                        self.CircuitsTablePE_IP_AddressArray.append("0.0.0.0")
                    else:
                        self.CircuitsTablePE_IP_AddressArray.append(self.db.results[i][21].strip())
                    if ((self.db.results[i][22]) == None):
                        self.CircuitsTableVLAN_IDArray.append("")
                    else:
                        self.CircuitsTableVLAN_IDArray.append(self.db.results[i][22].strip())
                    if ((self.db.results[i][23]) == None):
                        self.CircuitsTablePE_SwitchArray.append("")
                    else:
                        self.CircuitsTablePE_SwitchArray.append(self.db.results[i][23].strip())
                    if ((self.db.results[i][24]) == None):
                        self.CircuitsTablePE_LocationArray.append("")
                    else:
                        self.CircuitsTablePE_LocationArray.append(self.db.results[i][24].strip())
                    if (self.db.results[i][25] == None):
                        self.CircuitsTableNPA_NXXArray.append("")
                    else:
                        self.CircuitsTableNPA_NXXArray.append(self.db.results[i][25].strip())
                    if (self.db.results[i][26] == None):
                        self.CircuitsTableMonthlyCostArray.append(0.0)
                    else:
                        self.CircuitsTableMonthlyCostArray.append(self.db.results[i][26])
                    if (self.db.results[i][27] == None):
                        self.CircuitsTableOrderNumberArray.append("")
                    else:
                        self.CircuitsTableOrderNumberArray.append(self.db.results[i][27].strip())
                    if (((self.db.results[i][28]) == None) or (self.db.results[i][29] == 0)):
                        self.CircuitsTableDateInstalledArray.append("")
                        self.CircuitsTableDayInstalledArray.append(0)
                        self.CircuitsTableMonthInstalledArray.append(0)
                        self.CircuitsTableYearInstalledArray.append(0)
                    else:
                        self.CircuitsTableDateInstalledArray.append(self.db.results[i][28].strip())
                        self.CircuitsTableDayInstalledArray.append(self.db.results[i][29])
                        self.CircuitsTableMonthInstalledArray.append(self.db.results[i][30])
                        self.CircuitsTableYearInstalledArray.append(self.db.results[i][31])

                    if (((self.db.results[i][32]) == None) or (self.db.results[i][33] == 0)):
                        self.CircuitsTableDateActivatedArray.append("")
                        self.CircuitsTableDayActivatedArray.append(0)
                        self.CircuitsTableMonthActivatedArray.append(0)
                        self.CircuitsTableYearActivatedArray.append(0)
                    else:
                        self.CircuitsTableDateActivatedArray.append(self.db.results[i][32].strip())
                        self.CircuitsTableDayActivatedArray.append(self.db.results[i][33])
                        self.CircuitsTableMonthActivatedArray.append(self.db.results[i][34])
                        self.CircuitsTableYearActivatedArray.append(self.db.results[i][35])

                    if (((self.db.results[i][36]) == None) or (self.db.results[i][37] == 0)):
                        self.CircuitsTableDisconectedDateArray.append("")
                        self.CircuitsTableDayDisconectedArray.append(0)
                        self.CircuitsTableMonthDisconectedArray.append(0)
                        self.CircuitsTableYearDisconectedArray.append(0)
                    else:
                        self.CircuitsTableDisconectedDateArray.append(self.db.results[i][36].strip())
                        self.CircuitsTableDayDisconectedArray.append(self.db.results[i][37])
                        self.CircuitsTableMonthDisconectedArray.append(self.db.results[i][38])
                        self.CircuitsTableYearDisconectedArray.append(self.db.results[i][39])

                    if (((self.db.results[i][40]) == None) or (self.db.results[i][41] == 0)):
                        self.CircuitsTableExpirationDateArray.append("")
                        self.CircuitsTableDayExpirationArray.append(0)
                        self.CircuitsTableMonthExpirationArray.append(0)
                        self.CircuitsTableYearExpirationArray.append(0)
                    else:
                        self.CircuitsTableExpirationDateArray.append(self.db.results[i][40].strip())
                        self.CircuitsTableDayExpirationArray.append(self.db.results[i][41])
                        self.CircuitsTableMonthExpirationArray.append(self.db.results[i][42])
                        self.CircuitsTableYearExpirationArray.append(self.db.results[i][43])
                    if (self.db.results[i][44] == None):
                        self.CircuitsTableTerm_DayArray.append("")
                    else:
                        self.CircuitsTableTerm_DayArray.append(self.db.results[i][44].strip())
                    if (self.db.results[i][45] == None):
                        self.CircuitsTableTerm_TimeArray.append("")
                    else:
                        self.CircuitsTableTerm_TimeArray.append(self.db.results[i][45].strip())
                    if (self.db.results[i][46] == None):
                        self.CircuitsTableETFArray.append(0.0)
                    else:
                        self.CircuitsTableETFArray.append(self.db.results[i][46])
                    if (self.db.results[i][47] == None):
                        self.CircuitsTableContract_NoArray.append("")
                    else:
                        self.CircuitsTableContract_NoArray.append(self.db.results[i][47].strip())
                    if (self.db.results[i][48] == None):
                        self.CircuitsTableAccount_NoArray.append("")
                    else:
                        self.CircuitsTableAccount_NoArray.append(self.db.results[i][48].strip())
                    if (self.db.results[i][49] == None):
                        self.CircuitsTableExecutedByArray.append("")
                    else:
                        self.CircuitsTableExecutedByArray.append(self.db.results[i][49].strip())
                    i = i + 1
                i = 0
                while (i < len(self.CircuitsTablePriaryKeyArray)):
                    num = i + 1
                    tags = self.CircuitsTableStatusArray[i] # To use in the futire
                    item = [
                            self.CircuitsTablePriaryKeyArray[i],
                            self.CircuitsTableDescriptionArray[i],
                            self.CircuitsTableCarrierIDArray[i],
                            self.CircuitsTableCircuitTypeArray[i],
                            self.CircuitsTablePortSpeedArray[i],
                            self.CircuitsTableBandwidthArray[i],
                            #self.CircuitsTableCE_ASNArray[i],
                            #self.CircuitsTableCE_IP_AddressArray[i],
                            #self.CircuitsTablePE_ASNArray[i],
                            #self.CircuitsTablePE_IP_AddressArray[i],
                            #self.CircuitsTableVLAN_IDArray[i],
                            self.CircuitsTableMonthlyCostArray[i],
                            self.CircuitsTableETFArray[i],                            
                            self.CircuitsTableDateInstalledArray[i],
                            self.CircuitsTableDateActivatedArray[i],
                            self.CircuitsTableDisconectedDateArray[i],
                            self.CircuitsTableStatusArray[i],
                            self.CircuitsTableContract_NoArray[i],
                            self.CircuitsTableExpirationDateArray[i],
                            self.CircuitsTableExecutedByArray[i]
                            ]
                    self.CircuitsTableArrayTemp = [
                    self.CircuitsTablePriaryKeyArray[i],
                    self.CircuitsTableDescriptionArray[i],
                    self.CircuitsTableCountryIDArray[i],
                    self.CircuitsTableRegionIDArray[i],
                    self.CircuitsTableFacilityIDArray[i],
                    self.CircuitsTableSiteIDArray[i],
                    self.CircuitsTableCarrierIDArray[i],
                    self.CircuitsTableCircuitTypeArray[i],
                    self.CircuitsTablePortSpeedArray[i],
                    self.CircuitsTableBandwidthArray[i],
                    self.CircuitsTableStatusArray[i],
                    self.CircuitsTableDmarc_Info_1Array[i],
                    self.CircuitsTableDmarc_Info_2Array[i],
                    self.CircuitsTableLEC1Array[i],
                    self.CircuitsTableLEC2Array[i],
                    self.CircuitsTableLEC3Array[i],
                    self.CircuitsTableLEC4Array[i],
                    self.CircuitsTableLEC5Array[i],
                    self.CircuitsTableCE_ASNArray[i],
                    self.CircuitsTableCE_IP_AddressArray[i],
                    self.CircuitsTablePE_ASNArray[i],
                    self.CircuitsTablePE_IP_AddressArray[i],
                    self.CircuitsTableVLAN_IDArray[i],
                    self.CircuitsTablePE_SwitchArray[i],
                    self.CircuitsTablePE_LocationArray[i],
                    self.CircuitsTableNPA_NXXArray[i],
                    self.CircuitsTableMonthlyCostArray[i],
                    self.CircuitsTableOrderNumberArray[i],
                    self.CircuitsTableDateInstalledArray[i],
                    self.CircuitsTableDayInstalledArray[i],
                    self.CircuitsTableMonthInstalledArray[i],
                    self.CircuitsTableYearInstalledArray[i],
                    self.CircuitsTableDateActivatedArray[i],
                    self.CircuitsTableDayActivatedArray[i],
                    self.CircuitsTableMonthActivatedArray[i],
                    self.CircuitsTableYearActivatedArray[i],
                    self.CircuitsTableDisconectedDateArray[i],
                    self.CircuitsTableDayDisconectedArray[i],
                    self.CircuitsTableMonthDisconectedArray[i],
                    self.CircuitsTableYearDisconectedArray[i],
                    self.CircuitsTableExpirationDateArray[i],
                    self.CircuitsTableDayExpirationArray[i],
                    self.CircuitsTableMonthExpirationArray[i],
                    self.CircuitsTableYearExpirationArray[i],
                    self.CircuitsTableTerm_DayArray[i],
                    self.CircuitsTableTerm_TimeArray[i],
                    self.CircuitsTableETFArray[i],
                    self.CircuitsTableContract_NoArray[i],
                    self.CircuitsTableAccount_NoArray[i],
                    self.CircuitsTableExecutedByArray[i]
                    ]
                    self.results.append(self.CircuitsTableArrayTemp)
                    self.CircuitsTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonCircuitsAdd['state']    = ACTIVE
                self.ButtonCircuitsEdit['state']   = DISABLED
                self.ButtonCircuitsRemove['state'] = DISABLED
                self.ButtonCircuitsOK['state']     = DISABLED
                self.ButtonCircuitsCancel['state'] = DISABLED
                self.ButtonDevicePingPE64['state'] = DISABLED
                self.ButtonDevicePingPE1500['state'] = DISABLED
                self.ButtonDevicePingCE64['state'] = DISABLED
                self.ButtonDevicePingCE1500['state'] = DISABLED
                self.ButtonDeviceCircuits['state'] = DISABLED
                self.ButtonDeviceTraceroutePE['state'] = DISABLED
                self.ButtonDeviceTracerouteCE['state'] = DISABLED
                self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED

            else:
                if (self.Selection != 'edit'):
                    mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                    message = 'No Records found')
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Sites_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REFRESH Sites Window']    
            Logging.Log(Parameter)        
        self.Clean_Screen('sites','all')
        if self.db.Connect():
            # SQL Querry to the Circuits Table
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
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found')
                self.ComboBoxSitesID['state'] = DISABLED
                self.sql_querry = False
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- SITES SELECTION SECTION <END> ------------------------*
#*************************************************************************************
        
    def on_CircuitsWindow_quit(self):
        if (self.CircuitsWindowExist):
            self.CircuitsWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Circuits','CLOSE Window']    
                Logging.Log(Parameter)        
            self.CircuitsWindow.destroy()

    def on_Circuits_Table_Refresh(self): # I need to do more research on this call.
        self.on_country_combo_changed("event")
       
    def Call_Button_Circuits_Add(self):
        #-- reset the progess bar --
        self.Enable_Screen('add')
        self.Selection = 'add'
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','ADD Button']    
            Logging.Log(Parameter)        

    def Call_Button_Circuits_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','EDIT Button']    
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
        self.CircuitsInstalledData = {}
        if (self.Selection == 'edit'):
            self.CircuitsInstalledDateName = self.CircuitsInstalledDateFrameEntry.get()
            curItem = self.CircuitsTreeview.focus()  
            dic = self.CircuitsTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.CircuitsInstalledDateName) > 0):
                if (self.CircuitsTableDateInstalledArray[curItem] != 0):
                    self.CircuitsInstalledData['day_selected'] = self.CircuitsTableDayInstalledArray[curItem]
                    self.CircuitsInstalledData['month_selected'] = self.CircuitsTableMonthInstalledArray[curItem]
                    self.CircuitsInstalledData['year_selected'] = self.CircuitsTableYearInstalledArray[curItem]
        # ----- Activated Date ---------------------
        self.CircuitsActivatedDateName = self.CircuitsActivatedDateFrameEntry.get()
        curItem = self.CircuitsTreeview.focus()  
        dic = self.CircuitsTreeview.item(curItem)
        curItem = int(dic.get('text')) - 1
        self.CircuitsActivatedData = {}
        if (len(self.CircuitsActivatedDateName) > 0):
            if (self.CircuitsTableDateActivatedArray[curItem] != 0):
                self.CircuitsActivatedData['day_selected'] = self.CircuitsTableDayActivatedArray[curItem]
                self.CircuitsActivatedData['month_selected'] = self.CircuitsTableMonthActivatedArray[curItem]
                self.CircuitsActivatedData['year_selected'] = self.CircuitsTableYearActivatedArray[curItem]
                #print ("Day, Month, Year")
                #print (self.CircuitsActivatedData['day_selected'])
                #print (self.CircuitsActivatedData['month_selected'])
                #print (self.CircuitsActivatedData['year_selected'])
        # ----- Disconnected Date ---------------------
        self.CircuitsDisconnectedData = {}
        if (self.Selection == 'edit'):
            self.CircuitsDisconnectedDateName = self.CircuitsDisconnectedDateFrameEntry.get()
            curItem = self.CircuitsTreeview.focus()  
            dic = self.CircuitsTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.CircuitsDisconnectedDateName) > 0):
                if (self.CircuitsTableDisconectedDateArray[curItem] != 0):
                    self.CircuitsDisconnectedData['day_selected'] = self.CircuitsTableDayDisconectedArray[curItem]
                    self.CircuitsDisconnectedData['month_selected'] = self.CircuitsTableMonthDisconectedArray[curItem]
                    self.CircuitsDisconnectedData['year_selected'] = self.CircuitsTableYearDisconectedArray[curItem]
        # ----- Expiration Date ---------------------
        self.CircuitsExpirationData = {}
        if (self.Selection == 'edit'):
            self.CircuitsExpirationDateName = self.CircuitsExpirationDateFrameEntry.get()
            curItem = self.CircuitsTreeview.focus()  
            dic = self.CircuitsTreeview.item(curItem)
            curItem = int(dic.get('text')) - 1
            if (len(self.CircuitsExpirationDateName) > 0):
                if (self.CircuitsTableExpirationDateArray[curItem] != 0):
                    self.CircuitsExpirationData['day_selected'] = self.CircuitsTableDayExpirationArray[curItem]
                    self.CircuitsExpirationData['month_selected'] = self.CircuitsTableMonthExpirationArray[curItem]
                    self.CircuitsExpirationData['year_selected'] = self.CircuitsTableYearExpirationArray[curItem]

    #-------------- Using a Password Question to make sure it was the intent to be deleted <Begin> ---------------

    def Remove_Circuits_From_DB(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REMOVE Final Button']    
            Logging.Log(Parameter)        
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            #self.CircuitsID = self.CircuitIDFrameEntry.get()
            PrimaryKey = (self.CircuitsID)
            if (mbox.askyesnocancel(master=self.CircuitsFrame,title='Circuits',message = 'Are you Sure you want to Remove it?')):
                #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.CircuitsID)
                #print (PrimaryKey)
                sql = """
                        SELECT * FROM Circuits
                        WHERE Circuit_ID = '%s'
                       """ % (PrimaryKey)                
                if (self.db.Execute(sql)):
                    sql = "DELETE FROM Circuits WHERE Circuit_ID = '%s'" % (PrimaryKey)
                    if (self.db.Add_Move_Change_Data(sql)):
                        #self.db.Disconnect()
                        mbox.showwarning(master=self.CircuitsFrame,title='Circuits',
                        message = '*** The Circuit ID you entered was Removed ***')
                    else:
                        #self.db.Disconnect()
                        mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                        message = '*** The Circuit ID you entered was NOT Removed ***')               
                    self.on_sites_combo_changed("event")
                    self.Disable_Screen()
                else:
                    mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                    message = '*** The Circuit ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.GetPasswordWindow.destroy()
            self.Remove_Circuits_From_DB()
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

    def Call_Button_Circuits_Remove(self):
        self.CircuitsID = self.CircuitIDFrameEntry.get()
        self.Get_Usernanme_and_Password()
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','REMOVE Button']    
            Logging.Log(Parameter)        

    def Call_Button_Circuits_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','OK/UPDATE Button']    
            Logging.Log(Parameter)        
        if self.db.Connect():
            self.ButtonCircuitsAdd['state']    = DISABLED
            self.ButtonCircuitsEdit['state']   = DISABLED
            self.ButtonCircuitsRemove['state'] = DISABLED
            self.ButtonCircuitsOK['state']     = ACTIVE
            self.ButtonCircuitsCancel['state'] = ACTIVE
            self.ButtonDevicePingPE64['state'] = DISABLED
            self.ButtonDevicePingPE1500['state'] = DISABLED
            self.ButtonDevicePingCE64['state'] = DISABLED
            self.ButtonDevicePingCE1500['state'] = DISABLED
            self.ButtonDeviceCircuits['state'] = DISABLED
            self.ButtonDeviceTraceroutePE['state'] = DISABLED
            self.ButtonDeviceTracerouteCE['state'] = DISABLED
            self.ButtonDeviceLocalPointOfContacts['state'] = DISABLED
            
            self.Collect_Screen() # <-------------------- Collect all Data on Screen
            PrimaryKey = (self.CircuitsID)
            #self.Selection = 'notyet'
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(self.CircuitsID) > 0) and (len(self.Description) > 0) and (len(self.CircuitsBandwidth) > 0)):
                    if ((self.IPFormatCheck(self.CircuitCEIPAddress)) and (self.IPFormatCheck(self.CircuitPEIPAddress))):
                        sql = """
                                SELECT * FROM Circuits
                                WHERE Circuit_ID = '%s'
                              """ % (PrimaryKey)
                        if (self.db.Execute(sql)):
                            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                            message = '*** The Circuits ID you entered already exist ***')
                        else:
                            if (len(self.Description) > 200):
                                self.Description = self.Description[:200]
                            sql = "INSERT INTO Circuits(Circuit_ID, Description, Country_ID, Region_ID, Facility_ID, Site_ID, Carrier_ID, Circuit_Type, \
                                       Port_Speed, Bandwidth, Status, Dmarc_Info_1, Dmarc_Info_2, LEC1, LEC2, LEC3, LEC4, LEC5, CE_ASN, CE_IP_Address, \
                                       PE_ASN, PE_IP_Address, VLAN_ID, PE_Switch, PE_Location, NPA_NXX, Monthly_Cost, Order_Number, \
                                       Date_Installed, Day_Installed, Month_Installed, Year_Installed, \
                                       Date_Activated, Day_Activated, Month_Activated, Year_Activated, \
                                       Disconnect_Date, Day_Disconnect, Month_Disconnect, Year_Disconnect, \
                                       Expiration_Date, Day_Expiration, Month_Expiration, Year_Expiration, Term_Day, Term_Time, ETF, \
                                       Contract_No, Account_No, Executed_by_UserID)\
                                       VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s', \
                                       '%s','%s','%s','%s','%s','%s','%f','%s','%s','%d','%d','%d','%s','%d','%d','%d','%s','%d','%d','%d','%s','%d','%d','%d',\
                                       '%s','%s','%f','%s','%s','%s')" % (PrimaryKey, self.Description, self.CountryID, self.RegionID, self.FacilityID, self.SitesID,
                                        self.CircuitsCarrier, self.CircuitsTypeID, self.CircuitsPortID, self.CircuitsBandwidth, self.Circuitstatus, self.CircuitDMARK1,
                                        self.CircuitDMARK2, self.CircuitLEC1, self.CircuitLEC2, self.CircuitLEC3, self.CircuitLEC4, self.CircuitLEC5, self.CircuitCEASN,
                                        self.CircuitCEIPAddress, self.CircuitPEASN, self.CircuitPEIPAddress, self.CircuitVLANNo, self.CircuitPESwitch, self.CircuitPELocation,
                                        self.CircuitNPANXX, self.CircuitMonthlyCost, self.CircuitOrderNo,
                                        self.CircuitsInstalledDate, int(self.CircuitsInstalledDay), int(self.CircuitsInstalledMonth), int(self.CircuitsInstalledYear),
                                        self.CircuitsActivatedDate, int(self.CircuitsActivatedDay), int(self.CircuitsActivatedMonth), int(self.CircuitsActivatedYear),
                                        self.CircuitsDisconnectedDate, int(self.CircuitsDisconnectedDay), int(self.CircuitsDisconnectedMonth), int(self.CircuitsDisconnectedYear),
                                        self.CircuitsExpirationDate, int(self.CircuitsExpirationDay), int(self.CircuitsExpirationMonth), int(self.CircuitsExpirationYear),
                                        self.CircuitTerm, self.CircuitsTermTime, self.CircuitETF, self.CircuitsContractNo, self.CircuitAccountNo, self.Username)

                            '''

                            self.CircuitsTypeID = self.CircuitTypeIDArray[self.CircuitComboBoxTypeID.current()] 
                            self.CircuitsPortID = self.CircuitPortSpeedIDArray[self.CircuitsComboBoxPortSpeed.current()]
                            self.Circuitstatus = self.CircuitstatusValues[self.CircuitsComboBoxStatus.current()]
                            self.CircuitsCarrier = self.CircuitCarrierIDArray[self.CircuitComboBoxCarrier.current()]
                            self.CircuitsTermTime = self.CircuitTermValues[self.CircuitComboBoxTerm.current()]

                            # Setup Labels and Entry        
                            self.CircuitsBandwidth = self.CircuitsBandwidthFrameEntry.get()
                            self.CircuitTerm = self.CircuitTermFrameEntry.get()
                            self.CircuitsContractNo = self.CircuitsContractNoFrameEntry.get()
                            self.CircuitAccountNo = self.CircuitAccountNoFrameEntry.get()
                            self.CircuitOrderNo = self.CircuitOrderNoFrameEntry.get()
                            self.CircuitCEASN = self.CircuitCEASNFrameEntry.get()
                            self.CircuitCEIPAddress = self.CircuitCEIPAddressFrameEntry.get()
                            self.CircuitVLANNo = self.CircuitVLANNoFrameEntry.get()
                            self.CircuitNPANXX = self.CircuitNPANXXFrameEntry.get()
                            self.CircuitPEASN = self.CircuitPEASNFrameEntry.get()
                            self.CircuitPEIPAddress = self.CircuitPEIPAddressFrameEntry.get()
                            self.CircuitPESwitch = self.CircuitPESwitchFrameEntry.get()
                            self.CircuitPELocation = self.CircuitPELocationFrameEntry.get()
                            self.CircuitMonthlyCost = self.CircuitMonthlyCostFrameEntry.get()
                            self.CircuitLEC1 = self.CircuitLEC1FrameEntry.get()
                            self.CircuitLEC2 = self.CircuitLEC2FrameEntry.get()
                            self.CircuitLEC3 = self.CircuitLEC3FrameEntry.get()
                            self.CircuitLEC4 = self.CircuitLEC4FrameEntry.get()
                            self.CircuitLEC5 = self.CircuitLEC5FrameEntry.get()
                            self.CircuitDMARK1 = self.CircuitDMARK1FrameEntry.get()
                            self.CircuitDMARK2 = self.CircuitDMARK2FrameEntry.get()
                            self.CircuitCEASN = self.CircuitCEASNFrameEntry.get()
                            self.CircuitCEIPAddress = self.CircuitCEIPAddressFrameEntry.get()

                            # COST Entries (Float)
                            if (len(self.CircuitMonthlyCostFrameEntry.get()) > 0):    
                                self.CircuitMonthlyCost = float(self.CircuitMonthlyCostFrameEntry.get())
                            else:
                                self.CircuitMonthlyCost = 0
                            
                            if (len(self.CircuitETFFrameEntry.get()) > 0):    
                                self.CircuitETF = float(self.CircuitETFFrameEntry.get())
                            else:
                                self.CircuitETF = 0


                            '''

                            '''
                                0 - Circuit_ID              CHAR(100) NOT NULL PRIMARY KEY,
                                1 - Description             CHAR(200),
                                2 - Country_ID              CHAR(20),                    
                                3 - Region_ID               CHAR(20),
                                4 - Facility_ID             CHAR(20),
                                5 - Site_ID                 CHAR(20),
                                6 - Carrier_ID              CHAR(20),
                                7 - Circuit_Type            CHAR(40),
                                8 - Port_Speed              CHAR(20),
                                9 - Bandwidth               CHAR(20),
                                10- Status                  CHAR(20),
                                11- Dmarc_Info_1            CHAR(200),
                                12- Dmarc_Info_2            CHAR(200),
                                13- LEC1                    CHAR(50),
                                14- LEC2                    CHAR(50),
                                15- LEC3                    CHAR(50),
                                16- LEC4                    CHAR(50),
                                17- LEC5                    CHAR(50),
                                18- CE_ASN                  CHAR(20),
                                19- CE_IP_Address           CHAR(50),
                                20- PE_ASN                  CHAR(20),
                                21- PE_IP_Address           CHAR(50),
                                22- VLAN_ID                 CHAR(10),
                                23- PE_Switch               CHAR(100),
                                24- PE_Location             CHAR(100),
                                25- NPA_NXX                 CHAR(20),
                                26- Monthly_Cost            FLOAT,
                                27- Order_Number            CHAR(40),
                                28- Date_Installed          CHAR(20),
                                29- Day_Installed           INT,
                                30- Month_Installed         INT,
                                31- Year_Installed          INT,
                                32- Date_Activated          CHAR(20),
                                33- Day_Activated           INT,
                                34- Month_Activated         INT,
                                35- Year_Activated          INT,
                                36- Disconnect_Date         CHAR(20),
                                37- Day_Disconnect          INT,
                                38- Month_Disconnect        INT,
                                39- Year_Disconnect         INT,
                                40- Expiration_Date         CHAR(20),
                                41- Day_Expiration          INT,
                                42- Month_Expiration        INT,
                                43- Year_Expiration         INT,
                                44- Term_Day                CHAR(10),
                                45- Term_Time               CHAR(10),
                                46- ETF                     FLOAT,
                                47- Contract_No             CHAR(50),
                                48- Account_No              CAHR(50),
                                49- Executed_by_UserID      CHAR(20))"""

                            '''

                            #print (sql)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.CircuitsFrame,title='Circuits',
                                    message = '*** The Circuits ID you entered was Added ***')
                                self.on_sites_combo_changed("event")
                                self.Disable_Screen()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                                    message = '*** The Circuits ID you entered was NOT Added ***')
                                self.on_sites_combo_changed("event")
                                self.Disable_Screen()
                    else:
                     mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                        message = '*** The CE or PE IP Address is invalid ***')                         
                else:
                     mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                        message = '*** The Circuits ID, Circuits Name and Bandwidth Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                #print (PrimaryKey+"....")
                if ((len(self.Description) == 0) or (len(self.CircuitsBandwidth) == 0)):
                    mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                    message = '*** The Circuits Name or The Bandwidtth Cannot be BLANK ***')
                else:
                    if ((self.IPFormatCheck(self.CircuitCEIPAddress)) and (self.IPFormatCheck(self.CircuitPEIPAddress))):
                        sql = """
                                SELECT * FROM Circuits
                                WHERE Circuit_ID = '%s'
                              """ % (PrimaryKey)                   
                        if (self.db.Execute(sql)):                            
                            if (len(self.Description) > 200):
                                self.Description = self.Description[:200]
                            sql = "UPDATE Circuits SET Description = '%s', Country_ID = '%s', Region_ID = '%s', Facility_ID = '%s', Site_ID = '%s', \
                                       Carrier_ID = '%s', Circuit_Type = '%s', Port_Speed = '%s', Bandwidth = '%s', Status = '%s', Dmarc_Info_1 = '%s', \
                                       Dmarc_Info_2 = '%s', LEC1 = '%s', LEC2 = '%s', LEC3 = '%s', LEC4 = '%s', LEC5 = '%s', CE_ASN = '%s', CE_IP_Address = '%s', \
                                       PE_ASN = '%s', PE_IP_Address = '%s', VLAN_ID = '%s', PE_Switch = '%s', PE_Location = '%s', NPA_NXX = '%s', Monthly_Cost = '%f', \
                                       Order_Number  = '%s', \
                                       Date_Installed = '%s', Day_Installed = '%d', Month_Installed = '%d', Year_Installed = '%d', \
                                       Date_Activated = '%s', Day_Activated = '%d', Month_Activated = '%d', Year_Activated = '%d', \
                                       Disconnect_Date = '%s', Day_Disconnect = '%d', Month_Disconnect = '%d', Year_Disconnect = '%d', \
                                       Expiration_Date = '%s', Day_Expiration = '%d', Month_Expiration = '%d', Year_Expiration = '%d', Term_Day = '%s', \
                                       Term_Time = '%s', ETF = '%f', Contract_No = '%s', Account_No = '%s', Executed_by_UserID = '%s' \
                                   WHERE Circuit_ID = '%s'" %(self.Description, self.CountryID, self.RegionID, self.FacilityID, self.SitesID, \
                                        self.CircuitsCarrier, self.CircuitsTypeID, self.CircuitsPortID, self.CircuitsBandwidth, self.Circuitstatus, self.CircuitDMARK1,
                                        self.CircuitDMARK2, self.CircuitLEC1, self.CircuitLEC2, self.CircuitLEC3, self.CircuitLEC4, self.CircuitLEC5, self.CircuitCEASN,
                                        self.CircuitCEIPAddress, self.CircuitPEASN, self.CircuitPEIPAddress, self.CircuitVLANNo, self.CircuitPESwitch, self.CircuitPELocation,
                                        self.CircuitNPANXX, self.CircuitMonthlyCost, self.CircuitOrderNo,
                                        self.CircuitsInstalledDate, int(self.CircuitsInstalledDay), int(self.CircuitsInstalledMonth), int(self.CircuitsInstalledYear),
                                        self.CircuitsActivatedDate, int(self.CircuitsActivatedDay), int(self.CircuitsActivatedMonth), int(self.CircuitsActivatedYear),
                                        self.CircuitsDisconnectedDate, int(self.CircuitsDisconnectedDay), int(self.CircuitsDisconnectedMonth), int(self.CircuitsDisconnectedYear),
                                        self.CircuitsExpirationDate, int(self.CircuitsExpirationDay), int(self.CircuitsExpirationMonth), int(self.CircuitsExpirationYear),
                                        self.CircuitTerm, self.CircuitsTermTime, self.CircuitETF, self.CircuitsContractNo, self.CircuitAccountNo, self.Username, PrimaryKey)
                            '''
                    self.CircuitsTablePriaryKeyArray =  [] # Circuits ID
                    self.CircuitsTableDescriptionArray =  []                
                    self.CircuitsTableCountryIDArray =  [] 
                    self.CircuitsTableRegionIDArray =  []
                    self.CircuitsTableFacilityIDArray =  []
                    self.CircuitsTableSiteIDArray =  []
                    self.CircuitsTableCarrierIDArray =  []
                    self.CircuitsTableCircuitTypeArray = []
                    self.CircuitsTablePortSpeedArray = []
                    self.CircuitsTableBandwidthArray = []
                    self.CircuitsTableStatusArray = []
                    self.CircuitsTableDmarc_Info_1Array = []
                    self.CircuitsTableDmarc_Info_2Array = []
                    self.CircuitsTableLEC1Array = []
                    self.CircuitsTableLEC2Array = []
                    self.CircuitsTableLEC3Array = []
                    self.CircuitsTableLEC4Array = []
                    self.CircuitsTableLEC5Array = []
                    self.CircuitsTableCE_ASNArray = []
                    self.CircuitsTableCE_IP_AddressArray = []
                    self.CircuitsTablePE_ASNArray = []
                    self.CircuitsTablePE_IP_AddressArray = []
                    self.CircuitsTableVLAN_IDArray = []
                    self.CircuitsTablePE_SwitchArray = []
                    self.CircuitsTablePE_LocationArray = []
                    self.CircuitsTableNPA_NXXArray = []
                    self.CircuitsTableMonthlyCostArray = []
                    self.CircuitsTableOrderNumberArray = []
                    self.CircuitsTableDateInstalledArray = []
                    self.CircuitsTableDayInstalledArray = []
                    self.CircuitsTableMonthInstalledArray = []
                    self.CircuitsTableYearInstalledArray = []
                    self.CircuitsTableDateActivatedArray = []
                    self.CircuitsTableDayActivatedArray = []
                    self.CircuitsTableMonthActivatedArray = []
                    self.CircuitsTableYearActivatedArray = []
                    self.CircuitsTableDisconectedDateArray = []
                    self.CircuitsTableDayDisconectedArray = []
                    self.CircuitsTableMonthDisconectedArray = []
                    self.CircuitsTableYearDisconectedArray = []
                    self.CircuitsTableExpirationDateArray = []
                    self.CircuitsTableDayExpirationArray = []
                    self.CircuitsTableMonthExpirationArray = []
                    self.CircuitsTableYearExpirationArray = []
                    self.CircuitsTableTerm_DayArray = []
                    self.CircuitsTableTerm_TimeArray = []
                    self.CircuitsTableETFArray = []
                    self.CircuitsTableContract_NoArray = []
                    self.CircuitsTableAccount_NoArray = []
                    self.CircuitsTableExecutedByArray = []
                    
                        0 - Circuit_ID              CHAR(100) NOT NULL PRIMARY KEY,
                        1 - Description             CHAR(200),
                        2 - Country_ID              CHAR(20),                    
                        3 - Region_ID               CHAR(20),
                        4 - Facility_ID             CHAR(20),
                        5 - Site_ID                 CHAR(20),
                        6 - Carrier_ID              CHAR(20),
                        7 - Circuit_Type            CHAR(40),
                        8 - Port_Speed              CHAR(20),
                        9 - Bandwidth               CHAR(20),
                        10- Status                  CHAR(20),
                        11- Dmarc_Info_1            CHAR(200),
                        12- Dmarc_Info_2            CHAR(200),
                        13- LEC1                    CHAR(50),
                        14- LEC2                    CHAR(50),
                        15- LEC3                    CHAR(50),
                        16- LEC4                    CHAR(50),
                        17- LEC5                    CHAR(50),
                        18- CE_ASN                  CHAR(20),
                        19- CE_IP_Address           CHAR(50),
                        20- PE_ASN                  CHAR(20),
                        21- PE_IP_Address           CHAR(50),
                        22- VLAN_ID                 CHAR(10),
                        23- PE_Switch               CHAR(100),
                        24- PE_Location             CHAR(100),
                        25- NPA_NXX                 CHAR(20),
                        26- Monthly_Cost            FLOAT,
                        27- Order_Number            CHAR(40),
                        28- Date_Installed          CHAR(20),
                        29- Day_Installed           INT,
                        30- Month_Installed         INT,
                        31- Year_Installed          INT,
                        32- Date_Activated          CHAR(20),
                        33- Day_Activated           INT,
                        34- Month_Activated         INT,
                        35- Year_Activated          INT,
                        36- Disconnect_Date         CHAR(20),
                        37- Day_Disconnect          INT,
                        38- Month_Disconnect        INT,
                        39- Year_Disconnect         INT,
                        40- Expiration_Date         CHAR(20),
                        41- Day_Expiration          INT,
                        42- Month_Expiration        INT,
                        43- Year_Expiration         INT,
                        44- Term_Day                CHAR(10),
                        45- Term_Time               CHAR(10),
                        46- ETF                     FLOAT,
                        47- Contract_No             CHAR(50),
                        48- Account_No              CAHR(50),
                        49- Executed_by_UserID      CHAR(20))"""

                            '''
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.CircuitsFrame,title='Circuits',
                                    message = '*** The Circuits ID you entered was Updated ***')
                                if ((self.CountryID_Pre != self.CountryIDArray[self.ComboBoxCoutryID.current()]) or
                                    (self.RegionID_Pre != self.RegionIDArray[self.ComboBoxRegionID.current()]) or
                                    (self.FacilityID_Pre != self.FacilityIDArray[self.ComboBoxFacilityID.current()]) or
                                    (self.SitesID_Pre != self.SitesIDArray[self.ComboBoxSitesID.current()])):
                                    # ------ The Location Changed so we need to move to a new location !!!!!
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
                                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                                    message = '*** The Circuits ID you entered was NOT Upadted ***')
                            self.on_sites_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                            message = '*** The Circuits ID you try to Edit Does not exist Anymore ***')
                            self.on_sites_combo_changed("event")
                            self.Disable_Screen()
                    else:
                         mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                         message = '*** The CE or PE IP Address is invalid ***')                         
 
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Circuits_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','CANCEL Button']    
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
        self.Clean_Screen('Circuits','all')
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

    def on_Circuits_Tree_select_click(self,event):
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
        curItem = self.CircuitsTreeview.focus()  
        dic = self.CircuitsTreeview.item(curItem)
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
            self.ButtonCircuitsAdd['state']    = ACTIVE
            self.ButtonCircuitsEdit['state']   = ACTIVE
            self.ButtonCircuitsRemove['state'] = ACTIVE
            self.ButtonCircuitsOK['state']     = DISABLED
            self.ButtonCircuitsCancel['state'] = DISABLED
            self.ButtonDevicePingPE64['state'] = ACTIVE
            self.ButtonDevicePingPE1500['state'] = ACTIVE
            self.ButtonDevicePingCE64['state'] = ACTIVE
            self.ButtonDevicePingCE1500['state'] = ACTIVE
            self.ButtonDeviceCircuits['state'] = ACTIVE
            self.ButtonDeviceTraceroutePE['state'] = ACTIVE
            self.ButtonDeviceTracerouteCE['state'] = ACTIVE
            self.ButtonDeviceLocalPointOfContacts['state'] = ACTIVE
            self.Display_Screen(curItem)

    def On_Circuits_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_Circuits_Table_Refresh()

    #--------------------- Installed Date Calendar  <BEGIN> ------------------------------
    def on_CircuitsInstalledDateWindow_quit(self):
        self.CircuitsInstalledData = self.cal_Installed.kill_and_save()
        self.CircuitsCalendarInstalledDateExist = False
        self.CircuitsCalendarInstalledDateWindow.destroy()
        if (len(self.CircuitsInstalledData) > 0):
            self.CircuitsInstalledDate = str(self.CircuitsInstalledData['month_selected']) + '/' + str(self.CircuitsInstalledData['day_selected']) + '/' + str(self.CircuitsInstalledData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.CircuitsInstalledDate = ""
        self.CircuitsInstalledDateFrameEntry['state'] = 'normal'
        self.CircuitsInstalledDateFrameEntry.delete(0,END)
        self.CircuitsInstalledDateFrameEntry.insert(0,self.CircuitsInstalledDate)
        self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'

    def Call_Button_Installed_Date_Clear(self):
            self.CircuitsInstalledDateFrameEntry['state'] = 'normal'
            self.CircuitsInstalledDateFrameEntry.delete(0,END)
            self.CircuitsInstalledDateFrameEntry['state'] = 'readonly'
            
    def Call_Button_Activated_Date_Clear(self):
            self.CircuitsActivatedDateFrameEntry['state'] = 'normal'
            self.CircuitsActivatedDateFrameEntry.delete(0,END)
            self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'

    def Call_Button_Disconnected_Date_Clear(self):
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'normal'
            self.CircuitsDisconnectedDateFrameEntry.delete(0,END)
            self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'

    def Call_Button_Expiration_Date_Clear(self):
            self.CircuitsExpirationDateFrameEntry['state'] = 'normal'
            self.CircuitsExpirationDateFrameEntry.delete(0,END)
            self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Installed_Date(self):
        if not (self.CircuitsCalendarInstalledDateExist):
            self.CircuitsCalendarInstalledDateExist = True
            self.CircuitsCalendarInstalledDateWindow = Tk()
            self.CircuitsCalendarInstalledDateWindow.title("Installed")
            self.CircuitsCalendarInstalledDateWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitsInstalledDateWindow_quit)
            self.CircuitsCalendarInstalledDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.CircuitsInstalledData = {}
            if (self.Selection == 'edit'):
                self.CircuitsInstalledDateName = self.CircuitsInstalledDateFrameEntry.get()
                curItem = self.CircuitsTreeview.focus()  
                dic = self.CircuitsTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.CircuitsInstalledDateName) > 0):
                    if (self.CircuitsTableDateInstalledArray[curItem] != 0):
                        self.CircuitsInstalledData['day_selected'] = self.CircuitsTableDayInstalledArray[curItem]
                        self.CircuitsInstalledData['month_selected'] = self.CircuitsTableMonthInstalledArray[curItem]
                        self.CircuitsInstalledData['year_selected'] = self.CircuitsTableYearInstalledArray[curItem]
            self.cal_Installed = Class_Calendar(self.CircuitsCalendarInstalledDateWindow, self.CircuitsInstalledData)
            self.cal_Installed.setup()
    #--------------------- Installed Date Calendar  <END> ------------------------------

    #--------------------- Activated Date Calendar  <BEGIN> ------------------------------
    def on_CircuitsActivatedDateWindow_quit(self):
        self.CircuitsActivatedData = self.cal_Activated.kill_and_save()
        self.CircuitsCalendarActivatedDateExist = False
        self.CircuitsCalendarActivatedDateWindow.destroy()
        if (len(self.CircuitsActivatedData) > 0):
            self.CircuitsActivatedDate = str(self.CircuitsActivatedData['month_selected']) + '/' + str(self.CircuitsActivatedData['day_selected']) + '/' + str(self.CircuitsActivatedData['year_selected'])
            #print ("=> Day, Month, Year")
            #print (self.CircuitsActivatedData['day_selected'])
            #print (self.CircuitsActivatedData['month_selected'])
            #print (self.CircuitsActivatedData['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.CircuitsActivatedDate = ""
        self.CircuitsActivatedDateFrameEntry['state'] = 'normal'
        self.CircuitsActivatedDateFrameEntry.delete(0,END)
        self.CircuitsActivatedDateFrameEntry.insert(0,self.CircuitsActivatedDate)
        self.CircuitsActivatedDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Activated_Date(self):
        if not (self.CircuitsCalendarActivatedDateExist):
            self.CircuitsCalendarActivatedDateExist = True
            self.CircuitsCalendarActivatedDateWindow = Tk()
            self.CircuitsCalendarActivatedDateWindow.title("Activated")
            self.CircuitsCalendarActivatedDateWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitsActivatedDateWindow_quit)
            self.CircuitsCalendarActivatedDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.CircuitsActivatedData = {}
            if (self.Selection == 'edit'):
                self.CircuitsActivatedDateName = self.CircuitsActivatedDateFrameEntry.get()
                curItem = self.CircuitsTreeview.focus()  
                dic = self.CircuitsTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.CircuitsActivatedDateName) > 0):
                    if (self.CircuitsTableDateActivatedArray[curItem] != 0):
                        self.CircuitsActivatedData['day_selected'] = self.CircuitsTableDayActivatedArray[curItem]
                        self.CircuitsActivatedData['month_selected'] = self.CircuitsTableMonthActivatedArray[curItem]
                        self.CircuitsActivatedData['year_selected'] = self.CircuitsTableYearActivatedArray[curItem]
                        #print ("Day, Month, Year")
                        #print (self.CircuitsActivatedData['day_selected'])
                        #print (self.CircuitsActivatedData['month_selected'])
                        #print (self.CircuitsActivatedData['year_selected'])
 
            self.cal_Activated = Class_Calendar(self.CircuitsCalendarActivatedDateWindow, self.CircuitsActivatedData)
            self.cal_Activated.setup()
    #--------------------- Activated Date Calendar  <END> ------------------------------

    #--------------------- Disconnected Date Calendar  <BEGIN> ------------------------------
    def on_CircuitsDisconnectedDateWindow_quit(self):
        self.CircuitsDisconnectedData = self.cal_Disconnected.kill_and_save()
        self.CircuitsCalendarDisconnectedDateExist = False
        self.CircuitsCalendarDisconnectedDateWindow.destroy()
        if (len(self.CircuitsDisconnectedData) > 0):
            self.CircuitsDisconnectedDate = str(self.CircuitsDisconnectedData['month_selected']) + '/' + str(self.CircuitsDisconnectedData['day_selected']) + '/' + str(self.CircuitsDisconnectedData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.CircuitsDisconnectedDate = ""
        self.CircuitsDisconnectedDateFrameEntry['state'] = 'normal'
        self.CircuitsDisconnectedDateFrameEntry.delete(0,END)
        self.CircuitsDisconnectedDateFrameEntry.insert(0,self.CircuitsDisconnectedDate)
        self.CircuitsDisconnectedDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Disconnected_Date(self):
        if not (self.CircuitsCalendarDisconnectedDateExist):
            self.CircuitsCalendarDisconnectedDateExist = True
            self.CircuitsCalendarDisconnectedDateWindow = Tk()
            self.CircuitsCalendarDisconnectedDateWindow.title("Disconnected")
            self.CircuitsCalendarDisconnectedDateWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitsDisconnectedDateWindow_quit)
            self.CircuitsCalendarDisconnectedDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.CircuitsDisconnectedData = {}
            if (self.Selection == 'edit'):
                self.CircuitsDisconnectedDateName = self.CircuitsDisconnectedDateFrameEntry.get()
                curItem = self.CircuitsTreeview.focus()  
                dic = self.CircuitsTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.CircuitsDisconnectedDateName) > 0):
                    if (self.CircuitsTableDisconectedDateArray[curItem] != 0):
                        self.CircuitsDisconnectedData['day_selected'] = self.CircuitsTableDayDisconectedArray[curItem]
                        self.CircuitsDisconnectedData['month_selected'] = self.CircuitsTableMonthDisconectedArray[curItem]
                        self.CircuitsDisconnectedData['year_selected'] = self.CircuitsTableYearDisconectedArray[curItem]
            self.cal_Disconnected = Class_Calendar(self.CircuitsCalendarDisconnectedDateWindow, self.CircuitsDisconnectedData)
            self.cal_Disconnected.setup()
    #--------------------- Disconnected Date Calendar  <END> ------------------------------

    #--------------------- Expiration Date Calendar  <BEGIN> ------------------------------
    def on_CircuitsExpirationDateWindow_quit(self):
        self.CircuitsExpirationData = self.cal_Expiration.kill_and_save()
        self.CircuitsCalendarExpirationDateExist = False
        self.CircuitsCalendarExpirationDateWindow.destroy()
        if (len(self.CircuitsExpirationData) > 0):
            self.CircuitsExpirationDate = str(self.CircuitsExpirationData['month_selected']) + '/' + str(self.CircuitsExpirationData['day_selected']) + '/' + str(self.CircuitsExpirationData['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.CircuitsExpirationDate = ""
        self.CircuitsExpirationDateFrameEntry['state'] = 'normal'
        self.CircuitsExpirationDateFrameEntry.delete(0,END)
        self.CircuitsExpirationDateFrameEntry.insert(0,self.CircuitsExpirationDate)
        self.CircuitsExpirationDateFrameEntry['state'] = 'readonly'
        
    def Call_Button_Expiration_Date(self):
        if not (self.CircuitsCalendarExpirationDateExist):
            self.CircuitsCalendarExpirationDateExist = True
            self.CircuitsCalendarExpirationDateWindow = Tk()
            self.CircuitsCalendarExpirationDateWindow.title("Expiration")
            self.CircuitsCalendarExpirationDateWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitsExpirationDateWindow_quit)
            self.CircuitsCalendarExpirationDateWindow.call('tk', 'scaling', Windows_Scaling)
            self.CircuitsExpirationData = {}
            if (self.Selection == 'edit'):
                self.CircuitsExpirationDateName = self.CircuitsExpirationDateFrameEntry.get()
                curItem = self.CircuitsTreeview.focus()  
                dic = self.CircuitsTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.CircuitsExpirationDateName) > 0):
                    if (self.CircuitsTableExpirationDateArray[curItem] != 0):
                        self.CircuitsExpirationData['day_selected'] = self.CircuitsTableDayExpirationArray[curItem]
                        self.CircuitsExpirationData['month_selected'] = self.CircuitsTableMonthExpirationArray[curItem]
                        self.CircuitsExpirationData['year_selected'] = self.CircuitsTableYearExpirationArray[curItem]
            self.cal_Expiration = Class_Calendar(self.CircuitsCalendarExpirationDateWindow, self.CircuitsExpirationData)
            self.cal_Expiration.setup()
    #--------------------- Expiration Date Calendar  <END> ------------------------------


    def on_Circuits_type_combo_changed(self,event):
        #print (".")
        if self.db.Connect():
            # SQL Querry to the Circuits Table
            #self.CircuitsTypeVendorIDArray
            #self.CircuitComboBoxTypeID
            sql = """
                    SELECT * FROM Circuits_Model
                    WHERE Vendor_ID = '%s' AND Circuits_Type_ID = '%s'
                    ORDER BY Circuits_Model_Name ASC
                  """ % (self.CircuitsTypeVendorIDArray[self.CircuitComboBoxTypeID.current()],
                         self.CircuitTypeIDArray[self.CircuitComboBoxTypeID.current()])
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.CircuitPortSpeedIDArray = []
                self.CircuitPortSpeedNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CircuitPortSpeedIDArray.append(self.db.results[i][0].strip())
                    self.CircuitPortSpeedNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.CircuitsComboBoxPortSpeed['values'] = self.CircuitPortSpeedNameArray
                if (len(self.CircuitPortSpeedNameArray)== 0):
                    self.CircuitsComboBoxPortSpeed['state'] = DISABLED ######
                    self.Is_Get_Type_and_Model = False
                else:
                    self.CircuitsComboBoxPortSpeed['state'] = 'readonly'
                    self.CircuitsComboBoxPortSpeed.current(0)
                    #self.CircuitComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                sql = """
                        SELECT * FROM Circuits_Model
                        WHERE Vendor_ID = '%s'
                      """ % ('UNKNOWN')
                #print (sql)
                if (self.db.Execute(sql)):
                    #print ("found it")
                    i = 0
                    self.progress['maximum'] = len(self.db.results)
                    self.CircuitPortSpeedIDArray = []
                    self.CircuitPortSpeedNameArray = []
                    while (i < len(self.db.results)):
                        num = i + 1
                        self.CircuitPortSpeedIDArray.append(self.db.results[i][0].strip())
                        self.CircuitPortSpeedNameArray.append(self.db.results[i][1].strip())
                        i = i + 1
                        self.progress['value'] = i
                    self.CircuitsComboBoxPortSpeed['values'] = self.CircuitPortSpeedNameArray
                    if (len(self.CircuitPortSpeedNameArray)== 0):
                        self.CircuitsComboBoxPortSpeed['state'] = DISABLED ########
                        self.Is_Get_Type_and_Model = False
                    else:
                        self.CircuitsComboBoxPortSpeed['state'] = 'readonly'
                        self.CircuitsComboBoxPortSpeed.current(0)
                        #self.CircuitComboBoxTypeID.set("")
                        self.Is_Get_Type_and_Model = True
                else:
                    mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                    message = 'No Records found for Circuits Model')                
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def Get_Type_PortSpeed_and_Satus(self):
        #print ("")
        # self.CircuitComboBoxTypeID
        # self.CircuitsComboBoxPortSpeed
        # self.CircuitsComboBoxStatus
        self.Is_Get_Type_and_Model = False
        if self.db.Connect():
            # SQL Querry to the Circuits Type
            sql = """
                    SELECT * FROM Circuit_Type
                    ORDER BY Description ASC
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.CircuitTypeIDArray = []
                self.CircuitTypeNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CircuitTypeIDArray.append(self.db.results[i][0].strip())
                    self.CircuitTypeNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.CircuitComboBoxTypeID['values'] = self.CircuitTypeNameArray
                if (len(self.CircuitTypeNameArray)== 0):
                    self.CircuitComboBoxTypeID['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.CircuitComboBoxTypeID['state'] = DISABLED #'readonly'
                    self.CircuitComboBoxTypeID.current(0)
                    #self.CircuitComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found for Circuits Type')

            # SQL Querry to the Port Speed
            sql = """
                    SELECT * FROM PORT_SPEED
                    ORDER BY Description ASC
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.CircuitPortSpeedIDArray = []
                self.CircuitPortSpeedNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CircuitPortSpeedIDArray.append(self.db.results[i][0].strip())
                    self.CircuitPortSpeedNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.CircuitsComboBoxPortSpeed['values'] = self.CircuitPortSpeedNameArray
                if (len(self.CircuitPortSpeedNameArray)== 0):
                    self.CircuitsComboBoxPortSpeed['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.CircuitsComboBoxPortSpeed['state'] = DISABLED #'readonly'
                    self.CircuitsComboBoxPortSpeed.current(0)
                    #self.CircuitComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found for Port Speed')

            # SQL Querry to the Carrier
            sql = """
                    SELECT * FROM CARRIER
                    ORDER BY Description ASC
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                #print ("found it")
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.CircuitCarrierIDArray = []
                self.CircuitCarrierNameArray = []
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CircuitCarrierIDArray.append(self.db.results[i][0].strip())
                    self.CircuitCarrierNameArray.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.CircuitComboBoxCarrier['values'] = self.CircuitCarrierNameArray
                if (len(self.CircuitCarrierNameArray)== 0):
                    self.CircuitComboBoxCarrier['state'] = DISABLED
                    self.Is_Get_Type_and_Model = False
                else:
                    self.CircuitComboBoxCarrier['state'] = DISABLED #'readonly'
                    self.CircuitComboBoxCarrier.current(0)
                    #self.CircuitComboBoxTypeID.set("")
                    self.Is_Get_Type_and_Model = True
            else:
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                message = 'No Records found for Carrier')




                
        else:
            mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        # ---- Status ----
        self.CircuitstatusValues = ['Active','In Process','Research','Billed','Inactive']
        self.CircuitsComboBoxStatus['values'] = self.CircuitstatusValues
        self.CircuitsComboBoxStatus['state'] = DISABLED
        self.CircuitsComboBoxStatus.current(0)

        # ---- Term ----
        self.CircuitTermValues = ['Days','Months','Years']
        self.CircuitComboBoxTerm['values'] = self.CircuitTermValues
        self.CircuitComboBoxTerm['state'] = DISABLED
        self.CircuitComboBoxTerm.current(0)

    def On_Run_Traceroute(self,ip):
        #print ("PING.....")
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','TRACEROUTE Button']    
            Logging.Log(Parameter)        
        if (len(ip) > 6):
            '''
            import os
            if os.system("ping -c 1 google.com") == 0:
            ...     print "host appears to be up"
            '''
            cmd = ("tracert -d -w 500 " + ip)
            os.system(cmd)
            #C:\Users\rod90731\Downloads>
            cmd = ("tracert -d -w 500 " + ip + " > C:\\users\\" + self.Username + "\\Downloads\\traceroute-" + ip + ".txt")
            os.system(cmd)
            mbox.showwarning(master=self.CircuitsFrame,title='Traceroute',
                message = "The Traceroute output was saved at: \n" +
                          "C:\\users\\" + self.Username + "\\Downloads\\traceroute-" + ip + ".txt")


    def On_Traceroute(self,ip):
        try:
            self.thread = Thread(target=self.On_Run_Traceroute, args=(ip,))
            self.thread.daemon=True
            self.thread.start()
        except(KeyboardInterrupt, SystemExit):
            sys.exit("Interrupted by ctrl+c\n")

    def On_Run_Ping(self,ip,size):
        #print ("PING.....")
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING Button']    
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
            
    def Get_CE_IPAddress(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING CE 64 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitCEIPAddressFrameEntry.get()
        self.On_Ping(ip,'64')

    def Call_Button_PingCE64(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING CE 64 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitCEIPAddressFrameEntry.get()
        self.On_Ping(ip,'64')

    def Call_Button_PingCE1500(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING CE 1500 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitCEIPAddressFrameEntry.get()
        self.On_Ping(ip,'1500')

    def Call_Button_TracerouteCE(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','TRACEROUTE CE Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitCEIPAddressFrameEntry.get()
        self.On_Traceroute(ip)

    def Get_PE_IPAddress(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING PE 64 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitPEIPAddressFrameEntry.get()
        self.On_Ping(ip,'64')

    def Call_Button_PingPE64(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING PE 64 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitPEIPAddressFrameEntry.get()
        self.On_Ping(ip,'64')

    def Call_Button_PingPE1500(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','PING PE 1500 Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitPEIPAddressFrameEntry.get()
        self.On_Ping(ip,'1500')

    def Call_Button_TraceroutePE(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','TRACEROUTE PE Button']    
            Logging.Log(Parameter)        
        ip = self.CircuitPEIPAddressFrameEntry.get()
        self.On_Traceroute(ip)                    

    def Call_Button_Carriers(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','Carriers Button']    
            Logging.Log(Parameter)
            Carrier.Display_Carrier_Window()
        #mbox.showerror(master=self.CircuitsFrame,title='Under Construction',
        #                 message = '... Comming Soon...')
        
    def Call_Button_LocalPointOfContacts(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','Local POC Button']    
            Logging.Log(Parameter)        
        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]        
        Location = []
        Location = [self.CountryID, self.RegionID, self.FacilityID, self.SitesID]
        LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,Location)
        LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def Call_Button_Devices(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Circuits','Devices Button']    
            Logging.Log(Parameter)        
        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]        
        Location = []
        Location = [self.CountryID, self.RegionID, self.FacilityID, self.SitesID]
        Device = Class_Device(ODBC_DSN_name,Windows_Scaling,Location)
        Device.Display_Device_Window()


    def Call_Save_As(self):
        #self.data_ready = True    
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Circuit List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.circuitsTableColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.CircuitsFrame,title='Circuits',
                            message = '!!! The File was saved !!!')
                    self.CircuitsWindow.title("Circuits                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Circuits','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                            message = '*** The File was not saved, Perhaps It is already open ***')



    def Display_Circuits_Window(self): 
        if not self.CircuitsWindowExist:
            # Set up the Window
            self.CircuitsWindowExist = True
            self.CircuitsWindow = Tk()
            self.CircuitsWindow.geometry('1350x820+350+70')
            self.CircuitsWindow.title("Circuits")
            self.CircuitsWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitsWindow_quit)
            self.CircuitsWindow.call('tk', 'scaling', Windows_Scaling)
            CircuitsLabel = Label(self.CircuitsWindow,text="Helvetica", font=("Helvetica", 19))
            CircuitsLabel["text"] = "Circuits"
            CircuitsLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.CircuitsWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_CircuitsWindow_quit)
            self.CircuitsWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------

                    
            # Setup Frame
            self.CircuitsFrame = Frame(self.CircuitsWindow)
            self.CircuitsFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.CircuitsFrame.rowconfigure(0, weight=1)
            self.CircuitsFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.CircuitsFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 8, width=75, height=24)

                        self.ButtonCountryRefresh = Button(self.CircuitsFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 8, width=75, height=24)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.CircuitsFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 38, width=75, height=24)

                        self.ButtonRegionRefresh = Button(self.CircuitsFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 38, width=75, height=24)

                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd = Button(self.CircuitsFrame, text = '+ Facility', command = self.Display_Facility_Window, state=DISABLED)
                        self.ButtonFacilityAdd.place(x = 450, y = 68, width=75, height=24)

                        self.ButtonFacilityRefresh = Button(self.CircuitsFrame, text = 'Refresh', command = self.on_Facility_Table_Refresh, state=DISABLED)
                        self.ButtonFacilityRefresh.place(x = 550, y = 68, width=75, height=24)

                    if (Is_Sites_Available):
                        self.ButtonSitesAdd = Button(self.CircuitsFrame, text = '+ Site', command = self.Display_Sites_Window, state=DISABLED)
                        self.ButtonSitesAdd.place(x = 450, y = 98, width=75, height=24)

                        self.ButtonSitesRefresh = Button(self.CircuitsFrame, text = 'Refresh', command = self.on_Sites_Table_Refresh, state=DISABLED)
                        self.ButtonSitesRefresh.place(x = 550, y = 98, width=75, height=24)


                    self.ButtonCircuitsAdd = Button(self.CircuitsFrame, text = 'Add', command = self.Call_Button_Circuits_Add, state=DISABLED)
                    self.ButtonCircuitsAdd.place(x = 450, y = 128, width=75, height=25)

                    self.ButtonCircuitsEdit = Button(self.CircuitsFrame, text = 'Edit', command = self.Call_Button_Circuits_Edit, state=DISABLED)
                    self.ButtonCircuitsEdit.place(x = 550, y = 128, width=75, height=25)

                    self.ButtonCircuitsRemove = Button(self.CircuitsFrame, text = 'Remove', command = self.Call_Button_Circuits_Remove, state=DISABLED)
                    self.ButtonCircuitsRemove.place(x = 650, y = 128, width=75, height=25)

                    self.ButtonCircuitsOK = Button(self.CircuitsFrame, text = 'OK / UPDATE', command = self.Call_Button_Circuits_OK, state=DISABLED)
                    self.ButtonCircuitsOK.place(x = 750, y = 128, width=100, height=25)

                    self.ButtonCircuitsCancel = Button(self.CircuitsFrame, text = 'Cancel', command = self.Call_Button_Circuits_Cancel, state=DISABLED)
                    self.ButtonCircuitsCancel.place(x = 875, y = 128, width=75, height=25)

                    # Utilities Buttons
                    #self.ButtonCircuitDevice = Button(self.CircuitsFrame, text = 'Devices', command = self.Call_Button_Devices, state=DISABLED)
                    #self.ButtonCircuitDevice.place(x = 750, y = 8, width=100, height=25)

                    self.ButtonDeviceLocalPointOfContacts = Button(self.CircuitsFrame, text = 'Local POC', command = self.Call_Button_LocalPointOfContacts, state=DISABLED)
                    self.ButtonDeviceLocalPointOfContacts.place(x = 638, y = 8, width=100, height=25)

                    self.ButtonDeviceCircuits = Button(self.CircuitsFrame, text = 'Carriers', command = self.Call_Button_Carriers, state=DISABLED)
                    self.ButtonDeviceCircuits.place(x = 750, y = 8, width=100, height=25)

                    self.ButtonDevicePingPE64 = Button(self.CircuitsFrame, text = 'Ping PE 64', command = self.Call_Button_PingPE64, state=DISABLED)
                    self.ButtonDevicePingPE64.place(x = 638, y = 38, width=100, height=25)

                    self.ButtonDevicePingCE64 = Button(self.CircuitsFrame, text = 'Ping CE 64', command = self.Call_Button_PingCE64, state=DISABLED)
                    self.ButtonDevicePingCE64.place(x = 750, y = 38, width=100, height=25)

                    self.ButtonDevicePingPE1500 = Button(self.CircuitsFrame, text = 'Ping PE 1500', command = self.Call_Button_PingPE1500, state=DISABLED)
                    self.ButtonDevicePingPE1500.place(x = 638, y = 68, width=100, height=25)

                    self.ButtonDevicePingCE1500 = Button(self.CircuitsFrame, text = 'Ping CE 1500', command = self.Call_Button_PingCE1500, state=DISABLED)
                    self.ButtonDevicePingCE1500.place(x = 750, y = 68, width=100, height=25)

                    self.ButtonDeviceTraceroutePE = Button(self.CircuitsFrame, text = 'Traceroute PE', command = self.Call_Button_TraceroutePE, state=DISABLED)
                    self.ButtonDeviceTraceroutePE.place(x = 638, y = 98, width=100, height=25)

                    self.ButtonDeviceTracerouteCE = Button(self.CircuitsFrame, text = 'Traceroute CE', command = self.Call_Button_TracerouteCE, state=DISABLED)
                    self.ButtonDeviceTracerouteCE.place(x = 750, y = 98, width=100, height=25)



                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.CircuitsFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=158)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=40)

                    FacilityIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityIDFrameLabel["text"] = "Facility Name:"
                    FacilityIDFrameLabel.place(x=10, y=70)

                    SitesIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    SitesIDFrameLabel["text"] = "Site Name:"
                    SitesIDFrameLabel.place(x=10, y=100)

                    # Setup Labels and Entry
                    self.CircuitIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitIDFrameLabel["text"] = "Circuit ID:"
                    self.CircuitIDFrameLabel.place(x=10, y=130)
                    self.CircuitIDFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitIDFrameEntry['width']=50
                    self.CircuitIDFrameEntry.place(x=110, y=130)
                    self.CircuitIDFrameEntry['state'] = DISABLED
                    
                    self.CircuitsDescriptionFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsDescriptionFrameLabel["text"] = "Description:"
                    self.CircuitsDescriptionFrameLabel.place(x=10, y=160)
                    self.CircuitsDescriptionFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsDescriptionFrameEntry['width']=50
                    self.CircuitsDescriptionFrameEntry.place(x=110, y=160)
                    self.CircuitsDescriptionFrameEntry['state'] = DISABLED               

                    # ComboBox for Type, Model, Status
                    
                    TypeIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    TypeIDFrameLabel["text"] = "Type:"
                    TypeIDFrameLabel.place(x=10, y=190)
                    self.CircuitComboBoxTypeID = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 20)
                    #self.CircuitComboBoxTypeID.bind("<<ComboboxSelected>>", self.on_Circuits_type_combo_changed)
                    self.CircuitComboBoxTypeID.place(x = 55, y = 190)

                    self.ModelIDFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.ModelIDFrameLabel["text"] = "Port Speed:"
                    self.ModelIDFrameLabel.place(x=205, y=190)
                    self.CircuitsComboBoxPortSpeed = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 20)
                    self.CircuitsComboBoxPortSpeed.place(x = 280, y = 190)
                    #self.CircuitsComboBoxPortSpeed.bind("<<ComboboxSelected>>", self.on_country_combo_changed)

                    self.CircuitsBandwidthFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsBandwidthFrameLabel["text"] = "Bandwidth:"
                    self.CircuitsBandwidthFrameLabel.place(x=430, y=190)
                    self.CircuitsBandwidthFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsBandwidthFrameEntry['width']=17
                    self.CircuitsBandwidthFrameEntry.place(x=510, y=190)
                    self.CircuitsBandwidthFrameEntry['state'] = DISABLED

                    self.CircuitStatusFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitStatusFrameLabel["text"] = "Status:"
                    self.CircuitStatusFrameLabel.place(x=625, y=190)
                    self.CircuitsComboBoxStatus = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 20)
                    self.CircuitsComboBoxStatus.place(x = 675, y = 190)

                    self.CircuitTermFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitTermFrameLabel["text"] = "Term:"
                    self.CircuitTermFrameLabel.place(x=825, y=190)
                    self.CircuitTermFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitTermFrameEntry['width']=6
                    self.CircuitTermFrameEntry.place(x=870, y=190)
                    self.CircuitTermFrameEntry['state'] = DISABLED
                    self.CircuitComboBoxTerm = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 10)
                    self.CircuitComboBoxTerm.place(x = 915, y = 189)

                    # Setup Labels and Button Calendars Installed, Activated, Disconnected
                    self.CircuitsInstalledDateFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsInstalledDateFrameLabel["text"] = "Installed Date:"
                    self.CircuitsInstalledDateFrameLabel.place(x=10, y=220)
                    self.CircuitsInstalledDateFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsInstalledDateFrameEntry['width']=20
                    self.CircuitsInstalledDateFrameEntry.place(x=130, y=220)
                    self.CircuitsInstalledDateFrameEntry['state'] = DISABLED
                    self.CircuitsButtonInstalledDate = Button(self.CircuitsFrame, text = 'Calendar', command = self.Call_Button_Installed_Date, state=DISABLED)
                    self.CircuitsButtonInstalledDate.place(x = 260 , y = 217, width=75, height=25)
                    self.CircuitsButtonInstalledDateClear = Button(self.CircuitsFrame, text = 'Clear Date', command = self.Call_Button_Installed_Date_Clear, state=DISABLED)
                    self.CircuitsButtonInstalledDateClear.place(x = 345 , y = 217, width=75, height=25)


                    self.CircuitsContractNoFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsContractNoFrameLabel["text"] = "Contract No:"
                    self.CircuitsContractNoFrameLabel.place(x=430, y=220)
                    self.CircuitsContractNoFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsContractNoFrameEntry['width']=17
                    self.CircuitsContractNoFrameEntry.place(x=510, y=220)
                    self.CircuitsContractNoFrameEntry['state'] = DISABLED

                    self.CircuitAccountNoFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitAccountNoFrameLabel["text"] = "Account No:"
                    self.CircuitAccountNoFrameLabel.place(x=625, y=220)
                    self.CircuitAccountNoFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitAccountNoFrameEntry['width']=20
                    self.CircuitAccountNoFrameEntry.place(x=725, y=220)
                    self.CircuitAccountNoFrameEntry['state'] = DISABLED

                    self.CircuitOrderNoFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitOrderNoFrameLabel["text"] = "Order No:"
                    self.CircuitOrderNoFrameLabel.place(x=855, y=220)
                    self.CircuitOrderNoFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitOrderNoFrameEntry['width']=20
                    self.CircuitOrderNoFrameEntry.place(x=925, y=220)
                    self.CircuitOrderNoFrameEntry['state'] = DISABLED


                    self.CircuitCarrierFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitCarrierFrameLabel["text"] = "Carrier:"
                    self.CircuitCarrierFrameLabel.place(x=1060, y=220)
                    self.CircuitComboBoxCarrier = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 20)
                    self.CircuitComboBoxCarrier.place(x = 1145, y = 220)


                    self.CircuitsActivatedDateFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsActivatedDateFrameLabel["text"] = "Activated Date:"
                    self.CircuitsActivatedDateFrameLabel.place(x=10, y=250)
                    self.CircuitsActivatedDateFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsActivatedDateFrameEntry['width']=20
                    self.CircuitsActivatedDateFrameEntry.place(x=130, y=250)
                    self.CircuitsActivatedDateFrameEntry['state'] = DISABLED
                    self.CircuitsButtonActivatedDate = Button(self.CircuitsFrame, text = 'Calendar', command = self.Call_Button_Activated_Date, state=DISABLED)
                    self.CircuitsButtonActivatedDate.place(x = 260 , y = 247, width=75, height=25)
                    self.CircuitsButtonActivatedDateClear = Button(self.CircuitsFrame, text = 'Clear Date', command = self.Call_Button_Activated_Date_Clear, state=DISABLED)
                    self.CircuitsButtonActivatedDateClear.place(x = 345 , y = 247, width=75, height=25)


                    self.CircuitCEASNFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitCEASNFrameLabel["text"] = "CE ASN No:"
                    self.CircuitCEASNFrameLabel.place(x=430, y=250)
                    self.CircuitCEASNFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitCEASNFrameEntry['width']=17
                    self.CircuitCEASNFrameEntry.place(x=510, y=250)
                    self.CircuitCEASNFrameEntry['state'] = DISABLED

                    self.CircuitCEIPAddressFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitCEIPAddressFrameLabel["text"] = "CE IP Address:"
                    self.CircuitCEIPAddressFrameLabel.place(x=625, y=250)
                    self.CircuitCEIPAddressFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitCEIPAddressFrameEntry['width']=20
                    self.CircuitCEIPAddressFrameEntry.place(x=725, y=250)
                    self.CircuitCEIPAddressFrameEntry['state'] = DISABLED
                    self.CircuitCEIPAddressFrameEntry.bind("<Double-1>", self.Get_CE_IPAddress)

                    self.CircuitVLANNoFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitVLANNoFrameLabel["text"] = "VLAN No:"
                    self.CircuitVLANNoFrameLabel.place(x=855, y=250)
                    self.CircuitVLANNoFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitVLANNoFrameEntry['width']=20
                    self.CircuitVLANNoFrameEntry.place(x=925, y=250)
                    self.CircuitVLANNoFrameEntry['state'] = DISABLED

                    self.CircuitNPANXXFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitNPANXXFrameLabel["text"] = "NPA-NXX:"
                    self.CircuitNPANXXFrameLabel.place(x=1060, y=250)
                    self.CircuitNPANXXFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitNPANXXFrameEntry['width']=20
                    self.CircuitNPANXXFrameEntry.place(x=1145, y=250)
                    self.CircuitNPANXXFrameEntry['state'] = DISABLED


                    self.CircuitPEASNFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitPEASNFrameLabel["text"] = "PE ASN No:"
                    self.CircuitPEASNFrameLabel.place(x=430, y=280)
                    self.CircuitPEASNFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitPEASNFrameEntry['width']=17
                    self.CircuitPEASNFrameEntry.place(x=510, y=280)
                    self.CircuitPEASNFrameEntry['state'] = DISABLED

                    self.CircuitPEIPAddressFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitPEIPAddressFrameLabel["text"] = "PE IP Address:"
                    self.CircuitPEIPAddressFrameLabel.place(x=625, y=280)
                    self.CircuitPEIPAddressFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitPEIPAddressFrameEntry['width']=20
                    self.CircuitPEIPAddressFrameEntry.place(x=725, y=280)
                    self.CircuitPEIPAddressFrameEntry['state'] = DISABLED
                    self.CircuitPEIPAddressFrameEntry.bind("<Double-1>", self.Get_PE_IPAddress)

                    self.CircuitPESwitchFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitPESwitchFrameLabel["text"] = "PE Switch:"
                    self.CircuitPESwitchFrameLabel.place(x=855, y=280)
                    self.CircuitPESwitchFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitPESwitchFrameEntry['width']=20
                    self.CircuitPESwitchFrameEntry.place(x=925, y=280)
                    self.CircuitPESwitchFrameEntry['state'] = DISABLED

                    self.CircuitPELocationFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitPELocationFrameLabel["text"] = "PE Location:"
                    self.CircuitPELocationFrameLabel.place(x=1060, y=280)
                    self.CircuitPELocationFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitPELocationFrameEntry['width']=20
                    self.CircuitPELocationFrameEntry.place(x=1145, y=280)
                    self.CircuitPELocationFrameEntry['state'] = DISABLED

                    self.CircuitsDisconnectedDateFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsDisconnectedDateFrameLabel["text"] = "Disconnected Date:"
                    self.CircuitsDisconnectedDateFrameLabel.place(x=10, y=280)
                    self.CircuitsDisconnectedDateFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsDisconnectedDateFrameEntry['width']=20
                    self.CircuitsDisconnectedDateFrameEntry.place(x=130, y=280)
                    self.CircuitsDisconnectedDateFrameEntry['state'] = DISABLED
                    self.CircuitsButtonDisconnectedDate = Button(self.CircuitsFrame, text = 'Calendar', command = self.Call_Button_Disconnected_Date, state=DISABLED)
                    self.CircuitsButtonDisconnectedDate.place(x = 260 , y = 277, width=75, height=25)
                    self.CircuitsButtonDisconnectedDateClear = Button(self.CircuitsFrame, text = 'Clear Date', command = self.Call_Button_Disconnected_Date_Clear, state=DISABLED)
                    self.CircuitsButtonDisconnectedDateClear.place(x = 345 , y = 277, width=75, height=25)


                    self.CircuitsExpirationDateFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitsExpirationDateFrameLabel["text"] = "Expiration Date:"
                    self.CircuitsExpirationDateFrameLabel.place(x=10, y=310)
                    self.CircuitsExpirationDateFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitsExpirationDateFrameEntry['width']=20
                    self.CircuitsExpirationDateFrameEntry.place(x=130, y=310)
                    self.CircuitsExpirationDateFrameEntry['state'] = DISABLED
                    self.CircuitsButtonExpirationDate = Button(self.CircuitsFrame, text = 'Calendar', command = self.Call_Button_Expiration_Date, state=DISABLED)
                    self.CircuitsButtonExpirationDate.place(x = 260 , y = 307, width=75, height=25)
                    self.CircuitsButtonExpirationDateClear = Button(self.CircuitsFrame, text = 'Clear Date', command = self.Call_Button_Expiration_Date_Clear, state=DISABLED)
                    self.CircuitsButtonExpirationDateClear.place(x = 345 , y = 307, width=75, height=25)


                    self.CircuitMonthlyCostFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitMonthlyCostFrameLabel["text"] = "Monthly $:"
                    self.CircuitMonthlyCostFrameLabel.place(x=430, y=310)
                    self.CircuitMonthlyCostFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitMonthlyCostFrameEntry['width']=17
                    self.CircuitMonthlyCostFrameEntry.place(x=510, y=310)
                    self.CircuitMonthlyCostFrameEntry['state'] = DISABLED

                    self.CircuitETFFrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitETFFrameLabel["text"] = "Early Term Fees $:"
                    self.CircuitETFFrameLabel.place(x=625, y=310)
                    self.CircuitETFFrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitETFFrameEntry['width']=20
                    self.CircuitETFFrameEntry.place(x=750, y=310)
                    self.CircuitETFFrameEntry['state'] = DISABLED

                    self.CircuitDMARK1FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitDMARK1FrameLabel["text"] = "DMARK Info 1:"
                    self.CircuitDMARK1FrameLabel.place(x=10, y=340)
                    self.CircuitDMARK1FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitDMARK1FrameEntry['width']=137
                    self.CircuitDMARK1FrameEntry.place(x=130, y=340)
                    self.CircuitDMARK1FrameEntry['state']=DISABLED

                    self.CircuitDMARK2FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitDMARK2FrameLabel["text"] = "DMARK Info 2:"
                    self.CircuitDMARK2FrameLabel.place(x=10, y=370)
                    self.CircuitDMARK2FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitDMARK2FrameEntry['width']=137
                    self.CircuitDMARK2FrameEntry.place(x=130, y=370)
                    self.CircuitDMARK2FrameEntry['state']=DISABLED

                    self.CircuitLEC1FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitLEC1FrameLabel["text"] = "LEC 1:"
                    self.CircuitLEC1FrameLabel.place(x=10, y=400)
                    self.CircuitLEC1FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitLEC1FrameEntry['width']=50
                    self.CircuitLEC1FrameEntry.place(x=130, y=400)
                    self.CircuitLEC1FrameEntry['state']=DISABLED

                    self.CircuitLEC2FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitLEC2FrameLabel["text"] = "LEC 2:"
                    self.CircuitLEC2FrameLabel.place(x=460, y=400)
                    self.CircuitLEC2FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitLEC2FrameEntry['width']=50
                    self.CircuitLEC2FrameEntry.place(x=520, y=400)
                    self.CircuitLEC2FrameEntry['state']=DISABLED

                    self.CircuitLEC3FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitLEC3FrameLabel["text"] = "LEC 3:"
                    self.CircuitLEC3FrameLabel.place(x=840, y=400)
                    self.CircuitLEC3FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitLEC3FrameEntry['width']=50
                    self.CircuitLEC3FrameEntry.place(x=900, y=400)
                    self.CircuitLEC3FrameEntry['state']=DISABLED


                    self.CircuitLEC4FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitLEC4FrameLabel["text"] = "LEC 4:"
                    self.CircuitLEC4FrameLabel.place(x=10, y=430)
                    self.CircuitLEC4FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitLEC4FrameEntry['width']=50
                    self.CircuitLEC4FrameEntry.place(x=130, y=430)
                    self.CircuitLEC4FrameEntry['state']=DISABLED

                    self.CircuitLEC5FrameLabel = Label(self.CircuitsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.CircuitLEC5FrameLabel["text"] = "LEC 5:"
                    self.CircuitLEC5FrameLabel.place(x=460, y=430)
                    self.CircuitLEC5FrameEntry = Entry(self.CircuitsFrame)
                    self.CircuitLEC5FrameEntry['width']=50
                    self.CircuitLEC5FrameEntry.place(x=520, y=430)
                    self.CircuitLEC5FrameEntry['state']=DISABLED

                    #------------------ TREE VIEW For Circuits Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars                            
                    self.CircuitsTreeviewDataColumns = ('Circuits ID','Description','Carrier ID','Type','Speed','BW',
                                                       'Monthly Cost','ETF','Installed',
                                                       'Activated','Disconnected','Status','Contract No','Expiration',
                                                       'Added By')
                    
                    self.CircuitsTreeview = ttk.Treeview(self.CircuitsFrame,columns=self.CircuitsTreeviewDataColumns, height=8) # <--- Make sure the frame is correct !!
                    self.CircuitsTreeviewysb = Scrollbar(self.CircuitsFrame,orient=VERTICAL, command=self.CircuitsTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.CircuitsTreeviewxsb = Scrollbar(self.CircuitsFrame,orient=HORIZONTAL, command=self.CircuitsTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.CircuitsTreeview['yscroll'] = self.CircuitsTreeviewysb.set
                    self.CircuitsTreeview['xscroll'] = self.CircuitsTreeviewxsb.set

                    # setup headings and column 
                    self.CircuitsTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                    self.CircuitsTreeview.heading('#1', text='Circuits ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Circuits ID', False)) # E for East and W for West
                    self.CircuitsTreeview.heading('#2', text='Description',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Description', False)) # E for East and W for West
                    self.CircuitsTreeview.heading('#3', text='Carrier ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Carrier ID', False)) # E for East and W for West
                    self.CircuitsTreeview.heading('#4', text='Type', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Type', False)) # E for East and W for West
                    self.CircuitsTreeview.heading('#5', text='Speed', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Speed', False)) # E for East and W for West
                    self.CircuitsTreeview.heading('#6', text='BW', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'BW', False))
                    #self.CircuitsTreeview.heading('#6', text='CE ASN', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'CE ASN', False))
                    #self.CircuitsTreeview.heading('#7', text='CE IP', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'CE IP', False))
                    #self.CircuitsTreeview.heading('#8', text='PE ASN', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'PE ASN', False))
                    #self.CircuitsTreeview.heading('#9', text='PE IP', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'PE IP', False))
                    #self.CircuitsTreeview.heading('#10', text='VLAN No', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'VLAN No', False))
                    self.CircuitsTreeview.heading('#7', text='Monthly Cost', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Monthly Cost', False))
                    self.CircuitsTreeview.heading('#8', text='ETF', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'ETF', False))
                    self.CircuitsTreeview.heading('#9', text='Installed', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Installed', False))
                    self.CircuitsTreeview.heading('#10', text='Activated', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Activated', False))
                    self.CircuitsTreeview.heading('#11', text='Disconnected', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Disconnected', False))
                    self.CircuitsTreeview.heading('#12', text='Status', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Status', False))
                    self.CircuitsTreeview.heading('#13', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Contract No', False))
                    self.CircuitsTreeview.heading('#14', text='Expiration', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Expiration', False))
                    self.CircuitsTreeview.heading('#15', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Added By', False)) 
                    
                    self.CircuitsTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.CircuitsTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.CircuitsTreeview.column('#2', stretch=1, width=10)
                    self.CircuitsTreeview.column('#3', stretch=1, width=10)
                    self.CircuitsTreeview.column('#4', stretch=1, width=10)
                    self.CircuitsTreeview.column('#5', stretch=1, width=10)
                    self.CircuitsTreeview.column('#6', stretch=1, width=10)
                    self.CircuitsTreeview.column('#7', stretch=1, width=10)
                    self.CircuitsTreeview.column('#8', stretch=1, width=10)
                    self.CircuitsTreeview.column('#9', stretch=1, width=10)
                    self.CircuitsTreeview.column('#10', stretch=1, width=10)
                    self.CircuitsTreeview.column('#11', stretch=1, width=10)
                    self.CircuitsTreeview.column('#12', stretch=1, width=10)
                    self.CircuitsTreeview.column('#13', stretch=1, width=10)
                    self.CircuitsTreeview.column('#14', stretch=1, width=10)
                    self.CircuitsTreeview.column('#15', stretch=1, width=10)
                    #self.CircuitsTreeview.column('#16', stretch=1, width=10)
                    #self.CircuitsTreeview.column('#17', stretch=1, width=10)
                    #self.CircuitsTreeview.column('#18', stretch=1, width=10)
                    #self.CircuitsTreeview.column('#19', stretch=1, width=10)



                    # add tree and scrollbars to frame
                    self.CircuitsTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.CircuitsTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.CircuitsTreeviewxsb.grid(row=2, column=0, sticky=EW)

                    # create fonts and tags
                    # Use later to mark Business Units per color.
                    self.CircuitsTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                    #self.CircuitsTreeview.tag_configure('Active', font=('Helvetica', 8), background='green2')
                    self.CircuitsTreeview.tag_configure('Inactive', font=('Helvetica', 8), background='ivory3') # 'red'
                    self.CircuitsTreeview.tag_configure('In Process', font=('Helvetica', 8), background='green3')
                    self.CircuitsTreeview.tag_configure('Billed', font=('Helvetica', 8), background='yellow2')
                    self.CircuitsTreeview.tag_configure('Research', font=('Helvetica', 8), background='pink')

                   
                    # Bind the double Click
                    self.CircuitsTreeview.bind('<ButtonRelease-1>', self.on_Circuits_Tree_select_click) # When Select the Tree
                    #self.CircuitsTreeview.bind("<Double-1>", self.On_Circuits_Tree_Refresh)         

                    #------------------ TREE VIEW For Circuits Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.CircuitsFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 40)

                    self.ComboBoxFacilityID = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 50)
                    self.ComboBoxFacilityID.bind("<<ComboboxSelected>>", self.on_facility_combo_changed)
                    self.ComboBoxFacilityID.place(x = 110, y = 70)

                    self.ComboBoxSitesID = ttk.Combobox(self.CircuitsFrame, state='disabled', width = 50)
                    self.ComboBoxSitesID.bind("<<ComboboxSelected>>", self.on_sites_combo_changed)
                    self.ComboBoxSitesID.place(x = 110, y = 100)

                    # Get the Type and Models
                    self.Get_Type_PortSpeed_and_Satus()
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
                    self.CircuitsWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.CircuitsFrame,title='Circuits',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Circuits Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    Circuits = Class_Circuits("BV",Windows_Scaling,location)
    Circuits.Display_Circuits_Window()


if __name__ == '__main__':
    Main()

