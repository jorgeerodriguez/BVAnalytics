#==========================================================================
# Program: FindCircuitsEquipmentFacilities.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Feb-17-2018
# Date Last Modified: Feb-18-2018
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
    from Device import *
    Is_Device_Available = True
    Location = []
    Device = Class_Device(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Devices Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Device_Available = False  

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
    Parameter = ['Find Circuit,Equipment,Facilities','OPEN Window']    
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

class Class_FindCircuitsEquipmentFacilities:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPSummaryWindowExist = False
        self.FindCircuitWindowExist = False
        self.FindDeviceWindowExist = False
        self.CircuitsWindowExist = False
        self.FindFacilityWindowExist = False
        self.Username = os.getlogin()
        self.date = ""
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
        
                    
#*********************************************************************************************************************************************
#                                   Circuits Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************

    def on_FindCircuitWindow_quit(self):
        if (self.FindCircuitWindowExist):
            self.FindCircuitWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Circuts','CLOSE Window']    
                Logging.Log(Parameter)
            self.FindCircuitWindow.destroy()

    def Call_Circuits_ButtonDevice(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Device Window']    
            Logging.Log(Parameter)
        if (Is_Device_Available):
            Device = Class_Device(ODBC_DSN_name,Windows_Scaling,self.location)
            Device.Display_Device_Window()

    def Call_Circuits_ButtonCircuit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Circuit Window']    
            Logging.Log(Parameter)
        if (Is_Circuit_Available):
            Circuits = Class_Circuits(ODBC_DSN_name,Windows_Scaling,self.location)
            Circuits.Display_Circuits_Window()

    def Call_Circuits_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Facility Window']    
            Logging.Log(Parameter)
        if (Is_Facility_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_Circuits_ButtonPOC(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Local POC Window']    
            Logging.Log(Parameter)
        if (Is_LocalPointOfContacts_Available):
            LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,self.location)
            LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.CircuitsTreeview.focus()
        dic = self.CircuitsTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.CircuitsTreeview.next(curItem)
        self.CircuitsTreeview.selection_set((NextItem, NextItem))
        self.CircuitsTreeview.focus(NextItem)

        NextItem = self.CircuitsTreeview.focus()
        dic = self.CircuitsTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
        
            
    def Call_Circuits_ButtonValidate(self):
        #self.CircuitsTreeview.focus_set() Need to researh more on this one
        i = 0
        Total = len(self.CircuitsTablePriaryKeyArray)
        self.Device_Exist_Array = [] # YES or NO
        self.Circuit_Exist_Array = [] # YES or NO
        while (i < Total):
            sql = """
                    SELECT * FROM Devices
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                    """ % (self.CircuitsTableCountryIDArray[i],self.CircuitsTableRegionIDArray[i],self.CircuitsTableFacilityIDArray[i])
            if (self.db.Execute(sql)):
                self.Device_Exist_Array.append("YES")
            else:
                self.Device_Exist_Array.append("NO")
            i = i + 1

        #------------------------------- Deleting Treeview Begin --------
        x = self.CircuitsTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                self.CircuitsTreeview.delete(child)
        #------------------------------- Deleting Treeview End ----------
                                
        #-------------- Initializing Arrays <BEGIN> ----------------------
        i = 0
        while (i < len(self.CircuitsTablePriaryKeyArray)):
            num = i + 1
            tags = self.Device_Exist_Array[i] # To use in the futire
            '''
                self.CircuitsTreeviewDataColumns = ('Circuits ID','Country','Facility','Carrier ID','Type','Speed','BW',
                                                        'CE ASN','CE IP','VLAN','Monthly Cost','ETF','Installed',
                                                        'Activated','Disconected','Status','Contract No','Expiration')
            '''
            item = [
                    self.CircuitsTablePriaryKeyArray[i],
                    #self.CircuitsTableDescriptionArray[i],
                    self.CircuitsTableCountryIDArray[i],
                    self.CircuitsTableFacilityIDArray[i],
                    self.CircuitsTableCarrierIDArray[i],
                    self.CircuitsTableCircuitTypeArray[i],
                    self.CircuitsTablePortSpeedArray[i],
                    self.CircuitsTableBandwidthArray[i],
                    self.CircuitsTableCE_ASNArray[i],
                    self.CircuitsTableCE_IP_AddressArray[i],
                    #self.CircuitsTablePE_ASNArray[i],
                    #self.CircuitsTablePE_IP_AddressArray[i],
                    self.CircuitsTableVLAN_IDArray[i],
                    self.CircuitsTableMonthlyCostArray[i],
                    self.CircuitsTableETFArray[i],                            
                    self.CircuitsTableDateInstalledArray[i],
                    self.CircuitsTableDateActivatedArray[i],
                    self.CircuitsTableDisconectedDateArray[i],
                    self.CircuitsTableStatusArray[i],
                    self.CircuitsTableContract_NoArray[i],
                    self.CircuitsTableExpirationDateArray[i],
                    #self.CircuitsTableExecutedByArray[i]
                    ]
            self.CircuitsTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
            i = i + 1

        self.buttonFacility['state'] = DISABLED
        self.buttonDevice['state'] = DISABLED
        self.buttonCircuit['state'] = DISABLED
        self.buttonValidate['state'] = DISABLED
        self.buttonPOC['state'] = DISABLED
        


    def on_Circuits_Tree_select_click(self,event):
            curItem = self.CircuitsTreeview.focus()  
            dic = self.CircuitsTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []
                self.location = [self.CircuitsTableCountryIDArray[curItem],self.CircuitsTableRegionIDArray[curItem],
                            self.CircuitsTableFacilityIDArray[curItem],self.CircuitsTableSiteIDArray[curItem]]
                if (Is_Facility_Available):
                    self.buttonFacility['state'] = ACTIVE
                if (Is_Device_Available):
                    self.buttonDevice['state'] = ACTIVE
                if (Is_Circuit_Available):
                    self.buttonCircuit['state'] = ACTIVE
                if ((Is_Device_Available) and (Is_Circuit_Available)):
                    self.buttonValidate['state'] = ACTIVE
                if (Is_LocalPointOfContacts_Available):
                    self.buttonPOC['state'] = ACTIVE


    def on_Tree_Circuit_double_click(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Circuit Window']    
            Logging.Log(Parameter)
            curItem = self.CircuitsTreeview.focus()  
            dic = self.CircuitsTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []
                self.location = [self.CircuitsTableCountryIDArray[curItem],self.CircuitsTableRegionIDArray[curItem],
                            self.CircuitsTableFacilityIDArray[curItem],self.CircuitsTableSiteIDArray[curItem]]
                if (Is_Circuit_Available):
                    Circuits = Class_Circuits(ODBC_DSN_name,Windows_Scaling,self.location)
                    Circuits.Display_Circuits_Window()

    def Call_ButtonFindCircuits(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Circuts','Find Circuits']    
            Logging.Log(Parameter)
        self.buttonFacility['state'] = DISABLED
        self.buttonDevice['state'] = DISABLED
        self.buttonCircuit['state'] = DISABLED
        self.buttonValidate['state'] = DISABLED
        self.buttonPOC['state'] = DISABLED
        
        #print ("Execute SQL")
        value = self.FindCircuitFrameEntry.get() # Value to querry
        index = (self.ComboBoxCircuitTableFields.current())
        #field = ComboBoxCircuitTableFields.get())
        field = (self.CircuitTableFieldsSQL[index])
        # I need to work on this querry today!!!!
        #data = Build_fake_data()
        #Populate_tree(FindCircuitTreeview,data)
        self.TotalTMC = 0
        self.TotalETF = 0
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == '*'):
                        sql = """SELECT * FROM CIRCUITS 
                                ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Status """                    
                    else:
                        sql = """SELECT * FROM CIRCUITS 
                                WHERE %s LIKE '%s' ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Status """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Treeview Begin --------
                        x = self.CircuitsTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                self.CircuitsTreeview.delete(child)
                        #------------------------------- Deleting Treeview End ----------
                                
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

                        i = 0
                        self.data_ready = True
                        while (i < len(self.db.results)):                    
                            self.CircuitsTablePriaryKeyArray.append(self.db.results[i][0].strip())
                            if (self.db.results[i][1] == None):
                                self.CircuitsTableDescriptionArray.append("")
                            else:
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
                            self.TotalTMC = self.TotalTMC + self.CircuitsTableMonthlyCostArray[i]
                            self.TotalETF = self.TotalETF + self.CircuitsTableETFArray[i]
                            tags = self.CircuitsTableStatusArray[i] # To use in the futire
                            '''
                                self.CircuitsTreeviewDataColumns = ('Circuits ID','Country','Facility','Carrier ID','Type','Speed','BW',
                                                                        'CE ASN','CE IP','VLAN','Monthly Cost','ETF','Installed',
                                                                        'Activated','Disconected','Status','Contract No','Expiration')
                            '''
                            item = [
                                    self.CircuitsTablePriaryKeyArray[i],
                                    #self.CircuitsTableDescriptionArray[i],
                                    self.CircuitsTableCountryIDArray[i],
                                    self.CircuitsTableFacilityIDArray[i],
                                    self.CircuitsTableCarrierIDArray[i],
                                    self.CircuitsTableCircuitTypeArray[i],
                                    self.CircuitsTablePortSpeedArray[i],
                                    self.CircuitsTableBandwidthArray[i],
                                    self.CircuitsTableCE_ASNArray[i],
                                    self.CircuitsTableCE_IP_AddressArray[i],
                                    #self.CircuitsTablePE_ASNArray[i],
                                    #self.CircuitsTablePE_IP_AddressArray[i],
                                    self.CircuitsTableVLAN_IDArray[i],
                                    self.CircuitsTableMonthlyCostArray[i],
                                    self.CircuitsTableETFArray[i],                            
                                    self.CircuitsTableDateInstalledArray[i],
                                    self.CircuitsTableDateActivatedArray[i],
                                    self.CircuitsTableDisconectedDateArray[i],
                                    self.CircuitsTableStatusArray[i],
                                    self.CircuitsTableContract_NoArray[i],
                                    self.CircuitsTableExpirationDateArray[i],
                                    #self.CircuitsTableExecutedByArray[i]
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
                        tmc= ("{:,.2f}".format(self.TotalTMC));
                        self.FindCircuitFrameEntryTMC['state'] = 'normal'
                        self.FindCircuitFrameEntryTMC.delete(0,END)
                        self.FindCircuitFrameEntryTMC.insert(0,tmc)
                        self.FindCircuitFrameEntryTMC['state'] = 'readonly'

                        etf= ("{:,.2f}".format(self.TotalETF));
                        self.FindCircuitFrameEntryETF['state'] = 'normal'
                        self.FindCircuitFrameEntryETF.delete(0,END)
                        self.FindCircuitFrameEntryETF.insert(0,etf)
                        self.FindCircuitFrameEntryETF['state'] = 'readonly'                        
                    else:
                        mbox.showerror(master=self.FindCircuitFrame,title='Find Circuits',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindCircuitFrame,title='Find Circuits',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindCircuitFrame,title='Find Circuits',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindCircuits_From_Entry(self,event):
        self.Call_ButtonFindCircuits()

    def FindCircuit_Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Circuits List"
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
                    mbox.showinfo(master=self.FindCircuitFrame,title='Find Circuits List',
                            message = '!!! The File was saved !!!')
                    self.FindCircuitWindow.title("Find Circuits                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Vendor','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindCircuitFrame,title='Find Circuits List',
                            message = '*** The File was not saved, Perhaps It is already open ***')



    def Find_Circuits_Window(self):
        if not self.FindCircuitWindowExist:
            # Set up the Window
            self.FindCircuitWindowExist = True
            self.FindCircuitWindow = Tk()
            self.FindCircuitWindow.geometry('1350x700+350+70')
            self.FindCircuitWindow.title("Find Circuits")
            self.FindCircuitWindow.protocol("WM_DELETE_WINDOW", self.on_FindCircuitWindow_quit)
            self.FindCircuitWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindCircuitLabel = Label(self.FindCircuitWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindCircuitLabel["text"] = "Find Circuits"
            self.FindCircuitLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindCircuitWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.FindCircuit_Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindCircuitWindow_quit)
            self.FindCircuitWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.FindCircuitFrame = Frame(self.FindCircuitWindow)
            self.FindCircuitFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindCircuitFrame.rowconfigure(0, weight=1)
            self.FindCircuitFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindCircuitFrameLabel = Label(self.FindCircuitFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindCircuitFrameLabel["text"] = "Find:"
            self.FindCircuitFrameLabel.place(x=10, y=10)

            # Setup ComboBox
            self.CircuitTableFields = ['Circuit ID',
                        'Description',
                        'Country ID',                    
                        'Region ID',
                        'Facility ID',
                        #'Site_ID',
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
                        #'PE_Switch',
                        #'PE_Location',
                        #'NPA_NXX',
                        #'Monthly_Cost',
                        'Order Number',
                        'Date Installed',
                        #'Day_Installed',
                        #'Month_Installed',
                        #'Year_Installed',
                        'Date Activated',
                        #'Day_Activated',
                        #'Month_Activated',
                        #'Year_Activated',
                        'Disconnect Date',
                        #'Day_Disconnect',
                        #'Month_Disconnect',
                        #'Year_Disconnect',
                        'Expiration Date',
                        #'Day_Expiration',
                        #'Month_Expiration',
                        #'Year_Expiration',
                        'Term Day',
                        'Term Time',
                        #'ETF',
                        'Contract_No',
                        'Account_No']


            self.CircuitTableFieldsSQL = ['Circuit_ID',
                        'Description',
                        'Country_ID',                    
                        'Region_ID',
                        'Facility_ID',
                        #'Site_ID',
                        'Carrier_ID',
                        'Circuit_Type',
                        'Port_Speed',
                        'Bandwidth',
                        'Status',
                        'Dmarc_Info_1',
                        'Dmarc_Info_2',
                        'LEC1',
                        'LEC2',
                        'LEC3',
                        'LEC4',
                        'LEC5',
                        'CE_ASN',
                        'CE_IP_Address',
                        'PE_ASN',
                        'PE_IP_Address',
                        'VLAN_ID',
                        #'PE_Switch',
                        #'PE_Location',
                        #'NPA_NXX',
                        #'Monthly_Cost',
                        'Order_Number',
                        'Date_Installed',
                        #'Day_Installed',
                        #'Month_Installed',
                        #'Year_Installed',
                        'Date_Activated',
                        #'Day_Activated',
                        #'Month_Activated',
                        #'Year_Activated',
                        'Disconnect_Date',
                        #'Day_Disconnect',
                        #'Month_Disconnect',
                        #'Year_Disconnect',
                        'Expiration_Date',
                        #'Day_Expiration',
                        #'Month_Expiration',
                        #'Year_Expiration',
                        'Term_Day',
                        'Term_Time',
                        #'ETF',
                        'Contract_No',
                        'Account_No']

            self.ComboBoxCircuitTableFields = ttk.Combobox(self.FindCircuitFrame, state='readonly', width = 25)
            self.ComboBoxCircuitTableFields['values'] = self.CircuitTableFields
            self.ComboBoxCircuitTableFields.current(0)
            self.ComboBoxCircuitTableFields.place(x = 50, y = 10)
            #ComboBoxButtonWEBOSName['state'] = DISABLED

            # Setup Entry
            self.FindCircuitFrameEntry = Entry(self.FindCircuitFrame)
            self.FindCircuitFrameEntry['width']=50
            self.FindCircuitFrameEntry.place(x=235, y=10)
            self.FindCircuitFrameEntry.bind('<Return>', self.Call_ButtonFindCircuits_From_Entry)
            #FindCircuitFrameEntry['state'] = DISABLED

            # TOTALS
            self.FindCircuitFrameLabelTMC = Label(self.FindCircuitFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindCircuitFrameLabelTMC["text"] = "Total Monthly Cost:"
            self.FindCircuitFrameLabelTMC.place(x=10, y=40)

            self.FindCircuitFrameEntryTMC = Entry(self.FindCircuitFrame)
            self.FindCircuitFrameEntryTMC['width']=20
            self.FindCircuitFrameEntryTMC.place(x=135, y=40)
            self.FindCircuitFrameEntryTMC['state'] = 'readonly'


            self.FindCircuitFrameLabelETF = Label(self.FindCircuitFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindCircuitFrameLabelETF["text"] = "Total ETF Cost:"
            self.FindCircuitFrameLabelETF.place(x=280, y=40)

            self.FindCircuitFrameEntryETF = Entry(self.FindCircuitFrame)
            self.FindCircuitFrameEntryETF['width']=20
            self.FindCircuitFrameEntryETF.place(x=380, y=40)
            self.FindCircuitFrameEntryETF['state'] = 'readonly'

            self.TotalTMC = 0
            self.TotalETF = 0

            if self.db.Connect():                
                # Setup Button
                self.buttonFind = Button(self.FindCircuitFrame, text = 'Execute', command = self.Call_ButtonFindCircuits, state=ACTIVE)
                self.buttonFind.place(x = 550, y = 7, width=75, height=25)

                self.buttonPOC = Button(self.FindCircuitFrame, text = 'Local Point Of Contacts', command = self.Call_Circuits_ButtonPOC, state=DISABLED)
                self.buttonPOC.place(x = 690, y = 7, width=140, height=25)

                self.buttonCircuit = Button(self.FindCircuitFrame, text = 'Circuits', command = self.Call_Circuits_ButtonCircuit, state=DISABLED)
                self.buttonCircuit.place(x = 850, y = 7, width=75, height=25)

                self.buttonDevice = Button(self.FindCircuitFrame, text = 'Devices', command = self.Call_Circuits_ButtonDevice, state=DISABLED)
                self.buttonDevice.place(x = 950, y = 7, width=75, height=25)

                self.buttonFacility = Button(self.FindCircuitFrame, text = 'Facility', command = self.Call_Circuits_ButtonFacility, state=DISABLED)
                self.buttonFacility.place(x = 1050, y = 7, width=75, height=25)

                self.buttonValidate = Button(self.FindCircuitFrame, text = 'Validate HW', command = self.Call_Circuits_ButtonValidate, state=DISABLED)
                self.buttonValidate.place(x = 1150, y = 7, width=100, height=25)

                #------------------ TREE VIEW For Circuits Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars                            
                self.CircuitsTreeviewDataColumns = ('Circuits ID','Country','Facility','Carrier ID','Type','Speed','BW',
                                                       'CE ASN','CE IP','VLAN No','Monthly Cost','ETF','Installed',
                                                       'Activated','Disconected','Status','Contract No','Expiration')
                    
                self.CircuitsTreeview = ttk.Treeview(self.FindCircuitFrame,columns=self.CircuitsTreeviewDataColumns, height=28) # <--- Make sure the frame is correct !!
                self.CircuitsTreeviewysb = Scrollbar(self.FindCircuitFrame,orient=VERTICAL, command=self.CircuitsTreeview.yview)           # <--- Make sure the frame is correct !!
                self.CircuitsTreeviewxsb = Scrollbar(self.FindCircuitFrame,orient=HORIZONTAL, command=self.CircuitsTreeview.xview)         # <--- Make sure the frame is correct !!
                self.CircuitsTreeview['yscroll'] = self.CircuitsTreeviewysb.set
                self.CircuitsTreeview['xscroll'] = self.CircuitsTreeviewxsb.set

                # setup headings and column 
                self.CircuitsTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.CircuitsTreeview.heading('#1', text='Circuits ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Circuits ID', False)) # E for East and W for West
                self.CircuitsTreeview.heading('#2', text='Country',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Country', False)) # E for East and W for West
                self.CircuitsTreeview.heading('#3', text='Facility',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Facility', False)) # E for East and W for West                
                self.CircuitsTreeview.heading('#4', text='Carrier ID',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Carrier ID', False)) # E for East and W for West
                self.CircuitsTreeview.heading('#5', text='Type', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Type', False)) # E for East and W for West
                self.CircuitsTreeview.heading('#6', text='Speed', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Speed', False)) # E for East and W for West
                self.CircuitsTreeview.heading('#7', text='BW', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'BW', False))
                self.CircuitsTreeview.heading('#8', text='CE ASN', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'CE ASN', False))
                self.CircuitsTreeview.heading('#9', text='CE IP', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'CE IP', False))
                #self.CircuitsTreeview.heading('#10', text='PE ASN', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'PE ASN', False))
                #self.CircuitsTreeview.heading('#11', text='PE IP', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'PE IP', False))
                self.CircuitsTreeview.heading('#10', text='VLAN No', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'VLAN No', False))
                self.CircuitsTreeview.heading('#11', text='Monthly Cost', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Monthly Cost', False))
                self.CircuitsTreeview.heading('#12', text='ETF', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'ETF', False))
                self.CircuitsTreeview.heading('#13', text='Installed', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Installed', False))
                self.CircuitsTreeview.heading('#14', text='Activated', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Activated', False))
                self.CircuitsTreeview.heading('#15', text='Disconected', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Disconected', False))
                self.CircuitsTreeview.heading('#16', text='Status', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Status', False))
                self.CircuitsTreeview.heading('#17', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Contract No', False))
                self.CircuitsTreeview.heading('#18', text='Expiration', anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'Expiration', False))
                #self.CircuitsTreeview.heading('#19', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.CircuitsTreeview, 'xxxxx', False)) 
                    
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
                self.CircuitsTreeview.column('#16', stretch=1, width=10)
                self.CircuitsTreeview.column('#17', stretch=1, width=10)
                self.CircuitsTreeview.column('#18', stretch=1, width=10)
                #self.CircuitsTreeview.column('#19', stretch=1, width=10)

                # add tree and scrollbars to frame
                self.CircuitsTreeview.grid(row=1, column=0, sticky=NSEW)
                self.CircuitsTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.CircuitsTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.
                self.CircuitsTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                #self.CircuitsTreeview.tag_configure('Active', font=('Helvetica', 8), background='green2')
                #self.CircuitsTreeview.tag_configure('Inactive', font=('Helvetica', 8), background='red')
                self.CircuitsTreeview.tag_configure('Inactive', font=('Helvetica', 8), background='ivory3')
                self.CircuitsTreeview.tag_configure('In Process', font=('Helvetica', 8), background='green3')
                self.CircuitsTreeview.tag_configure('Billed', font=('Helvetica', 8), background='yellow2')
                self.CircuitsTreeview.tag_configure('Research', font=('Helvetica', 8), background='pink')
                self.CircuitsTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.CircuitsTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                

                   
                # Bind the double Click
                self.CircuitsTreeview.bind('<ButtonRelease-1>', self.on_Circuits_Tree_select_click) # When Select the Tree
                self.CircuitsTreeview.bind("<Double-1>", self.on_Tree_Circuit_double_click)         

                #------------------ TREE VIEW For Circuits Database <END> -----------------------------------
                
                # Bind the double Click
                #self.FindCircuitTreeview.bind('<ButtonRelease-1>', self.on_ICMP_Tree_Device_select_click)
                #self.FindCircuitTreeview.bind("<Double-1>", self.on_Tree_Circuit_double_click)
                
                self.FindCircuitWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find Circuits',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")

#*********************************************************************************************************************************************
#                                   Circuits Section <END>                                                                                   *
#*********************************************************************************************************************************************

#*********************************************************************************************************************************************
#                                   Devices Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************


    def on_FindDeviceWindow_quit(self):
        if (self.FindDeviceWindowExist):
            self.FindDeviceWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Devices','CLOSE Window']    
                Logging.Log(Parameter)
            self.FindDeviceWindow.destroy()

    def Call_Devices_ButtonDevice(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Device Window']    
            Logging.Log(Parameter)
        if (Is_Device_Available):
            Device = Class_Device(ODBC_DSN_name,Windows_Scaling,self.location)
            Device.Display_Device_Window()

    def Call_Devices_ButtonCircuit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Circuit Window']    
            Logging.Log(Parameter)
        if (Is_Circuit_Available):
            Devices = Class_Circuits(ODBC_DSN_name,Windows_Scaling,self.location)
            Devices.Display_Circuits_Window()

    def Call_Devices_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Facility Window']    
            Logging.Log(Parameter)
        if (Is_Facility_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_Device_ButtonPOC(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Local POC Window']    
            Logging.Log(Parameter)
        if (Is_LocalPointOfContacts_Available):
            LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,self.location)
            LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.DeviceTreeview.focus()
        dic = self.DeviceTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.DeviceTreeview.next(curItem)
        self.DeviceTreeview.selection_set((NextItem, NextItem))
        self.DeviceTreeview.focus(NextItem)

        NextItem = self.DeviceTreeview.focus()
        dic = self.DeviceTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))

    def Call_Devices_ButtonOutsourceCost(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Update Outsource Cost']    
            Logging.Log(Parameter)
        # Set up the Window
        self.UpdateOutsourceCostWindowExist = True
        self.UpdateOutsourceCostWindow = Tk()
        self.UpdateOutsourceCostWindow.geometry('550x100+350+70')
        self.UpdateOutsourceCostWindow.title("Update Per Country Outsource Cost")
        #self.UpdateOutsourceCostWindow.protocol("WM_DELETE_WINDOW", self.on_FindDeviceWindow_quit)
        self.UpdateOutsourceCostWindow.call('tk', 'scaling', Windows_Scaling)
        self.UpdateOutsourceCostLabel = Label(self.UpdateOutsourceCostWindow,text="Helvetica", font=("Helvetica", 13))
        self.UpdateOutsourceCostLabel["text"] = "Update Per Country Outsource Cost"
        self.UpdateOutsourceCostLabel.pack()
                
        # Setup Frame
        self.UpdateOutsourceCostFrame = Frame(self.UpdateOutsourceCostWindow)
        self.UpdateOutsourceCostFrame.pack(side=TOP, fill=BOTH, expand=Y)

        # set frame resizing priorities
        self.UpdateOutsourceCostFrame.rowconfigure(0, weight=1)
        self.UpdateOutsourceCostFrame.columnconfigure(0, weight=1)

        # Create Progress Bar
        self.progress = ttk.Progressbar(self.UpdateOutsourceCostFrame, orient="horizontal",length=500, mode="determinate")
        Total = len(self.DeviceTablePriaryKeyArray)
        self.progress['maximum'] = Total
        self.progress['value'] = 0
        self.progress.place(x=10, y=10)
        if (mbox.askyesnocancel(master=self.UpdateOutsourceCostFrame,title='Update the Outsource Cost per Country per Device Type',message = 'Are you Sure you want to Proceed?')):
            Total = len(self.DeviceTablePriaryKeyArray)
            self.progress['maximum'] = Total
            self.progress['value'] = 0
            self.progress.place(x=10, y=10)
            i = 0
            while (i < Total):
                self.progress['value'] = i
                #sleep(0.1)
                sql = """
                        SELECT * FROM Device_Model
                        WHERE Device_Model_ID = '%s'
                      """ % (self.DeviceTableDeviceModelIDArray[i])
                if (self.db.Execute(sql)):
                    self.Device_Type_Outsource = self.db.results[0][3]
                    #print ("found Device Model")
                    Primarykey = self.DeviceTableDeviceModelIDArray[i] + "-" + self.DeviceTableCountryIDArray[i]
                    sql = """
                            SELECT * FROM Device_Model_Country
                            WHERE Device_Model_ID_Country_ID = '%s'
                          """ % (Primarykey)
                    if (self.db.Execute(sql)):
                        self.Device_Type_Outsource = self.db.results[0][3]
                        #print ("found Device Model per Country " + Primarykey)
                else:
                    self.Device_Type_Outsource = 0.0
                    
                sql = """
                        SELECT * FROM Devices
                        WHERE Device_Model_ID = '%s' AND Country_ID = '%s' AND Device_ID = '%s'
                      """ % (self.DeviceTableDeviceModelIDArray[i],self.DeviceTableCountryIDArray[i],self.DeviceTablePriaryKeyArray[i])
                if (self.db.Execute(sql)):
                    if (self.DeviceTableStatusArray[i] == 'obsolete'):
                        self.Device_Type_Outsource = 0.0
                    sql = "UPDATE Devices SET OutSource_Cost = '%f', Executed_by_UserID = '%s' \
                                    WHERE (Device_Model_ID  = '%s' AND Country_ID = '%s' AND Device_ID = '%s')" %(self.Device_Type_Outsource,self.Username,
                                                                self.DeviceTableDeviceModelIDArray[i],self.DeviceTableCountryIDArray[i],self.DeviceTablePriaryKeyArray[i])
                    if (self.db.Add_Move_Change_Data(sql)):
                        #print ("YES")
                        self.progress['value'] = i
                i = i + 1
                #print (self.Device_Type_Outsource)
            self.progress['value'] = Total
            self.Call_ButtonFindDevices()
        else:
            self.UpdateOutsourceCostWindow.destroy()
        #self.UpdateOutsourceCostWindow.mainloop()
        
            
    def Call_Devices_ButtonValidate(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Validate Circuits']    
            Logging.Log(Parameter)
        #self.DeviceTreeview.focus_set() Need to researh more on this one
        i = 0
        Total = len(self.DeviceTablePriaryKeyArray)
        self.Device_Exist_Array = [] # YES or NO
        #self.Device_Exist_Array = [] # YES or NO
        while (i < Total):
            sql = """
                    SELECT * FROM Circuits
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                    """ % (self.DeviceTableCountryIDArray[i],self.DeviceTableRegionIDArray[i],self.DeviceTableFacilityIDArray[i])
            if (self.db.Execute(sql)):
                self.Device_Exist_Array.append("YES")
            else:
                self.Device_Exist_Array.append("NO")
            i = i + 1

        #------------------------------- Deleting Treeview Begin --------
        x = self.DeviceTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                self.DeviceTreeview.delete(child)
        #------------------------------- Deleting Treeview End ----------
                                
        #-------------- Initializing Arrays <BEGIN> ----------------------
        i = 0
        '''
        self.DeviceTreeviewDataColumns = ('Device Name','Country','Facility','Type','Model','IP Address',
                                            'Outsrc. Cost','Mnt. Cost','Installed',
                                            'Activated','Disconected','Status','Contract No','Expiration',
                                            'Serial No.')

        '''
        while (i < len(self.DeviceTablePriaryKeyArray)):
            num = i + 1
            tags = self.Device_Exist_Array[i] # To use in the future
            item = [
                    self.DeviceTablePriaryKeyArray[i],
                    #self.DeviceTableDeviceDescriptionArray[i],
                    self.DeviceTableCountryIDArray[i],
                    self.DeviceTableFacilityIDArray[i],
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
                    #self.DeviceTableExecutedByArray[i],
                    ]
            self.DeviceTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
            i = i + 1
            
        self.buttonDeviceFacility['state'] = DISABLED
        self.buttonDeviceDevices['state'] = DISABLED
        self.buttonDeviceCircuits['state'] = DISABLED
        self.buttonValidateDevices['state'] = DISABLED
        self.buttonDevicePOC['state'] = DISABLED
        self.buttonUpdateOutsourceCost['state'] = DISABLED


    def on_Devices_Tree_select_click(self,event):
            curItem = self.DeviceTreeview.focus()  
            dic = self.DeviceTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []
                self.location = [self.DeviceTableCountryIDArray[curItem],self.DeviceTableRegionIDArray[curItem],
                            self.DeviceTableFacilityIDArray[curItem],self.DeviceTableSiteIDArray[curItem]]
                if (Is_Facility_Available):
                    self.buttonDeviceFacility['state'] = ACTIVE
                if (Is_Device_Available):
                    self.buttonDeviceDevices['state'] = ACTIVE
                if (Is_Circuit_Available):
                    self.buttonDeviceCircuits['state'] = ACTIVE
                if ((Is_Device_Available) and (Is_Circuit_Available)):
                    self.buttonValidateDevices['state'] = ACTIVE
                if (Is_LocalPointOfContacts_Available):
                    self.buttonDevicePOC['state'] = ACTIVE
                self.buttonUpdateOutsourceCost['state'] = ACTIVE

    def on_Tree_Device_double_click(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Device Window']    
            Logging.Log(Parameter)
        curItem = self.DeviceTreeview.focus()  
        dic = self.DeviceTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            self.location = []
            self.location = [self.DeviceTableCountryIDArray[curItem],self.DeviceTableRegionIDArray[curItem],
                            self.DeviceTableFacilityIDArray[curItem],self.DeviceTableSiteIDArray[curItem]]
            if (Is_Device_Available):
                Device = Class_Device(ODBC_DSN_name,Windows_Scaling,self.location)
                Device.Display_Device_Window()


    def Call_ButtonFindDevices(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Devices','Find Devices']    
            Logging.Log(Parameter)
        self.buttonDeviceFacility['state'] = DISABLED
        self.buttonDeviceDevices['state'] = DISABLED
        self.buttonDeviceCircuits['state'] = DISABLED
        self.buttonValidateDevices['state'] = DISABLED
        self.buttonDevicePOC['state'] = DISABLED
        self.buttonUpdateOutsourceCost['state'] = DISABLED
        
        #print ("Execute SQL")
        value = self.FindDeviceFrameEntry.get() # Value to querry
        index = (self.ComboBoxDeviceTableFields.current())
        #field = ComboBoxDeviceTableFields.get())
        field = (self.DeviceTableFieldsSQL[index])
        # I need to work on this querry today!!!!
        #data = Build_fake_data()
        #Populate_tree(FindDeviceTreeview,data)
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == '*'):
                        sql = """SELECT * FROM Devices 
                                ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Status """
                    else:
                        sql = """SELECT * FROM Devices 
                                WHERE %s LIKE '%s' ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Status """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Treeview Begin --------
                        x = self.DeviceTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                self.DeviceTreeview.delete(child)
                        #------------------------------- Deleting Treeview End ----------
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
                            'Last ICMP_Status',
                            'ICMP Capable',
                            'Monitor via ICMP',
                            'Last Updated CMDB_Date',
                            'Last Updated CMDB_Day',
                            'Last Updated CMDB_Month',
                            'Last Updated CMDB_Year'
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
                            34 MAC_Address             CHAR(20))"""

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
                                self.DeviceTableMonitorviaICMPArray.append(self.db.results[i][41].strip())

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
                        '''
                        self.DeviceTreeviewDataColumns = ('Device Name','Country','Facility','Type','Model','IP Address',
                                                            'Outsrc. Cost','Mnt. Cost','Installed',
                                                           'Activated','Disconected','Status','Contract No','Expiration',
                                                           'Serial No.')

                        '''
                        self.TotalOutsource = 0
                        self.TotalMaintenance = 0
                        while (i < len(self.DeviceTablePriaryKeyArray)):
                            self.TotalOutsource = self.TotalOutsource + self.DeviceTableOutsourceCostArray[i]
                            self.TotalMaintenance = self.TotalMaintenance + self.DeviceTableMaintenanceCostArray[i]                           
                            num = i + 1
                            tags = self.DeviceTableStatusArray[i] # To use in the futire
                            item = [
                                    self.DeviceTablePriaryKeyArray[i],
                                    #self.DeviceTableDeviceDescriptionArray[i],
                                    self.DeviceTableCountryIDArray[i],
                                    self.DeviceTableFacilityIDArray[i],
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
                                    #self.DeviceTableExecutedByArray[i],
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

                        Outsource= ("{:,.2f}".format(self.TotalOutsource));
                        self.FindDeviceFrameEntryOutsource['state'] = 'normal'
                        self.FindDeviceFrameEntryOutsource.delete(0,END)
                        self.FindDeviceFrameEntryOutsource.insert(0,Outsource)
                        self.FindDeviceFrameEntryOutsource['state'] = 'readonly'

                        Maintenance= ("{:,.2f}".format(self.TotalMaintenance));
                        self.FindDeviceFrameEntryMaintenance['state'] = 'normal'
                        self.FindDeviceFrameEntryMaintenance.delete(0,END)
                        self.FindDeviceFrameEntryMaintenance.insert(0,Maintenance)
                        self.FindDeviceFrameEntryMaintenance['state'] = 'readonly'                        

                    else:
                        mbox.showerror(master=self.FindDeviceFrame,title='Find Equipment',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindDeviceFrame,title='Find Equipment',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindDeviceFrame,title='Find Equipment',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindDevices_From_Entry(self,event):
        self.Call_ButtonFindDevices()

    def FindDevices_Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Find Equipment List"
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
                    mbox.showinfo(master=self.FindDeviceFrame,title='Find Equipment List',
                            message = '!!! The File was saved !!!')
                    self.FindDeviceWindow.title("Find Equipment                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Find Equipment','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindDeviceFrame,title='Find Equipment List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Find_Devices_Window(self):
        if not self.FindDeviceWindowExist:
            # Set up the Window
            self.FindDeviceWindowExist = True
            self.FindDeviceWindow = Tk()
            self.FindDeviceWindow.geometry('1350x700+350+70')
            self.FindDeviceWindow.title("Find Equipment")
            self.FindDeviceWindow.protocol("WM_DELETE_WINDOW", self.on_FindDeviceWindow_quit)
            self.FindDeviceWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindDeviceLabel = Label(self.FindDeviceWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindDeviceLabel["text"] = "Find Equipment"
            self.FindDeviceLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindDeviceWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.FindDevices_Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindDeviceWindow_quit)
            self.FindDeviceWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.FindDeviceFrame = Frame(self.FindDeviceWindow)
            self.FindDeviceFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindDeviceFrame.rowconfigure(0, weight=1)
            self.FindDeviceFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindDeviceFrameLabel = Label(self.FindDeviceFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindDeviceFrameLabel["text"] = "Find:"
            self.FindDeviceFrameLabel.place(x=10, y=10)

            # Setup ComboBox
            self.DeviceTableFields = ['Device ID',
                                'Description',
                                'Country ID',                    
                                'Region ID',
                                'Facility ID',
                                'Site ID',
                                'Device Type ID',
                                'Device Model ID',
                                'IP Address',
                                #'IP6_Address',
                                'Contract No',
                                'Expiration Date',
                                #'OutSource_Cost',
                                #'Maintenance_Cost',
                                'Status',
                                #'Monthly_Cost',
                                'Date Installed',
                                #'Day_Installed',
                                #'Month_Installed',
                                #'Year_Installed',
                                'Date Activated',
                                #'Day_Activated',
                                #'Month_Activated',
                                #'Year_Activated',
                                'Disconnect Date',
                                #'Day_Disconnect',
                                #'Month_Disconnect',
                                #'Year_Disconnect',
                                #'Day_Expiration',
                                #'Month_Expiration',
                                #'Year_Expiration',
                                'Serial No',
                                #'Executed_by_UserID',
                                'Notes',
                                'MAC Address',
                                'NAT IP Address',
                                'Management IP4 Address',
                                'Last Success ICMP',
                                'Last ICMP Status',
                                'ICMP Capable',
                                'Monitor via ICMP',
                                'Last Updated CMDB Date'                                         
                                ]

            self.DeviceTableFieldsSQL = ['Device_ID',
                                'Device_Description',
                                'Country_ID',                    
                                'Region_ID',
                                'Facility_ID',
                                'Site_ID',
                                'Device_Type_ID',
                                'Device_Model_ID',
                                'IP4_Address',
                                #'IP6_Address',
                                'Contract_No',
                                'Expiration_Date',
                                #'OutSource_Cost',
                                #'Maintenance_Cost',
                                'Status',
                                #'Monthly_Cost',
                                'Date_Installed',
                                #'Day_Installed',
                                #'Month_Installed',
                                #'Year_Installed',
                                'Date_Activated',
                                #'Day_Activated',
                                #'Month_Activated',
                                #'Year_Activated',
                                'Disconnect_Date',
                                #'Day_Disconnect,
                                #'Month_Disconnect',
                                #'Year_Disconnect',
                                #'Day_Expiration',
                                #'Month_Expiration',
                                #'Year_Expiration',
                                'Serial_No',
                                #'Executed_by_UserID',
                                'Notes',
                                'MAC_Address',
                                'NAT_IP4_Address',
                                'Management_IP4_Address',
                                'Last_Success_ICMP',
                                'Last_ICMP_Status',
                                'ICMP_Capable',
                                'Monitor_via_ICMP',
                                'Last_Updated_CMDB_Date'                                         
                                ]

            self.ComboBoxDeviceTableFields = ttk.Combobox(self.FindDeviceFrame, state='readonly', width = 25)
            self.ComboBoxDeviceTableFields['values'] = self.DeviceTableFields
            self.ComboBoxDeviceTableFields.current(0)
            self.ComboBoxDeviceTableFields.place(x = 50, y = 10)
            #ComboBoxButtonWEBOSName['state'] = DISABLED

            # Setup Entry
            self.FindDeviceFrameEntry = Entry(self.FindDeviceFrame)
            self.FindDeviceFrameEntry['width']=50
            self.FindDeviceFrameEntry.place(x=235, y=10)
            self.FindDeviceFrameEntry.bind('<Return>', self.Call_ButtonFindDevices_From_Entry)
            #FindDeviceFrameEntry['state'] = DISABLED


            # TOTALS
            self.FindDeviceFrameLabelOutsource = Label(self.FindDeviceFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindDeviceFrameLabelOutsource["text"] = "Montly Outsrc Cost:"
            self.FindDeviceFrameLabelOutsource.place(x=10, y=40)

            self.FindDeviceFrameEntryOutsource = Entry(self.FindDeviceFrame)
            self.FindDeviceFrameEntryOutsource['width']=20
            self.FindDeviceFrameEntryOutsource.place(x=145, y=40)
            self.FindDeviceFrameEntryOutsource['state'] = 'readonly'


            self.FindDeviceFrameLabelMaintenance = Label(self.FindDeviceFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindDeviceFrameLabelMaintenance["text"] = "Maintenance Cost:"
            self.FindDeviceFrameLabelMaintenance.place(x=290, y=40)

            self.FindDeviceFrameEntryMaintenance = Entry(self.FindDeviceFrame)
            self.FindDeviceFrameEntryMaintenance['width']=20
            self.FindDeviceFrameEntryMaintenance.place(x=405, y=40)
            self.FindDeviceFrameEntryMaintenance['state'] = 'readonly'

            self.TotalOutsource = 0
            self.TotalMaintenance = 0

            if self.db.Connect():                
                # Setup Button
                self.buttonFindDevice = Button(self.FindDeviceFrame, text = 'Execute', command = self.Call_ButtonFindDevices, state=ACTIVE)
                self.buttonFindDevice.place(x = 550, y = 7, width=75, height=25)

                self.buttonDevicePOC = Button(self.FindDeviceFrame, text = 'Local Point Of Contacts', command = self.Call_Device_ButtonPOC, state=DISABLED)
                self.buttonDevicePOC.place(x = 690, y = 7, width=140, height=25)

                self.buttonDeviceCircuits = Button(self.FindDeviceFrame, text = 'Circuits', command = self.Call_Devices_ButtonCircuit, state=DISABLED)
                self.buttonDeviceCircuits.place(x = 850, y = 7, width=75, height=25)

                self.buttonDeviceDevices = Button(self.FindDeviceFrame, text = 'Devices', command = self.Call_Devices_ButtonDevice, state=DISABLED)
                self.buttonDeviceDevices.place(x = 950, y = 7, width=75, height=25)

                self.buttonDeviceFacility = Button(self.FindDeviceFrame, text = 'Facility', command = self.Call_Devices_ButtonFacility, state=DISABLED)
                self.buttonDeviceFacility.place(x = 1050, y = 7, width=75, height=25)

                self.buttonValidateDevices = Button(self.FindDeviceFrame, text = 'Validate Circuits', command = self.Call_Devices_ButtonValidate, state=DISABLED)
                self.buttonValidateDevices.place(x = 1150, y = 7, width=100, height=25)

                self.buttonUpdateOutsourceCost = Button(self.FindDeviceFrame, text = 'Update Outsource Cost Per Country', command = self.Call_Devices_ButtonOutsourceCost, state=DISABLED)
                self.buttonUpdateOutsourceCost.place(x = 1050, y = 37, width=200, height=25)


                #------------------ TREE VIEW For Device Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars                            
                self.DeviceTreeviewDataColumns = ('Device Name','Country','Facility','Type','Model','IP Address',
                                                    'OutSrc. Cost','Mnt. Cost','Installed',
                                                   'Activated','Disconected','Status','Contract No','Expiration',
                                                   'Serial No.')
                    
                self.DeviceTreeview = ttk.Treeview(self.FindDeviceFrame,columns=self.DeviceTreeviewDataColumns, height=28) # <--- Make sure the frame is correct !!
                self.DeviceTreeviewysb = Scrollbar(self.FindDeviceFrame,orient=VERTICAL, command=self.DeviceTreeview.yview)           # <--- Make sure the frame is correct !!
                self.DeviceTreeviewxsb = Scrollbar(self.FindDeviceFrame,orient=HORIZONTAL, command=self.DeviceTreeview.xview)         # <--- Make sure the frame is correct !!
                self.DeviceTreeview['yscroll'] = self.DeviceTreeviewysb.set
                self.DeviceTreeview['xscroll'] = self.DeviceTreeviewxsb.set

                # setup headings and column 
                self.DeviceTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.DeviceTreeview.heading('#1', text='Device Name',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Device Name', False)) # E for East and W for West
                self.DeviceTreeview.heading('#2', text='Country',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Country', False)) # E for East and W for West
                self.DeviceTreeview.heading('#3', text='Facility',   anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Facility', False)) # E for East and W for West                    
                self.DeviceTreeview.heading('#4', text='Type', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Type', False)) # E for East and W for West
                self.DeviceTreeview.heading('#5', text='Model', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Model', False)) # E for East and W for West
                self.DeviceTreeview.heading('#6', text='IP Address', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'IP Address', False))
                self.DeviceTreeview.heading('#7', text='OutSrc. Cost', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'OutSrc. Cost', False))
                self.DeviceTreeview.heading('#8', text='Mnt. Cost', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Mnt. Cost', False))
                self.DeviceTreeview.heading('#9', text='Installed', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Installed', False))
                self.DeviceTreeview.heading('#10', text='Activated', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Activated', False))
                self.DeviceTreeview.heading('#11', text='Disconected', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Disconected', False))
                self.DeviceTreeview.heading('#12', text='Status', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Status', False))
                self.DeviceTreeview.heading('#13', text='Contract No', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Contract No', False))
                self.DeviceTreeview.heading('#14', text='Expiration', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Expiration', False))
                self.DeviceTreeview.heading('#15', text='Serial No.', anchor=W,command=lambda: self.treeview_sort_column(self.DeviceTreeview, 'Serial No.', False))
                
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
                self.DeviceTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.DeviceTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                   
                # Bind the double Click
                self.DeviceTreeview.bind('<ButtonRelease-1>', self.on_Devices_Tree_select_click) # When Select the Tree
                self.DeviceTreeview.bind("<Double-1>", self.on_Tree_Device_double_click)         

                #------------------ TREE VIEW For Device Database <END> -----------------------------------                
                self.FindDeviceWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find Equipment',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")

#*********************************************************************************************************************************************
#                                   Devices Section <END>                                                                                   *
#*********************************************************************************************************************************************


#*********************************************************************************************************************************************
#                                   Facility Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************



    def on_FindFacilityWindow_quit(self):
        if (self.FindFacilityWindowExist):
            self.FindFacilityWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Facility','CLOSE Window']    
                Logging.Log(Parameter)
            self.FindFacilityWindow.destroy()

    def Call_Facility_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Facility Window']    
            Logging.Log(Parameter)
        if (Is_Facility_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_Facility_ButtonCircuit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Circuit Window']    
            Logging.Log(Parameter)
        if (Is_Circuit_Available):
            Facility = Class_Circuits(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Circuits_Window()

    def Call_Facility_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Facility Window']    
            Logging.Log(Parameter)
        if (Is_Facility_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_Facility_ButtonPOC(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Local POC Window']    
            Logging.Log(Parameter)
        if (Is_LocalPointOfContacts_Available):
            LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,self.location)
            LocalPointOfContacts.Display_LocalPointOfContacts_Window()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.FacilityTreeview.focus()
        dic = self.FacilityTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.FacilityTreeview.next(curItem)
        self.FacilityTreeview.selection_set((NextItem, NextItem))
        self.FacilityTreeview.focus(NextItem)

        NextItem = self.FacilityTreeview.focus()
        dic = self.FacilityTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
        
            
    def Call_Facility_ButtonValidate(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Validate Circuits and Devices']    
            Logging.Log(Parameter)
        #self.FacilityTreeview.focus_set() Need to researh more on this one
        i = 0
        Total = len(self.FacilityTablePriaryKeyArray)
        self.Facility_HW_Exist_Array = [] # YES or NO
        self.Facility_Circuit_Exist_Array = [] # YES or NO
        self.FacilityValidateSelected = "YES"
        while (i < Total):
            sql = """
                    SELECT * FROM Circuits
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                    """ % (self.FacilityTableCountryIDArray[i],self.FacilityTableRegionIDArray[i],self.FacilityTableFacilityIDArray[i])
            if (self.db.Execute(sql)):
                self.Facility_Circuit_Exist_Array.append("YES")
            else:
                self.Facility_Circuit_Exist_Array.append("NO")

            sql = """
                    SELECT * FROM Devices
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                    """ % (self.FacilityTableCountryIDArray[i],self.FacilityTableRegionIDArray[i],self.FacilityTableFacilityIDArray[i])
            if (self.db.Execute(sql)):
                self.Facility_HW_Exist_Array.append("YES")
            else:
                self.Facility_HW_Exist_Array.append("NO")                
            i = i + 1

        #------------------------------- Deleting Treeview Begin --------
        x = self.FacilityTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                self.FacilityTreeview.delete(child)
        #------------------------------- Deleting Treeview End ----------
                                
        #-------------- Initializing Arrays <BEGIN> ----------------------
        i = 0
        while (i < len(self.FacilityTablePriaryKeyArray)):
            num = i + 1
            if (self.Facility_HW_Exist_Array[i] == "YES" and self.Facility_Circuit_Exist_Array[i] == "YES"):
                tags = "YES"
            else:
                tags = "NO"
            item = [
                    #self.FacilityTablePriaryKeyArray[i],
                    self.FacilityTableCountryIDArray[i], 
                    #self.FacilityTableRegionIDArray[i],
                    self.FacilityTableFacilityIDArray[i],
                    self.FacilityTableFacilityNameArray[i],
                    self.FacilityTableBusinessUnitArray[i],
                    self.FacilityTableOfficeTypeArray[i],
                    self.FacilityTableOfficeSizeArray[i],
                    self.FacilityTableFacilityLeaseEndNameArray[i],
                    #self.FacilityTableFacilityLeaseEndMonthArray[i],
                    #self.FacilityTableFacilityLeaseEndDayArray[i],
                    #self.FacilityTableFacilityLeaseEndYearArray[i],
                    self.FacilityTableFacilityOfficeAddress1Array[i],
                    #self.FacilityTableFacilityOfficeAddress2Array[i],
                    self.FacilityTableFacilityOfficeCityArray[i],
                    self.FacilityTableFacilityOfficeStateArray[i],
                    #self.FacilityTableFacilityOfficeZIPCodeArray[i],
                    #self.FacilityTableFacilityOfficeCountryArray[i],  
                    self.FacilityTableFacilityOfficeHeadcountArray[i],
                    #self.FacilityTableFacilityOfficeHRIDLocationArray[i],
                    #self.FacilityTableFacilityOfficeHRIDArray[i],
                    self.FacilityTableOfficeStatusArray[i],
                    #self.FacilityHasDevices[i],
                    #self.FacilityHasCircuits[i],
                    ]
            self.FacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
            i = i + 1

        self.buttonFacilityFacility['state'] = DISABLED
        self.buttonValidateFacility['state'] = DISABLED
        self.buttonFacilityPOC['state'] = DISABLED


    def on_Facility_Tree_select_click(self,event):
        curItem = self.FacilityTreeview.focus()  
        dic = self.FacilityTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            self.location = []              
            self.location = [self.FacilityTableCountryIDArray[curItem],self.FacilityTableRegionIDArray[curItem],
                            self.FacilityTableFacilityIDArray[curItem],self.FacilityValidateSelected]
            if (Is_Facility_Available):
                self.buttonFacilityFacility['state'] = ACTIVE
            if ((Is_Facility_Available) and (Is_Circuit_Available)):
                self.buttonValidateFacility['state'] = ACTIVE
            if (Is_LocalPointOfContacts_Available):
                self.buttonFacilityPOC['state'] = ACTIVE
                    

    def on_Tree_Facility_double_click(self,event):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Facility Window']    
            Logging.Log(Parameter)
        curItem = self.FacilityTreeview.focus()  
        dic = self.FacilityTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            self.location = []
            self.location = [self.FacilityTableCountryIDArray[curItem],self.FacilityTableRegionIDArray[curItem],
                            self.FacilityTableFacilityIDArray[curItem],self.FacilityValidateSelected]
            if (Is_Facility_Available):
                Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
                Facility.Display_Facility_Window()


    def Call_ButtonFindFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Facility','Find Facility']    
            Logging.Log(Parameter)
        self.buttonFacilityFacility['state'] = DISABLED
        self.buttonValidateFacility['state'] = DISABLED
        self.buttonFacilityPOC['state'] = DISABLED

        #print ("Execute SQL")
        value = self.FindFacilityFrameEntry.get() # Value to querry
        index = (self.ComboBoxFacilityTableFields.current())
        #field = ComboBoxFacilityTableFields.get())
        field = (self.FacilityTableFieldsSQL[index])
        self.FacilityValidateSelected = "NO"
        # I need to work on this querry today!!!!
        #data = Build_fake_data()
        #Populate_tree(FindFacilityTreeview,data)
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == '*'):
                        sql = """SELECT * FROM Facility 
                                ORDER BY Region_ID,Country_ID, Facility_ID,
                                Status """
                    else:
                        sql = """SELECT * FROM Facility 
                                WHERE %s LIKE '%s' ORDER BY Region_ID,Country_ID, Facility_ID,
                                Status """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Treeview Begin --------
                        x = self.FacilityTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                self.FacilityTreeview.delete(child)
                        #------------------------------- Deleting Treeview End ----------

                        #-------------- Initializing Arrays <BEGIN> ----------------------
                        self.FacilityTablePriaryKeyArray =  []
                        self.FacilityTableCountryIDArray =  [] 
                        self.FacilityTableRegionIDArray =  []
                        self.FacilityTableFacilityIDArray =  []
                        self.FacilityTableFacilityNameArray =  []
                        self.FacilityTableOfficeTypeArray =  []
                        self.FacilityTableOfficeSizeArray =  []
                        self.FacilityTableFacilityLeaseEndNameArray =  []
                        self.FacilityTableLaseNameArray =  []
                        self.FacilityTableFacilityLeaseEndMonthArray =  []
                        self.FacilityTableFacilityLeaseEndDayArray =  []
                        self.FacilityTableFacilityLeaseEndYearArray =  []
                        self.FacilityTableFacilityOfficeAddress1Array =  []
                        self.FacilityTableFacilityOfficeAddress2Array =  []
                        self.FacilityTableFacilityOfficeCityArray =  []             
                        self.FacilityTableFacilityOfficeStateArray =  []            
                        self.FacilityTableFacilityOfficeCountryArray =  []             
                        self.FacilityTableFacilityOfficeZIPCodeArray =  []
                        self.FacilityTableFacilityOfficeHeadcountArray =  []           
                        self.FacilityTableFacilityOfficeHRIDLocationArray = []
                        self.FacilityTableFacilityOfficeHRIDArray = []                
                        self.FacilityTableOfficeStatusArray =  []                               
                        self.FacilityTableBusinessUnitArray =  []
                        self.FacilityTableExecutedByUserArray = []
                        self.FacilityTableStrategyArray =  []
                        self.FacilityTableActionArray =  []
                        self.FacilityTableNotesArray =  []
                        self.FacilityTableFacilityNewOfficeAddress1Array =  []
                        self.FacilityTableFacilityNewOfficeAddress2Array =  []
                        self.FacilityTableFacilityNewOfficeCityArray =  []             
                        self.FacilityTableFacilityNewOfficeStateArray =  []            
                        self.FacilityTableFacilityNewOfficeCountryArray =  []             
                        self.FacilityTableFacilityNewOfficeZIPCodeArray =  []
                        self.FacilityHasDevices = []
                        self.FacilityHasCircuits = []
                        self.FacilityTableArrayColumns = []
                        self.FacilityTableArrayTemp = []
                        self.results = []


                        self.FacilityTableArrayColumns = (
                                'Country ID',
                                'Region ID',
                                'Facility ID',
                                'Facility Name',
                                'Business Unit ID',
                                'Office Type',
                                'Office Size',
                                'Lease End Date',
                                'Lease End Month',
                                'Lease End Day',
                                'Lease End Year',
                                'Address1',
                                'Address2',
                                'City',
                                'State',
                                'Zip Code',
                                'Country',
                                'Headcount',
                                'HR ID_Location',
                                'HR Office ID',
                                'Status',
                                'Executed by UserID',
                                'Strategy',
                                'Action',
                                'Notes',
                                'New Address1',
                                'New Address2',
                                'New City',
                                'New State',
                                'New Zip Code',
                                'New Country')

                        
                        #-------------- Initializing Arrays <END> ----------------------
                        self.data_ready = True
                        i = 0
                        while (i < len(self.db.results)): # I NEED TO WORK ON THIS JORGE fix all the fields and then the Select function need to build Array for all fields
                            self.FacilityTablePriaryKeyArray.append(self.db.results[i][0].strip())
                            self.FacilityTableCountryIDArray.append(self.db.results[i][1].strip()) 
                            self.FacilityTableRegionIDArray.append(self.db.results[i][2].strip())
                            self.FacilityTableFacilityIDArray.append(self.db.results[i][3].strip())
                            self.FacilityTableFacilityNameArray.append(self.db.results[i][4].strip())
                            self.FacilityTableBusinessUnitArray.append(self.db.results[i][5].strip())
                            self.FacilityTableOfficeTypeArray.append(self.db.results[i][6].strip())
                            self.FacilityTableOfficeSizeArray.append(self.db.results[i][7].strip())
                            if (self.db.results[i][8] == None):
                                self.FacilityTableFacilityLeaseEndNameArray.append("")
                            else:
                                self.FacilityTableFacilityLeaseEndNameArray.append(self.db.results[i][8].strip())
                            self.FacilityTableFacilityLeaseEndMonthArray.append(self.db.results[i][9])
                            self.FacilityTableFacilityLeaseEndDayArray.append(self.db.results[i][10])
                            self.FacilityTableFacilityLeaseEndYearArray.append(self.db.results[i][11])
                            self.FacilityTableFacilityOfficeAddress1Array.append(self.db.results[i][12].strip())
                            if (self.db.results[i][13] == None):
                                self.FacilityTableFacilityOfficeAddress2Array.append("")
                            else:
                                self.FacilityTableFacilityOfficeAddress2Array.append(self.db.results[i][13].strip())
                            self.FacilityTableFacilityOfficeCityArray.append(self.db.results[i][14].strip())
                            if (self.db.results[i][15] == None):
                                self.FacilityTableFacilityOfficeStateArray.append("")
                            else:
                                self.FacilityTableFacilityOfficeStateArray.append(self.db.results[i][15].strip())
                            self.FacilityTableFacilityOfficeZIPCodeArray.append(self.db.results[i][16].strip())
                            self.FacilityTableFacilityOfficeCountryArray.append(self.db.results[i][17].strip())  
                            self.FacilityTableFacilityOfficeHeadcountArray.append(self.db.results[i][18])
                            self.FacilityTableFacilityOfficeHRIDArray.append(self.db.results[i][19].strip())
                            self.FacilityTableFacilityOfficeHRIDLocationArray.append(self.db.results[i][20])
                            self.FacilityTableOfficeStatusArray.append(self.db.results[i][21].strip())        
                            self.FacilityTableExecutedByUserArray.append(self.db.results[i][22].strip())
                            if (self.db.results[i][23] == None):
                                self.FacilityTableStrategyArray.append("")
                            else:
                                self.FacilityTableStrategyArray.append(self.db.results[i][23].strip())
                            if (self.db.results[i][24] == None):
                                self.FacilityTableActionArray.append("")
                            else:
                                self.FacilityTableActionArray.append(self.db.results[i][24].strip())
                            if (self.db.results[i][25] == None):
                                self.FacilityTableNotesArray.append("")
                            else:
                                self.FacilityTableNotesArray.append(self.db.results[i][25].strip())
                            if (self.db.results[i][26] == None):
                                self.FacilityTableFacilityNewOfficeAddress1Array.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeAddress1Array.append(self.db.results[i][26].strip())
                            if (self.db.results[i][27] == None):
                                self.FacilityTableFacilityNewOfficeAddress2Array.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeAddress2Array.append(self.db.results[i][27].strip())
                            if (self.db.results[i][28] == None):
                                self.FacilityTableFacilityNewOfficeCityArray.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeCityArray.append(self.db.results[i][28].strip())
                            if (self.db.results[i][29] == None):
                                self.FacilityTableFacilityNewOfficeStateArray.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeStateArray.append(self.db.results[i][29].strip())
                            if (self.db.results[i][30] == None):
                                self.FacilityTableFacilityNewOfficeZIPCodeArray.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeZIPCodeArray.append(self.db.results[i][30].strip())
                            if (self.db.results[i][31] == None):
                                self.FacilityTableFacilityNewOfficeCountryArray.append("")
                            else:
                                self.FacilityTableFacilityNewOfficeCountryArray.append(self.db.results[i][31].strip())
                            Location = []
                            Location = [self.db.results[i][1].strip(),self.db.results[i][2].strip(),self.db.results[i][3].strip()]
                            i = i + 1

                        i = 0
                        '''
                        self.FacilityTreeviewDataColumns = ('Country ID','Facility ID','Facility Name','Business Unit','Office Type','Office Size','Lease End','Addres',
                                                    'City','State','Head Count','Status')
                        '''
                        while (i < len(self.FacilityTablePriaryKeyArray)):
                            num = i + 1
                            tags = self.FacilityTableOfficeStatusArray[i].replace(' ','_')
                            item = [
                                    #self.FacilityTablePriaryKeyArray[i],
                                    self.FacilityTableCountryIDArray[i], 
                                    #self.FacilityTableRegionIDArray[i],
                                    self.FacilityTableFacilityIDArray[i],
                                    self.FacilityTableFacilityNameArray[i],
                                    self.FacilityTableBusinessUnitArray[i],
                                    self.FacilityTableOfficeTypeArray[i],
                                    self.FacilityTableOfficeSizeArray[i],
                                    self.FacilityTableFacilityLeaseEndNameArray[i],
                                    #self.FacilityTableFacilityLeaseEndMonthArray[i],
                                    #self.FacilityTableFacilityLeaseEndDayArray[i],
                                    #self.FacilityTableFacilityLeaseEndYearArray[i],
                                    self.FacilityTableFacilityOfficeAddress1Array[i],
                                    #self.FacilityTableFacilityOfficeAddress2Array[i],
                                    self.FacilityTableFacilityOfficeCityArray[i],
                                    self.FacilityTableFacilityOfficeStateArray[i],
                                    #self.FacilityTableFacilityOfficeZIPCodeArray[i],
                                    #self.FacilityTableFacilityOfficeCountryArray[i],  
                                    self.FacilityTableFacilityOfficeHeadcountArray[i],
                                    #self.FacilityTableFacilityOfficeHRIDLocationArray[i],
                                    #self.FacilityTableFacilityOfficeHRIDArray[i],
                                    self.FacilityTableOfficeStatusArray[i],
                                    #self.FacilityHasDevices[i],
                                    #self.FacilityHasCircuits[i],
                                    ]
                            self.FacilityTableArrayTemp = [
                                    #self.FacilityTablePriaryKeyArray[i],
                                    self.FacilityTableCountryIDArray[i], 
                                    self.FacilityTableRegionIDArray[i],
                                    self.FacilityTableFacilityIDArray[i],
                                    self.FacilityTableFacilityNameArray[i],
                                    self.FacilityTableBusinessUnitArray[i],
                                    self.FacilityTableOfficeTypeArray[i],
                                    self.FacilityTableOfficeSizeArray[i],
                                    self.FacilityTableFacilityLeaseEndNameArray[i],
                                    self.FacilityTableFacilityLeaseEndMonthArray[i],
                                    self.FacilityTableFacilityLeaseEndDayArray[i],
                                    self.FacilityTableFacilityLeaseEndYearArray[i],
                                    self.FacilityTableFacilityOfficeAddress1Array[i],
                                    self.FacilityTableFacilityOfficeAddress2Array[i],
                                    self.FacilityTableFacilityOfficeCityArray[i],
                                    self.FacilityTableFacilityOfficeStateArray[i],
                                    self.FacilityTableFacilityOfficeZIPCodeArray[i],
                                    self.FacilityTableFacilityOfficeCountryArray[i],  
                                    self.FacilityTableFacilityOfficeHeadcountArray[i],
                                    self.FacilityTableFacilityOfficeHRIDLocationArray[i],
                                    self.FacilityTableFacilityOfficeHRIDArray[i],
                                    self.FacilityTableOfficeStatusArray[i],
                                    #self.FacilityHasDevices[i],
                                    #self.FacilityHasCircuits[i],
                                    self.FacilityTableExecutedByUserArray[i],
                                    self.FacilityTableStrategyArray[i],
                                    self.FacilityTableActionArray[i],
                                    self.FacilityTableNotesArray[i],
                                    self.FacilityTableFacilityNewOfficeAddress1Array[i],
                                    self.FacilityTableFacilityNewOfficeAddress2Array[i],
                                    self.FacilityTableFacilityNewOfficeCityArray[i],
                                    self.FacilityTableFacilityNewOfficeStateArray[i],
                                    self.FacilityTableFacilityNewOfficeZIPCodeArray[i],
                                    self.FacilityTableFacilityNewOfficeCountryArray[i]
                                    ]
                            self.results.append(self.FacilityTableArrayTemp)
                            self.FacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.FindFacilityFrame,title='Find Equipment',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindFacilityFrame,title='Find Equipment',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindFacilityFrame,title='Find Facility',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindFacility_From_Entry(self,event):
        self.Call_ButtonFindFacility()

    def FindFacility_Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Find Facility List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.FacilityTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.FindFacilityFrame,title='Find Facility List',
                            message = '!!! The File was saved !!!')
                    self.FindFacilityWindow.title("Find Facility                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Find Facility','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindFacilityFrame,title='Find Facility List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Find_Facility_Window(self):
        if not self.FindFacilityWindowExist:
            # Set up the Window
            self.FindFacilityWindowExist = True
            self.FindFacilityWindow = Tk()
            self.FindFacilityWindow.geometry('1350x700+350+70')
            self.FindFacilityWindow.title("Find Facility")
            self.FindFacilityWindow.protocol("WM_DELETE_WINDOW", self.on_FindFacilityWindow_quit)
            self.FindFacilityWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindFacilityLabel = Label(self.FindFacilityWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindFacilityLabel["text"] = "Find Facility"
            self.FindFacilityLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindFacilityWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.FindFacility_Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindFacilityWindow_quit)
            self.FindFacilityWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------

                
            # Setup Frame
            self.FindFacilityFrame = Frame(self.FindFacilityWindow)
            self.FindFacilityFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindFacilityFrame.rowconfigure(0, weight=1)
            self.FindFacilityFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindFacilityFrameLabel = Label(self.FindFacilityFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindFacilityFrameLabel["text"] = "Find:"
            self.FindFacilityFrameLabel.place(x=10, y=10)

            # Setup ComboBox
            self.FacilityTableFields = ['Country ID',
                                'Region ID',
                                'Facility ID',
                                'Facility Name',
                                'Business Unit ID',
                                'Office Type',
                                'Office Size',
                                'Lease End Date',
                                #'Lease End Month',
                                #'Lease End Day',
                                #'Lease End Year',
                                'Address1',
                                'Address2',
                                'City',
                                'State',
                                'Zip Code',
                                'Country',
                                'Headcount',
                                'HR ID Location',
                                'HR Office ID',
                                'Status',
                                'Strategy',
                                'Action',
                                'Notes']

            self.FacilityTableFieldsSQL = [
                                'Country_ID',
                                'Region_ID',
                                'Facility_ID',
                                'Facility_Name',
                                'Business_Unit_ID',
                                'Office_Type',
                                'Office_Size',
                                'Lease_End_Date',
                                #'Lease_End_Month',
                                #'Lease_End_Day',
                                #'Lease_End_Year',
                                'Address1',
                                'Address2',
                                'City',
                                'State',
                                'Zip_Code',
                                'Country',
                                'Headcount',
                                'HR_ID_Location',
                                'HR_Office_ID',
                                'Status',
                                'Strategy',
                                'Action',
                                'Notes']

            self.ComboBoxFacilityTableFields = ttk.Combobox(self.FindFacilityFrame, state='readonly', width = 25)
            self.ComboBoxFacilityTableFields['values'] = self.FacilityTableFields
            self.ComboBoxFacilityTableFields.current(0)
            self.ComboBoxFacilityTableFields.place(x = 50, y = 10)
            #ComboBoxButtonWEBOSName['state'] = DISABLED

            # Setup Entry
            self.FindFacilityFrameEntry = Entry(self.FindFacilityFrame)
            self.FindFacilityFrameEntry['width']=50
            self.FindFacilityFrameEntry.place(x=235, y=10)
            self.FindFacilityFrameEntry.bind('<Return>', self.Call_ButtonFindFacility_From_Entry)
            #FindFacilityFrameEntry['state'] = DISABLED

            if self.db.Connect():                
                # Setup Button
                self.buttonFindFacility = Button(self.FindFacilityFrame, text = 'Execute', command = self.Call_ButtonFindFacility, state=ACTIVE)
                self.buttonFindFacility.place(x = 550, y = 7, width=75, height=25)

                self.buttonFacilityPOC = Button(self.FindFacilityFrame, text = 'Local Point Of Contacts', command = self.Call_Facility_ButtonPOC, state=DISABLED)
                self.buttonFacilityPOC.place(x = 690, y = 7, width=140, height=25)

                self.buttonFacilityFacility = Button(self.FindFacilityFrame, text = 'Facility', command = self.Call_Facility_ButtonFacility, state=DISABLED)
                self.buttonFacilityFacility.place(x = 850, y = 7, width=75, height=25)

                self.buttonValidateFacility = Button(self.FindFacilityFrame, text = 'Validate HW & Circuits', command = self.Call_Facility_ButtonValidate, state=DISABLED)
                self.buttonValidateFacility.place(x = 950, y = 7, width=140, height=25)



                #------------------ TREE VIEW For Facility Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.FacilityTreeviewDataColumns = ('Country ID','Facility ID','Facility Name','Business Unit','Office Type','Office Size','Lease End','Address',
                                                    'City','State','Head Count','Status')
                    
                self.FacilityTreeview = ttk.Treeview(self.FindFacilityFrame,columns=self.FacilityTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.FacilityTreeviewysb = Scrollbar(self.FindFacilityFrame,orient=VERTICAL, command=self.FacilityTreeview.yview)           # <--- Make sure the frame is correct !!
                self.FacilityTreeviewxsb = Scrollbar(self.FindFacilityFrame,orient=HORIZONTAL, command=self.FacilityTreeview.xview)         # <--- Make sure the frame is correct !!
                self.FacilityTreeview['yscroll'] = self.FacilityTreeviewysb.set
                self.FacilityTreeview['xscroll'] = self.FacilityTreeviewxsb.set

                # setup headings and column 
                self.FacilityTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.FacilityTreeview.heading('#1', text='Country ID',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Country ID', False)) # E for East and W for West
                self.FacilityTreeview.heading('#2', text='Facility ID',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Facility ID', False)) # E for East and W for West
                self.FacilityTreeview.heading('#3', text='Facility Name', anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Facility Name', False)) # E for East and W for West
                self.FacilityTreeview.heading('#4', text='Business Unit', anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Business Unit', False)) 
                self.FacilityTreeview.heading('#5', text='Office Type',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Office Type', False)) 
                self.FacilityTreeview.heading('#6', text='Office Size',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Office Size', False)) 
                self.FacilityTreeview.heading('#7', text='Lease End',     anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Lease End', False))
                self.FacilityTreeview.heading('#8', text='Address',       anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Address', False))
                self.FacilityTreeview.heading('#9', text='City',          anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'City', False))
                self.FacilityTreeview.heading('#10', text='State',         anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'State', False))
                self.FacilityTreeview.heading('#11', text='Head Count',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Head Count', False)) 
                self.FacilityTreeview.heading('#12', text='Status',       anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Status', False))
                    
                self.FacilityTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.FacilityTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.FacilityTreeview.column('#2', stretch=1, width=10)
                self.FacilityTreeview.column('#3', stretch=1, width=10)
                self.FacilityTreeview.column('#4', stretch=1, width=10)
                self.FacilityTreeview.column('#5', stretch=1, width=10)
                self.FacilityTreeview.column('#6', stretch=1, width=10)
                self.FacilityTreeview.column('#7', stretch=1, width=10)
                self.FacilityTreeview.column('#8', stretch=1, width=10)
                self.FacilityTreeview.column('#9', stretch=1, width=10)
                self.FacilityTreeview.column('#10', stretch=1, width=10)
                self.FacilityTreeview.column('#11', stretch=1, width=10)
                self.FacilityTreeview.column('#12', stretch=1, width=10)

                # add tree and scrollbars to frame
                self.FacilityTreeview.grid(row=1, column=0, sticky=NSEW)
                self.FacilityTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.FacilityTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.

                self.FacilityTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.FacilityTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                self.FacilityTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.FacilityTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.FacilityTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.FacilityTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                #self.FacilityTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.FacilityTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.FacilityTreeview.tag_configure('Closed', font=('Helvetica', 8), background='ivory3') # 'red3'

                   
                # Bind the double Click
                self.FacilityTreeview.bind('<ButtonRelease-1>', self.on_Facility_Tree_select_click) # When Select the Tree
                self.FacilityTreeview.bind("<Double-1>", self.on_Tree_Facility_double_click)         
                
                #------------------ TREE VIEW For Facility Database <END> -----------------------------------                
                self.FindFacilityWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find Equipment',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   Facility Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Circuits Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindCEF = Class_FindCircuitsEquipmentFacilities("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindCEF.Find_Facility_Window()


if __name__ == '__main__':
    Main()
