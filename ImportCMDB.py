#==========================================================================
# Program: ImportCMDB.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: May-10-2017
# Date Last Modified: May-14-2017
# Summary: This is Class to for the ImportCMDB
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
    Parameter = ['ImportCMDB','OPEN Window']    
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

class Class_ImportCMDB:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.ImportCMDBWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.version = self.db.Get_Version()
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


    def on_ImportCMDBWindow_quit(self):
        if (self.ImportCMDBWindowExist):
            self.ImportCMDBWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['ImportCMDB','CLOSE Window']    
                Logging.Log(Parameter) 
            self.ImportCMDBWindow.destroy()

    def Call_Save_As_Sample_DXC(self):
        self.DXC_File_Columns = (
                    'name',
                    'parentequipment',
                    'managementIP',
                    'managementIPmask',
                    'realIPaddress',
                    'realIPmask',
                    'site',
                    'building-floor-room-tile',
                    'region/group',
                    'type',
                    'function/role',
                    'brand',
                    'model',
                    'serialnumber',
                    'detectedserialnumber',
                    'macaddress',
                    'step',
                    'creationdate',
                    'lastchangedate',
                    'movetoproddate',
                    'obsolescencedate')
            
        self.DXC_File_Sample = (
                    'bv-coppukrt1.net.eds.com',
                    'No Required',
                    '7.9.141.1',
                    '255.255.255.0',
                    '10.186.45.1',
                    '255.255.255.0',
                    'CoppermillsLaneWalthamstowLondon',
                    'TWCommsroom',
                    'APAC',
                    'Router',
                    'SiteRouter',
                    'Cisco',
                    'CISCO1921/K9(C1900series)',
                    'FCZ1608207E',
                    'FCZ1608207E',
                    'MAC Address if required',
                    'obsolete',
                    '8/29/201211:39:18PM',
                    '6/6/20172:43:24PM',
                    '8/29/201211:39:18PM',
                    '6/6/20172:43:24PM')
        
        self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Hardware"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.DXC_File_Columns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')                          #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                ExcellFile.Add_DataToWorksheet(self.DXC_File_Sample,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = '!!! The File was saved !!!')
                    self.ImportCMDBWindow.title("Import CMDB                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['ImportCMDB','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.ImportCMDBFrame,title='ImportCMDB',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Call_Save_As_Sample_Cisco(self):        
        self.Cisco_File_Columns =(
        'Contract Type',
        'Contract #',
        'Item Status',
        'Item Name',
        'Serial Number',
        'Qty',
        'Item Begin Date',
        'Item End Date',
        'Last Date of Support',
        'Ship Date',
        'Item Type',
        'Product Group',
        'Install Site ID',
        'Install Site Name',
        'Install Site Address',
        'Install Site Address 2',
        'City',
        'State',
        'Zipcode',
        'Country',
        'Maintenance Bill-To ID',
        'Maintenance Bill-To Name',
        'Maintenance SO#',
        'Maintenance PO#',
        'Instance ID',
        'Contract Status',
        'Currency Code')
        
        self.Cisco_File_Sample = (
        'SNT',
        '91311424',
        'EXPIRED',
        'CISCO1841',
        'FCZ133371MN',
        '1',
        '11/23/2012',
        '12/31/2013',
        '10/31/2016',
        '8/18/2009',
        'CHASSIS',
        '1800.ROUTER.1841 SERIES',
        '403832204',
        'BLACK & VEATCH LTD',
        'GROSVENOR HOUSE',
        '69 LONDON ROAD',
        'REDHILL',
        'SURREY',
        'RH1 1LQ',
        'UNITED KINGDOM',
        '1619354',
        'TALLGRASS TECHNOLOGIES',
        '10513060',
        '61827625',
        '486810440',
        'EXPIRED',
        ' ')
        
        self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Contract Summary Detail - BO"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.Cisco_File_Columns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')                          #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                ExcellFile.Add_DataToWorksheet(self.Cisco_File_Sample,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = '!!! The File was saved !!!')
                    self.ImportCMDBWindow.title("Import CMDB                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['ImportCMDB','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.ImportCMDBFrame,title='ImportCMDB',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Call_Button_ImportCMDB_Validate_File(self):
        if Is_logging_Available:
            Parameter = []
            if (self.Menu_Option == "DXC"):
                Parameter = ['ImportCMDB','DXC Validate']    
            if (self.Menu_Option == "Cisco"):
                Parameter = ['ImportCMDB','Cisco Validate']    
            Logging.Log(Parameter)
            
        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.ImportCMDBTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.ImportCMDBTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------


        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        No_of_Rows = len(self.ExcellFileData)
        No_of_Columns = len(self.ExcellFileData[0])
        Column_Names = self.ExcellFileData[0]
        #------------------------------------------------------------------> DXC <-----------------------------------------------------------------------
        if (self.Menu_Option == "DXC"):
            '''
            # Columns required on the DXC File
            0  name
            1  parentequipment
            2  managementIP
            3  managementIPmask
            4  realIPaddress
            5  realIPmask
            6  site
            7  building-floor-room-tile
            8  region/group
            9  type
            10 function/role
            11 brand
            12 model
            13 serialnumber
            14 detectedserialnumber
            15 macaddress
            16 step
            17 creationdate
            18 lastchangedate
            19 movetoproddate
            20 obsolescencedate
            '''
            self.DXC_File_Columns = (
                    'name',
                    'parentequipment',
                    'managementIP',
                    'managementIPmask',
                    'realIPaddress',
                    'realIPmask',
                    'site',
                    'building-floor-room-tile',
                    'region/group',
                    'type',
                    'function/role',
                    'brand',
                    'model',
                    'serialnumber',
                    'detectedserialnumber',
                    'macaddress',
                    'step',
                    'creationdate',
                    'lastchangedate',
                    'movetoproddate',
                    'obsolescencedate')
            
            self.DXC_File_Sample = (
                    'bv-coppukrt1.net.eds.com',
                    'No Required',
                    '7.9.141.1',
                    '255.255.255.0',
                    '10.186.45.1',
                    '255.255.255.0',
                    'CoppermillsLaneWalthamstowLondon',
                    'TWCommsroom',
                    'APAC',
                    'Router',
                    'SiteRouter',
                    'Cisco',
                    'CISCO1921/K9(C1900series)',
                    'FCZ1608207E',
                    'FCZ1608207E',
                    'MAC Address if required',
                    'obsolete',
                    '8/29/201211:39:18PM',
                    '6/6/20172:43:24PM',
                    '8/29/201211:39:18PM',
                    '6/6/20172:43:24PM')
            
            File_OK = False
            if (No_of_Columns == 21):
                i = 0
                while (i < No_of_Columns):
                    if (self.DXC_File_Columns[i] == Column_Names[i]):
                        File_OK = True
                        i = i + 1
                    else:
                        File_OK = False
                        i = No_of_Columns + 1
            if (File_OK):
                self.progress['value'] = 100
                self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                self.ButtonImportCMDBDXCImport['state'] = ACTIVE
                self.ExcellFile.Save_File()
                mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = '!!! The Excel File matched the Column names required to be validated !!!')
            else:
                mbox.showerror(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = ('*** The Excel File does not Match the Column names required to be validated *** \n\n' + str(self.DXC_File_Columns)))                                         

        #------------------------------------------------------------------> Cisco <-----------------------------------------------------------------------
        if (self.Menu_Option == "Cisco"):
            '''
            # Columns required on the Cisco File
            0 'Contract Type',
            1  'Contract #',
            2  'Item Status',
            3  'Item Name',
            4  'Serial Number',
            5  'Qty',
            6  'Item Begin Date',
            7  'Item End Date',
            8  'Last Date of Support',
            9  'Ship Date',
            10 'Item Type',
            11 'Product Group',
            12 'Install Site ID',
            13 'Install Site Name',
            14 'Install Site Address',
            15 'Install Site Address 2',
            16 'City',
            17 'State',
            18 'Zipcode',
            19 'Country',
            20 'Maintenance Bill-To ID',
            21 'Maintenance Bill-To Name',
            22 'Maintenance SO#',
            23 'Maintenance PO#',
            24 'Instance ID',
            25 'Contract Status',
            26 'Currency Code')
            '''
            self.Cisco_File_Columns =(
            'Contract Type',
            'Contract #',
            'Item Status',
            'Item Name',
            'Serial Number',
            'Qty',
            'Item Begin Date',
            'Item End Date',
            'Last Date of Support',
            'Ship Date',
            'Item Type',
            'Product Group',
            'Install Site ID',
            'Install Site Name',
            'Install Site Address',
            'Install Site Address 2',
            'City',
            'State',
            'Zipcode',
            'Country',
            'Maintenance Bill-To ID',
            'Maintenance Bill-To Name',
            'Maintenance SO#',
            'Maintenance PO#',
            'Instance ID',
            'Contract Status',
            'Currency Code')
            
            self.Cisco_File_Sample = (
            'SNT',
            '91311424',
            'EXPIRED',
            'CISCO1841',
            'FCZ133371MN',
            '1',
            '11/23/2012',
            '12/31/2013',
            '10/31/2016',
            '8/18/2009',
            'CHASSIS',
            '1800.ROUTER.1841 SERIES',
            '403832204',
            'BLACK & VEATCH LTD',
            'GROSVENOR HOUSE',
            '69 LONDON ROAD',
            'REDHILL',
            'SURREY',
            'RH1 1LQ',
            'UNITED KINGDOM',
            '1619354',
            'TALLGRASS TECHNOLOGIES',
            '10513060',
            '61827625',
            '486810440',
            'EXPIRED',
            ' ')
            
            File_OK = False
            if (No_of_Columns == 27):
                i = 0
                while (i < No_of_Columns):
                    if (self.Cisco_File_Columns[i] == Column_Names[i]):
                        File_OK = True
                        i = i + 1
                    else:
                        File_OK = False
                        i = No_of_Columns + 1
            if (File_OK):
                self.progress['value'] = 100
                self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                self.ButtonImportCMDBDXCImport['state'] = ACTIVE
                self.ExcellFile.Save_File()
                mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = '!!! The Excel File matched the Column names required to be validated !!!')
            else:
                mbox.showerror(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = ('*** The Excel File does not Match the Column names required to be validated *** \n\n' + str(self.Cisco_File_Columns)))                                         


    def Get_Device_Information(self,index):
            self.Device_ID   = self.db.results[index][0].strip()
            self.Country_ID  = self.db.results[index][2].strip()
            self.Region_ID   = self.db.results[index][3].strip()
            self.Facility_ID = self.db.results[index][4].strip()
            self.Site_ID     = self.db.results[index][5].strip()
            if (self.db.results[index][7] == None):
                self.IP4_Address = ""
            else:
                self.IP4_Address = self.db.results[index][7].strip()

            if (self.db.results[index][12] == None):
                self.Status = ""
            else:
                self.Status = self.db.results[index][12].strip()
                                    
            if (self.db.results[index][14] == None):
                self.Date_Installed = ""
            else:
                self.Date_Installed = self.db.results[index][14].strip()
                                    
            if (self.db.results[index][18] == None):
                self.Date_Activated = ""
            else:
                self.Date_Activated = self.db.results[index][18].strip()
                                    
            if (self.db.results[index][22] == None):
                self.Disconnect_Date = ""
            else:
                self.Disconnect_Date = self.db.results[index][22].strip()

            if (self.db.results[0][30] == None):
                self.Serial_No = ""
            else:
                self.Serial_No = self.db.results[index][30].strip()
                                
            if (self.db.results[index][33] == None):
                self.Device_Model_ID = ""
            else:
                self.Device_Model_ID = self.db.results[index][33].strip()
                                
            if (self.db.results[index][36] == None):
                self.Management_IP4_Address = ""
            else:
                self.Management_IP4_Address = self.db.results[index][36].strip()
                
            if (self.db.results[index][41] == None):
                self.Last_Updated_CMDB_Date = ""
            else:
                self.Last_Updated_CMDB_Date = self.db.results[index][41].strip()

            if (((self.db.results[index][26]) == None) or (self.db.results[index][27] == 0)):
                self.Expiration_Date  = ""
                self.Day_Expiration   = 0
                self.Month_Expiration = 0
                self.Year_Expiration  = 0
            else:
                self.Expiration_Date  = self.db.results[index][26].strip()
                self.Day_Expiration   = self.db.results[index][27]
                self.Month_Expiration = self.db.results[index][28]
                self.Year_Expiration  = self.db.results[index][29]


    def Add_Device_Table(self):
        #self.Get_Device_Information(0)
        sql = """
                SELECT * FROM Device_Model
                WHERE Device_Model_ID = '%s'
                """ % (self.DeviceModelID)
        if (self.db.Execute(sql)):
            self.DeviceTypeID = self.db.results[0][5].strip() # Device Type
            PrimaryKey = self.DeviceID
            #------------------ Need to Add a new Entry ---------------------------------------------
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
                    '%s','%s','%d','%d','%d')" %(PrimaryKey,self.DeviceDescription, self.Country_ID, self.Region_ID, self.Facility_ID, self.Site_ID,
                    self.DeviceTypeID, self.DeviceIPAddress, "::", self.DeviceContract, self.DeviceOutSourceCost, self.DeviceMaintenanceCost,
                    self.DeviceStatus, 0.0, self.DeviceInstalledDate, int(self.DeviceInstalledDay), int(self.DeviceInstalledMonth), int(self.DeviceInstalledYear),
                    self.DeviceActivatedDate, int(self.DeviceActivatedDay), int(self.DeviceActivatedMonth), int(self.DeviceActivatedYear),
                    self.DeviceDisconnectedDate, int(self.DeviceDisconnectedDay), int(self.DeviceDisconnectedMonth), int(self.DeviceDisconnectedYear),
                    self.DeviceExpirationDate, int(self.DeviceExpirationDay), int(self.DeviceExpirationMonth), int(self.DeviceExpirationYear),
                    self.DeviceSerialNo,self.Username, self.DeviceNotes, self.DeviceModelID, self.DeviceMACAddress, self.DeviceTableNATIP4Address,
                    self.DeviceManagementIP4Address, self.DeviceTableLastSuccessICMP, self.DeviceTableLastICMPStatus,
                    self.DeviceTableICMPCapable,self.DeviceTableMonitorviaICMP,self.DeviceTableLastUpdatedCMDBDate,int(self.DeviceTableLastUpdatedCMDBDay),
                    int(self.DeviceTableLastUpdatedCMDBMonth),int(self.DeviceTableLastUpdatedCMDBYear))
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Device Added to the Database:," + self.DeviceID + ","
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][0],self.row+1,0+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.ExcellFile.Add_DataToWorksheetSingleCell("Added to the Database",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
            else:
                self.Action = self.Action + "Device Was not Added to the Database:," + self.DeviceID + ","
                self.ExcellFile.Add_DataToWorksheetSingleCell("NOT Added to the Database",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
        else:
            self.Action = self.Action + "Device Model Not Found in Database, Need to add a NEW Device Model," + self.DeviceModelID + ","    
            self.ExcellFile.Add_DataToWorksheetSingleCell("Device Model Not Found in Database",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
        

    def Update_Device_Table(self,Field,PrimaryKey):
        # Format <Database != Excell File>
        #--------------------------------------------- Device ID -------------------------------
        if (self.Device_ID != self.DeviceID):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Device_ID = '%s', Device_Description = '%s' \
                    WHERE %s  = '%s'" %(self.Username, self.DeviceID, self.DeviceID, Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Device ID Changed from:," + self.Device_ID + ",to:," + self.DeviceID + ","
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][0],self.row+1,0+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Device ID Changed, was not updated in the Database,"

        #--------------------------------------------- IP Address -------------------------------
        if (self.IP4_Address != self.DeviceIPAddress):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', IP4_Address = '%s' \
                    WHERE %s  = '%s'" %(self.Username, self.DeviceIPAddress, Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "IP Address Changed,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][4],self.row+1,4+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "IP Address Changed, was not updated in the Database,"

        #--------------------------------------------- Status -------------------------------
        if (self.Status !=  self.DeviceStatus):
            if (self.Status == 'obsolete'):
                self.Action = self.Action + "Status Changed from obsolete, Need More Research"
            else:                                            
                sql = "UPDATE Devices SET Executed_by_UserID = '%s', Status = '%s' \
                        WHERE %s  = '%s'" %(self.Username, self.DeviceStatus, Field, PrimaryKey)
                if (self.db.Add_Move_Change_Data(sql)):
                    self.Action = self.Action + "Status Changed to:," + self.DeviceStatus + ","
                    self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][16],self.row+1,16+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                    self.Change_happened = True
                else:
                    self.Action = self.Action + "Status Changed, was not updated in the Database,"

        #--------------------------------------------- Serial No. -------------------------------
        if (self.Serial_No !=  self.DeviceSerialNo):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Serial_No = '%s' \
                    WHERE %s  = '%s'" %(self.Username, self.DeviceSerialNo, Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Serial No. Changed,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][4],self.row+1,4+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Serial No. Changed, was not updated in the Database,"

        #--------------------------------------------- Device Model  -------------------------------
        if ((self.Device_Model_ID !=  self.DeviceModelID) and (len(self.DeviceModelID) > 0)):
            sql = """
                    SELECT * FROM Device_Model
                    WHERE Device_Model_ID = '%s'
                    """ % (self.DeviceModelID)
            if (self.db.Execute(sql)):
                self.DeviceTypeID = self.db.results[0][5].strip() # Device Type
                sql = "UPDATE Devices SET Executed_by_UserID = '%s', Device_Model_ID = '%s', Device_Type_ID = '%s' \
                        WHERE %s  = '%s'" %(self.Username, self.DeviceModelID, self.DeviceTypeID, Field, PrimaryKey)
                if (self.db.Add_Move_Change_Data(sql)):
                    self.Action = self.Action + "Device Model Changed,"
                    self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][12],self.row+1,12+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                    self.Change_happened = True
                else:
                    self.Action = self.Action + "Device Model Changed, was not updated in the Database,"
            else:
                self.Action = self.Action + "Device Model Changed but Not Found in Database, Need to add a NEW Device Model," + self.DeviceModelID + ","
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][12],self.row+1,12+1,"Hardware","Fill",steelblue,11,'Bold')
                self.ExcellFile.Add_DataToWorksheetSingleCell("Device Model Not Found in Database",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
        #--------------------------------------------- Management IP Address  -------------------------------
        if (self.Management_IP4_Address !=  self.DeviceManagementIP4Address):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Management_IP4_Address = '%s' \
                    WHERE %s  = '%s'" %(self.Username, self.DeviceManagementIP4Address, Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Managed IP Address Changed,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][2],self.row+1,2+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Management IP Address Changed, was not updated in the Database,"

        #--------------------------------------------- Date Installed  -------------------------------
        if (self.Date_Installed !=  self.DeviceInstalledDate):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Date_Installed = '%s', Day_Installed = '%d', Month_Installed = '%d', Year_Installed = '%d' \
                    WHERE %s  = '%s'" %(self.Username,self.DeviceInstalledDate,int(self.DeviceInstalledDay),int(self.DeviceInstalledMonth),
                                                int(self.DeviceInstalledYear), Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Changed Installed Date,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][17],self.row+1,17+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Date Installed Changed, was not updated in the Database,"
           
        #--------------------------------------------- Date Activated  -------------------------------
        if (self.Date_Activated !=  self.DeviceActivatedDate):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Date_Activated = '%s', Day_Activated = '%d', Month_Activated = '%d', Year_Activated = '%d' \
                    WHERE %s  = '%s'" %(self.Username,self.DeviceActivatedDate,int(self.DeviceActivatedDay),int(self.DeviceActivatedMonth),
                                                int(self.DeviceActivatedYear), Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Changed Activated Date,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][19],self.row+1,19+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Date Activated Changed, was not updated in the Database,"

        #--------------------------------------------- Date Disconnected  -------------------------------
        
        if (self.Disconnect_Date !=  self.DeviceDisconnectedDate):
            sql = "UPDATE Devices SET Executed_by_UserID = '%s', Disconnect_Date = '%s', Day_Disconnect = '%d', Month_Disconnect = '%d', Year_Disconnect = '%d' \
                    WHERE %s  = '%s'" %(self.Username,self.DeviceDisconnectedDate,int(self.DeviceDisconnectedDay),int(self.DeviceDisconnectedMonth),
                                                int(self.DeviceDisconnectedYear), Field, PrimaryKey)
            if (self.db.Add_Move_Change_Data(sql)):
                self.Action = self.Action + "Changed Disconnected Date,"
                self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][20],self.row+1,20+1,"Hardware","Fill",lightsalmon_1,11,'Bold')
                self.Change_happened = True
            else:
                self.Action = self.Action + "Date Disconnected Changed, was not updated in the Database,"
                
        #--------------------------------------------- Updated CMDB Date  -------------------------------
        sql = "UPDATE Devices SET Executed_by_UserID = '%s', Last_Updated_CMDB_Date = '%s', Last_Updated_CMDB_Day = '%d', Last_Updated_CMDB_Month = '%d', \
                Last_Updated_CMDB_Year = '%d' \
                WHERE %s  = '%s'" %(self.Username,str(self.DeviceTableLastUpdatedCMDBDate), int(self.DeviceTableLastUpdatedCMDBDay),int(self.DeviceTableLastUpdatedCMDBMonth),
                                                                           int(self.DeviceTableLastUpdatedCMDBYear), Field, PrimaryKey)
        if (self.db.Add_Move_Change_Data(sql)):
            if (self.Change_happened ):
                self.Action = self.Action + "Inventory Updated,"
                self.ExcellFile.Add_DataToWorksheetSingleCell("Updated",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
            else:
                self.Action = self.Action + "No Changes,"
                self.ExcellFile.Add_DataToWorksheetSingleCell("No Changes",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
        else:
            self.Action = self.Action + "CMDB Date, was not updated in the Database,"

    def Conver_Device_Model_ID(self):
        if (str(self.ExcellFileData[self.row][12]) == "MerakiMR32"):
            self.ExcellFileData[self.row][12] = "Meraki MR32"

        if (str(self.ExcellFileData[self.row][12]) == "Meraki MR 33"):
            self.ExcellFileData[self.row][12] = "Meraki MR33"

        if (str(self.ExcellFileData[self.row][12]) == "MerakiMR33"):
            self.ExcellFileData[self.row][12] = "Meraki MR33"

        if (str(self.ExcellFileData[self.row][12]) == "MerakiMX65W"):
            self.ExcellFileData[self.row][12] = "Meraki MX65W"

        if (str(self.ExcellFileData[self.row][12]) == "MerakiMS225-48FP"):
            self.ExcellFileData[self.row][12] = "Meraki MS225-48FP"

        if (str(self.ExcellFileData[self.row][12]) == "UX1000W/SBA"):
            self.ExcellFileData[self.row][12] = "UX1000 W/SBA"

        if (str(self.ExcellFileData[self.row][12]) == "UX2000W/SBA"):
            self.ExcellFileData[self.row][12] = "UX2000 W/SBA"

        if (str(self.ExcellFileData[self.row][12]) == "CISCO1921/K9(C1900series)"):
            self.ExcellFileData[self.row][12] = "CISCO1921/K9 (C1900series)"

        if (str(self.ExcellFileData[self.row][12]) == "WS-C2924-XL-EN(C2900XLseries)"):
            self.ExcellFileData[self.row][12] = "WS-C2924M-XL"

        if (str(self.ExcellFileData[self.row][12]) == "WS-C2950-24(C2950series)"):
            self.ExcellFileData[self.row][12] = "WS-C2950-24"

        if (str(self.ExcellFileData[self.row][12]) == "E3500-48-PoE"):
            self.ExcellFileData[self.row][12] = "E3500yl-48G-PoE+"

        if (str(self.ExcellFileData[self.row][12]) == "1050(1050L)"):
            self.ExcellFileData[self.row][12] = "1050"

        if (str(self.ExcellFileData[self.row][12]) == "WS-C3750G-48PS-S(C3750series)"):
            self.ExcellFileData[self.row][12] = "WS-C3750G-48PS-S"

        if (str(self.ExcellFileData[self.row][12]) == "550(550H)"):
            self.ExcellFileData[self.row][12] = "550 (550H)"

        if (str(self.ExcellFileData[self.row][12]) == "CX3070"):
            self.ExcellFileData[self.row][12] = "CX3070 (CX3070L)"

        if (str(self.ExcellFileData[self.row][12]) == "MerakiMS225-24P"):
            self.ExcellFileData[self.row][12] = "Meraki MS225-24P"

        if (str(self.ExcellFileData[self.row][12]) == "Model#4260"):
            self.ExcellFileData[self.row][12] = "Model #4260"

        # Added it on Sep 6 2018
        if (str(self.ExcellFileData[self.row][12]) == "WS-C2960-X"):
            self.ExcellFileData[self.row][12] = "WS-C2960X-48LPD-L"

        if (str(self.ExcellFileData[self.row][12]) == "WS-C2960X"):
            self.ExcellFileData[self.row][12] = "WS-C2960X-48LPD-L"


        if (str(self.ExcellFileData[self.row][12]) == "2960X-48LPD-L"):
            self.ExcellFileData[self.row][12] = "WS-C2960X-48LPD-L"

        if (str(self.ExcellFileData[self.row][12]) == "WS-C2960X-48FPD-L (C2960X series)"):
            self.ExcellFileData[self.row][12] = "WS-C2960X-48FPD-L"

        if (str(self.ExcellFileData[self.row][12]) == "MS225-24P"):
            self.ExcellFileData[self.row][12] = "Meraki MS225-24P"



    def Call_Button_ImportCMDB_Import_File(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['ImportCMDB','DXC Import']    
            Logging.Log(Parameter)
        if (mbox.askyesnocancel(master=self.ImportCMDBFrame,title='Import CMDB',message = 'Are you Sure you want to IMPORT the Network Inventory?')):            
            if self.db.Connect():
                #-------------------------------------------------------------------> DXC <-------------------------------------------------------
                if (self.Menu_Option == "DXC"):
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    self.ButtonImportCMDBDXCImport['state'] = DISABLED        
                    No_of_Rows = len(self.ExcellFileData)
                    No_of_Columns = len(self.ExcellFileData[0])
                    Column_Names = self.ExcellFileData[0]
                    self.progress['maximum'] = No_of_Rows
                    self.progress['value'] = 0
                    column = 0
                    self.row = 1
                    i = 0
                    num = 0
                    self.results = []
                    self.progress['maximum'] = No_of_Rows
                    self.progress['value'] = 0
                    while (self.row < No_of_Rows):
                        #column = 0
                        #while (column <No_of_Columns):
                        #========================================== Here is all the Work to valildate the Device Database Each Column NEEDS to be EXACT =============
                        tags = "-" # To use in the future                          
                        '''
                        # Columns required on the DXC File
                        0  name
                        1  parentequipment
                        2  managementIP
                        3  managementIPmask
                        4  realIPaddress
                        5  realIPmask
                        6  site
                        7  building-floor-room-tile
                        8  region/group
                        9  type
                        10 function/role
                        11 brand
                        12 model
                        13 serialnumber
                        14 detectedserialnumber
                        15 macaddress
                        16 step
                        17 creationdate
                        18 lastchangedate
                        19 movetoproddate
                        20 obsolescencedate
                        '''

                        '''
                        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
                        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
                        self.SitesID = self.SitesIDArray[self.ComboBoxSitesID.current()]
                        ################# For the Cisco Maintenance price ##################
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
                        '''
                        
                        if ((len(self.ExcellFileData[self.row][0]) > 0) and (self.ExcellFileData[self.row][0].find(".") > 0)):
                            self.CountryID = "UNKNOWN"
                            self.RegionID = "UNKNOWN"
                            self.FacilityID = "UNKNOWN"
                            self.SitesID = "UNKNOWN"
                            device_id = self.ExcellFileData[self.row][0].split(".")
                            self.DeviceID = device_id[0]
                            self.DeviceDescription = device_id[0]
                            self.DeviceManagementIP4Address = self.ExcellFileData[self.row][2]
                            self.DeviceIPAddress = self.ExcellFileData[self.row][4]
                            if ((len(self.DeviceIPAddress) > 6) and (self.DeviceIPAddress.find(".") > 0)):
                                x = self.DeviceIPAddress.split('.')
                                self.network = x[0] + "." + x[1] + "." +x[2]+"."
                            else:
                                self.network = "0.0.0.0"
                            self.DeviceTypeID = self.ExcellFileData[self.row][9]
                            self.Conver_Device_Model_ID()
                            self.DeviceModelID = str(self.ExcellFileData[self.row][12])
                            if (len(self.DeviceModelID) > 0):
                                x = self.DeviceModelID.strip()
                                self.DeviceModelID = x
                            self.DeviceSerialNo = str(self.ExcellFileData[self.row][13])
                            if (len(self.DeviceSerialNo) > 100):
                                x = self.DeviceSerialNo[:99]
                                self.DeviceSerialNo = x
                            detectedserialnumber = self.ExcellFileData[self.row][14]
                            self.DeviceStatus  = self.ExcellFileData[self.row][16]           
                            self.DeviceContract = ""
                            self.DeviceMACAddress = self.ExcellFileData[self.row][15]
                            self.DeviceOutSourceCost = 0
                            self.DeviceMaintenanceCost = 0
                            self.DeviceNotes = "Added by Importing the File"
                            self.DeviceTableNATIP4Address = "0.0.0.0"           
                            self.DeviceTableICMPCapable = "YES"
                            self.DeviceTableMonitorviaICMP = "NO"
                            self.DeviceTableLastSuccessICMP = ""
                            self.DeviceTableLastICMPStatus = ""
                            today = datetime.date.today()
                            self.DeviceTableLastUpdatedCMDBDate = str(today.month) + "/" + str(today.day) + "/" +str(today.year)
                            self.DeviceTableLastUpdatedCMDBDay = today.day
                            self.DeviceTableLastUpdatedCMDBMonth = today.month
                            self.DeviceTableLastUpdatedCMDBYear = today.year
                            ################
                            if ((len(str(self.ExcellFileData[self.row][17])) > 0)and (str(self.ExcellFileData[self.row][17]) != "0000-00-00 00:00:00")): # <------------- Install Date
                                self.DeviceInstalledMonth = str(self.ExcellFileData[self.row][17].month)
                                self.DeviceInstalledDay   = str(self.ExcellFileData[self.row][17].day)
                                self.DeviceInstalledYear  = str(self.ExcellFileData[self.row][17].year)
                                self.DeviceInstalledDate = str(self.DeviceInstalledMonth) +"/"+str(self.DeviceInstalledDay) +"/"+ str(self.DeviceInstalledYear)
                            else:
                                self.DeviceInstalledMonth = "0"
                                self.DeviceInstalledDay = "0"
                                self.DeviceInstalledYear = "0"
                                self.DeviceInstalledDate = ""
                            '''
                            if ((len(self.ExcellFileData[self.row][17]) > 8) and (self.ExcellFileData[self.row][17].find("/") > 0) and (self.ExcellFileData[self.row][17].find(":") > 0)): 
                                date_and_time = self.ExcellFileData[self.row][17].split(":")
                                y = date_and_time[0]
                                date = y.split("/")
                                self.DeviceInstalledMonth = date[0]
                                self.DeviceInstalledDay = date[1]
                                self.DeviceInstalledYear = date[2][:4]
                                self.DeviceInstalledDate = str(self.DeviceInstalledMonth) +"/"+str(self.DeviceInstalledDay) +"/"+ str(self.DeviceInstalledYear)
                            else:
                                self.DeviceInstalledMonth = "0"
                                self.DeviceInstalledDay = "0"
                                self.DeviceInstalledYear = "0"
                                self.DeviceInstalledDate = ""
                            '''
                            ################
                            if ((len(str(self.ExcellFileData[self.row][19])) > 0) and (str(self.ExcellFileData[self.row][19]) != "0000-00-00 00:00:00")): # <------------- Activated Date
                                self.DeviceActivatedMonth = str(self.ExcellFileData[self.row][19].month)
                                self.DeviceActivatedDay   = str(self.ExcellFileData[self.row][19].day)
                                self.DeviceActivatedYear  = str(self.ExcellFileData[self.row][19].year)
                                self.DeviceActivatedDate = str(self.DeviceActivatedMonth) +"/"+str(self.DeviceActivatedDay) +"/"+ str(self.DeviceActivatedYear)
                            else:
                                self.DeviceActivatedMonth = "0"
                                self.DeviceActivatedDay = "0"
                                self.DeviceActivatedYear = "0"
                                self.DeviceActivatedDate = ""
                            '''
                            if ((len(self.ExcellFileData[self.row][19]) > 8) and (self.ExcellFileData[self.row][19].find("/") > 0) and (self.ExcellFileData[self.row][19].find(":") > 0)): 
                                date_and_time = self.ExcellFileData[self.row][19].split(":")
                                y = date_and_time[0]
                                date = y.split("/")
                                self.DeviceActivatedMonth = date[0]
                                self.DeviceActivatedDay = date[1]
                                self.DeviceActivatedYear = date[2][:4]
                                self.DeviceActivatedDate = str(self.DeviceActivatedMonth) +"/"+str(self.DeviceActivatedDay) +"/"+ str(self.DeviceActivatedYear)                                
                            else:
                                self.DeviceActivatedMonth = "0"
                                self.DeviceActivatedDay = "0"
                                self.DeviceActivatedYear = "0"
                                self.DeviceActivatedDate = ""
                            '''
                            ################
                            #print (self.ExcellFileData[self.row][0])
                            #print (self.ExcellFileData[self.row][17])
                            #print (self.ExcellFileData[self.row][19])
                            #print (self.ExcellFileData[self.row][20])
                            #if (str(self.ExcellFileData[self.row][20]) == "0000-00-00 00:00:00"):
                            #    print ("ZEROS")
                            if ((len(str(self.ExcellFileData[self.row][20])) > 0) and (str(self.ExcellFileData[self.row][20]) != "0000-00-00 00:00:00")):# <------------- Disconnected Date
                                self.DeviceDisconnectedMonth = str(self.ExcellFileData[self.row][20].month)
                                self.DeviceDisconnectedDay   = str(self.ExcellFileData[self.row][20].day)
                                self.DeviceDisconnectedYear  = str(self.ExcellFileData[self.row][20].year)
                                self.DeviceDisconnectedDate = str(self.DeviceDisconnectedMonth) +"/"+str(self.DeviceDisconnectedDay) +"/"+ str(self.DeviceDisconnectedYear)
                            else:
                                self.DeviceDisconnectedMonth = "0"
                                self.DeviceDisconnectedDay = "0"
                                self.DeviceDisconnectedYear = "0"
                                self.DeviceDisconnectedDate = ""
                            '''
                            if ((len(self.ExcellFileData[self.row][20]) > 8) and (self.ExcellFileData[self.row][20].find("/") > 0) and (self.ExcellFileData[self.row][20].find(":") > 0)): 
                                date_and_time = self.ExcellFileData[self.row][20].split(":")
                                y = date_and_time[0]
                                date = y.split("/")
                                self.DeviceDisconnectedMonth = date[0]
                                self.DeviceDisconnectedDay = date[1]
                                self.DeviceDisconnectedYear = date[2][:4]
                                self.DeviceDisconnectedDate = str(self.DeviceDisconnectedMonth) +"/"+str(self.DeviceDisconnectedDay) +"/"+ str(self.DeviceDisconnectedYear)
                            else:
                                self.DeviceDisconnectedMonth = "0"
                                self.DeviceDisconnectedDay = "0"
                                self.DeviceDisconnectedYear = "0"
                                self.DeviceDisconnectedDate = ""
                            '''
                            
                            ##################
                            self.DeviceExpirationDate = ""
                            self.DeviceExpirationMonth = "0"
                            self.DeviceExpirationDay = "0"
                            self.DeviceExpirationYear = "0"
                            self.Action = ""
                            PrimaryKey = self.DeviceID
                            sql = """
                                    SELECT * FROM Devices
                                    WHERE Device_ID = '%s'
                                  """ % (PrimaryKey)
                            if (self.db.Execute(sql)): # <------------------------------------------------ Device ID Found! --------------------------------------
                                self.Get_Device_Information(0)
                                self.Change_happened = False
                                today_str = str(today.month) + "/" + str(today.day) + "/" +str(today.year)
                                if ((self.Last_Updated_CMDB_Date == today_str) or (self.Status == 'obsolete')):
                                    self.Action = self.Action + "Device ID, was found and has been updated already, or it is Obsolte, or Possible Duplicate Device Name in the File,"
                                    self.ExcellFile.Add_DataToWorksheetSingleCell("No Changes",self.row+1,22,"Hardware","Fill",lightsalmon_1,11,'Bold')
                                else:
                                    self.Action = self.Action + "Device ID, was found and the Following Changes were made:,"
                                    self.Update_Device_Table('Device_ID',self.DeviceID)
                            else:
                                PrimaryKey = self.DeviceSerialNo
                                sql = """
                                        SELECT * FROM Devices
                                        WHERE Serial_No = '%s'
                                      """ % (PrimaryKey)
                                if (self.db.Execute(sql)): # <-------------------------------------------- Serial No. Found! --------------------------------------
                                    self.Get_Device_Information(0)
                                    self.Add_Device_Table()
                                    self.Action = self.Action + "Serial No. Found but Added with a NEW Device ID,"
                                else:                                
                                    PrimaryKey = self.DeviceIPAddress
                                    sql = """
                                            SELECT * FROM Devices
                                            WHERE IP4_Address = '%s'
                                          """ % (PrimaryKey)
                                    if (self.db.Execute(sql)): # <---------------------------------------- IP Address Found! ----------------------------------------
                                        self.Get_Device_Information(0)
                                        self.Add_Device_Table()
                                        self.Action = self.Action + "IP Address Found but Added with a NEW Device ID,"
                                    else:
                                        sql = """SELECT * FROM Devices 
                                                WHERE %s LIKE '%s' """ % ('IP4_Address',('%'+self.network+'%'))
                                        if (self.db.Execute(sql)): # <------------------------------------ IP Network Found! -------------------------------------------
                                            self.Get_Device_Information(0)
                                            self.Add_Device_Table()
                                            self.Action = self.Action + "IP Network Found and Added a NEW Device ID,"
                                        else:
                                            self.Add_Device_Table()
                                            self.ExcellFile.Add_DataToWorksheetSingleCell("NOT Found But IF Added to UNKNOWN",self.row+1,23,"Hardware","Fill",lightsalmon_1,11,'Bold')
                                            self.Action = self.Action + "NOT Found But Added to UNKNOWN,"
                            #self.Action = "["+self.network +"] " +self.DeviceInstalledDate + " - " + self.DeviceActivatedDate + " - " + self.DeviceDisconnectedDate
                            
                            item = [self.DeviceID,          # Name
                                    self.DeviceIPAddress,   # Real IP
                                    self.DeviceModelID,     # Model
                                    self.DeviceSerialNo,    # Serial No
                                    self.DeviceStatus,      # Status
                                    self.Action]
                            self.results.append(item)
                            num = num + 1
                            self.ImportCMDBTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                        #print (self.ExcellFileData[self.row][column])
                        #column = column + 1
                        self.row = self.row + 1
                        self.progress['value'] = self.row
                        #print ("["+str(self.row)+"]:["+str(column)+"]")
                    self.ExcellFile.Save_File()
                    self.data_ready = True
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    self.ButtonImportCMDBDXCImport['state'] = DISABLED                    
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = 'The Excel File has been imported, please look at the Processed File for more Details')
                    Temporary_File_Name = self.file_name.replace(".","-BVAnalytics.")
                    self.ImportCMDBProcessedFrameEntry['state'] = 'normal'
                    self.ImportCMDBProcessedFrameEntry.delete(0,END)
                    self.ImportCMDBProcessedFrameEntry.insert(0,Temporary_File_Name)
                    self.ImportCMDBProcessedFrameEntry['state'] = 'readonly'                               
                    #self.ButtonImportCMDBDXCValidate['state'] = ACTIVE
                    #self.ButtonImportCMDBDXCImport['state'] = ACTIVE
                #-------------------------------------------------------------------> Cisco <-------------------------------------------------------
                if (self.Menu_Option == "Cisco"):
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    self.ButtonImportCMDBDXCImport['state'] = DISABLED        
                    No_of_Rows = len(self.ExcellFileData)
                    No_of_Columns = len(self.ExcellFileData[0])
                    Column_Names = self.ExcellFileData[0]
                    self.progress['maximum'] = No_of_Rows
                    self.progress['value'] = 0
                    column = 0
                    self.row = 1
                    i = 0
                    num = 0
                    self.results = []
                    self.progress['maximum'] = No_of_Rows
                    self.progress['value'] = 0
                    while (self.row < No_of_Rows):
                        #column = 0
                        #while (column <No_of_Columns):
                        #========================================== Here is all the Work to valildate the Device Database Each Column NEEDS to be EXACT =============
                        tags = "-" # To use in the future                          
                        '''
                        # Columns required on the Cisco File
                        0 'Contract Type',
                        1  'Contract #',
                        2  'Item Status',
                        3  'Item Name',
                        4  'Serial Number',
                        5  'Qty',
                        6  'Item Begin Date',
                        7  'Item End Date',
                        8  'Last Date of Support',
                        9  'Ship Date',
                        10 'Item Type',
                        11 'Product Group',
                        12 'Install Site ID',
                        13 'Install Site Name',
                        14 'Install Site Address',
                        15 'Install Site Address 2',
                        16 'City',
                        17 'State',
                        18 'Zipcode',
                        19 'Country',
                        20 'Maintenance Bill-To ID',
                        21 'Maintenance Bill-To Name',
                        22 'Maintenance SO#',
                        23 'Maintenance PO#',
                        24 'Instance ID',
                        25 'Contract Status',
                        26 'Currency Code')
                        '''                        
                        if (len(str(self.ExcellFileData[self.row][4])) > 0):
                            self.DeviceModelID = str(self.ExcellFileData[self.row][3])
                            self.DeviceSerialNo = str(self.ExcellFileData[self.row][4])
                            if (len(str(self.ExcellFileData[self.row][9])) > 0): # <------------- Ship Date
                                self.DeviceInstalledYear  = str(self.ExcellFileData[self.row][9].year)
                                self.DeviceInstalledMonth = str(self.ExcellFileData[self.row][9].month)
                                self.DeviceInstalledDay   = str(self.ExcellFileData[self.row][9].day)
                                self.DeviceInstalledDate = str(self.DeviceInstalledMonth) +"/"+str(self.DeviceInstalledDay) +"/"+ str(self.DeviceInstalledYear)
                            else:
                                self.DeviceInstalledMonth = "0"
                                self.DeviceInstalledDay = "0"
                                self.DeviceInstalledYear = "0"
                                self.DeviceInstalledDate = ""

                            if (len(str(self.ExcellFileData[self.row][7])) > 0): # <------------- Expriration Date
                                self.DeviceExpirationYear  = str(self.ExcellFileData[self.row][7].year)
                                self.DeviceExpirationMonth = str(self.ExcellFileData[self.row][7].month)
                                self.DeviceExpirationDay   = str(self.ExcellFileData[self.row][7].day)
                                self.DeviceExpirationDate = str(self.DeviceExpirationMonth) +"/"+str(self.DeviceExpirationDay) +"/"+ str(self.DeviceExpirationYear)                                
                            else:
                                self.DeviceExpirationMonth = "0"
                                self.DeviceExpirationDay = "0"
                                self.DeviceExpirationYear = "0"
                                self.DeviceExpirationDate = ""
                            self.DeviceContract = str(self.ExcellFileData[self.row][0]) + "/" + str(self.ExcellFileData[self.row][1])
                            self.Action = ""
                            PrimaryKey = self.DeviceSerialNo                            
                            sql = """
                                    SELECT * FROM Devices
                                    WHERE Serial_No = '%s'
                                  """ % (PrimaryKey)
                            if (self.db.Execute(sql)): # <------------------------------------------------ Serial No. Found! --------------------------------------
                                i = 0
                                while (i < len(self.db.results)):
                                    self.Get_Device_Information(i)
                                    i = i + 1
                                    self.Change_happened = False
                                    if (self.Expiration_Date == self.DeviceExpirationDate):
                                        self.Action = self.Action + "Serial No., was found and has been updated already,"
                                        self.ExcellFile.Add_DataToWorksheetSingleCell("No Changes",self.row+1,28,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                    else:
                                        #--------------------------------------------- Date Installed / Shiped  -------------------------------
                                        sql = "UPDATE Devices SET Executed_by_UserID = '%s', Date_Installed = '%s', Day_Installed = '%d', Month_Installed = '%d', Year_Installed = '%d', \
                                                Expiration_Date = '%s', Day_Expiration = '%d', Month_Expiration = '%d', Year_Expiration = '%d', Contract_No = '%s' \
                                                WHERE Serial_No = '%s'" %(self.Username,self.DeviceInstalledDate,int(self.DeviceInstalledDay),int(self.DeviceInstalledMonth),
                                                                        int(self.DeviceInstalledYear), self.DeviceExpirationDate, int(self.DeviceExpirationDay), int(self.DeviceExpirationMonth), int(self.DeviceExpirationYear),
                                                                        self.DeviceContract, PrimaryKey)
                                        if (self.db.Add_Move_Change_Data(sql)):
                                            self.Action = self.Action + "Serial No., was found and the Following Changes were made:, Ship Date:," + self.DeviceInstalledDate + ", Expiration Date:,"+self.DeviceExpirationDate + ","
                                            self.ExcellFile.Add_DataToWorksheetSingleCell("Updated",self.row+1,28,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                            self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][0],self.row+1,0+1,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                            self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][1],self.row+1,1+1,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                            self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][7],self.row+1,7+1,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                            self.ExcellFile.Add_DataToWorksheetSingleCell(self.ExcellFileData[self.row][9],self.row+1,9+1,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                            self.Change_happened = True
                                        else:
                                            self.Action = self.Action + "Date Installed/Ship and Expriration Changed, was not updated in the Database,"
                                            self.ExcellFile.Add_DataToWorksheetSingleCell("NOT Updated",self.row+1,28,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                    item = [self.Device_ID,         # Name
                                            self.IP4_Address,       # Real IP
                                            self.DeviceModelID,     # Model
                                            self.DeviceSerialNo,    # Serial No
                                            self.Status,            # Status
                                            self.Action]
                                    self.results.append(item)
                                    num = num + 1
                                    self.ImportCMDBTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)                                        
                            else:
                                self.ExcellFile.Add_DataToWorksheetSingleCell("NOT Found",self.row+1,28,"Contract Summary Detail - BO","Fill",lightsalmon_1,11,'Bold')
                                self.Action = self.Action + "NOT Found,"
                        self.row = self.row + 1
                        self.progress['value'] = self.row
                    self.ExcellFile.Save_File()
                    self.data_ready = True
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    self.ButtonImportCMDBDXCImport['state'] = DISABLED                    
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = 'The Excel File has been imported, please look at the Processed File for more Details')
                    Temporary_File_Name = self.file_name.replace(".","-BVAnalytics.")
                    self.ImportCMDBProcessedFrameEntry['state'] = 'normal'
                    self.ImportCMDBProcessedFrameEntry.delete(0,END)
                    self.ImportCMDBProcessedFrameEntry.insert(0,Temporary_File_Name)
                    self.ImportCMDBProcessedFrameEntry['state'] = 'readonly'                               
            else:
                mbox.showerror(master=self.ImportCMDBFrame,title='ImportCMDB',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")


    def Open_File(self,Menu_Option):
        self.data_ready = False
        self.Menu_Option = Menu_Option
        input_file_name = ""
        if (Is_SaveAs_Available):
            input_file_name = tkinter.filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),
                                                                                                       ("All Files", "*.*")])
        else:
            mbox.showerror(master=self.ImportCMDBFrame, title='Opening an Exell File', message = 'The Excell File Libraries are not Installed')        
        if input_file_name != "":
            self.file_name = input_file_name
            self.ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
            if (self.Menu_Option == "DXC"): #<----------------------------------------------------------------------- DXC File Format ---------
                self.ExcellFileData = self.ExcellFile.Read_DataFromWorksheet("Hardware") # <- from the CMDB Tab
                if (len(self.ExcellFileData) > 0):
                    self.ImportCMDBImportFrameEntry['state'] = 'normal'
                    self.ImportCMDBImportFrameEntry.delete(0,END)
                    self.ImportCMDBImportFrameEntry.insert(0,self.file_name)
                    self.ImportCMDBImportFrameEntry['state'] = 'readonly'
                    self.ButtonImportCMDBDXCValidate['state'] = ACTIVE
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                        message = 'The Excel File has been succesfully opened')
                else:
                    self.ImportCMDBImportFrameEntry['state'] = 'normal'
                    self.ImportCMDBImportFrameEntry.delete(0,END)
                    self.ImportCMDBImportFrameEntry.insert(0,"Error Opening File")
                    self.ImportCMDBImportFrameEntry['state'] = 'readonly'
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                        message = 'The Excel File you selected did not have the TAB named "Hardware" on it')
            
            if (self.Menu_Option == "Cisco"): #<----------------------------------------------------------------------- Cisco File Format ---------
                self.ExcellFileData = self.ExcellFile.Read_DataFromWorksheet("Contract Summary Detail - BO") # <- from the Excell Tab
                if (len(self.ExcellFileData) > 0):
                    self.ImportCMDBImportFrameEntry['state'] = 'normal'
                    self.ImportCMDBImportFrameEntry.delete(0,END)
                    self.ImportCMDBImportFrameEntry.insert(0,self.file_name)
                    self.ImportCMDBImportFrameEntry['state'] = 'readonly'
                    self.ButtonImportCMDBDXCValidate['state'] = ACTIVE
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                        message = 'The Excel File has been succesfully opened')
                else:
                    self.ImportCMDBImportFrameEntry['state'] = 'normal'
                    self.ImportCMDBImportFrameEntry.delete(0,END)
                    self.ImportCMDBImportFrameEntry.insert(0,"Error Opening File")
                    self.ImportCMDBImportFrameEntry['state'] = 'readonly'
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                        message = 'The Excel File you selected did not have the TAB named "Contract Summary Detail - BO" on it')



    #-------------- Using a Password Question to make sure it was the intent to be deleted <Begin> ---------------
    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.GetPasswordWindow.destroy()
            if (self.Menu_Option == "DXC"):
                self.Open_File("DXC")
            if (self.Menu_Option == "Cisco"):
                self.Open_File("Cisco")
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


    def Open_File_DXC(self):
        self.Menu_Option = "DXC"
        self.Get_Usernanme_and_Password()
        #self.Open_File("DXC")

    def Open_File_Cisco(self):
        # need to add Pasword request
        self.Menu_Option = "Cisco"
        self.Get_Usernanme_and_Password()
        #self.Open_File("Cisco")


    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "ImportCMDB List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.ImportCMDBTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                #Column = 1
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = self.results[i]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    #Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ImportCMDBFrame,title='Import CMDB',
                            message = '!!! The File was saved !!!')
                    self.ImportCMDBWindow.title("Import CMDB                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['ImportCMDB','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.ImportCMDBFrame,title='ImportCMDB',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_ImportCMDB_Window(self):
        global ImportCMDBLabel
        if not self.ImportCMDBWindowExist:
            self.ImportCMDBWindowExist = True
            # Set up the Window
            self.ImportCMDBWindow = Tk()
            self.ImportCMDBWindow.geometry('930x650+350+70')
            self.ImportCMDBWindow.title("Import Network Inventory")
            self.ImportCMDBWindow.protocol("WM_DELETE_WINDOW", self.on_ImportCMDBWindow_quit)
            self.ImportCMDBWindow.call('tk', 'scaling', self.Windows_Scaling)
            ImportCMDBLabel = Label(self.ImportCMDBWindow,text="Helvetica", font=("Helvetica", 19))
            ImportCMDBLabel["text"] = "Import Network Inventory"
            ImportCMDBLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.ImportCMDBWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Open DXC Network Gear Inventory File", command=self.Open_File_DXC)
                filemenu.add_command(label="Build a Sample DXC Network Gear Inventory File",command=self.Call_Save_As_Sample_DXC)
                filemenu.add_separator()
                filemenu.add_command(label="Open Cisco SmartNET Report File", command=self.Open_File_Cisco)
                filemenu.add_command(label="Build a Sample Cisco SmartNET Report File", command=self.Call_Save_As_Sample_Cisco)
                filemenu.add_separator()
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_ImportCMDBWindow_quit)
            self.ImportCMDBWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.ImportCMDBFrame = Frame(self.ImportCMDBWindow)
            self.ImportCMDBFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.ImportCMDBFrame.rowconfigure(0, weight=1)
            self.ImportCMDBFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    ImportCMDBImportFrameLabel = Label(self.ImportCMDBFrame,text="Helvetica", font=("Helvetica", 10))
                    ImportCMDBImportFrameLabel["text"] = "Imported File Name:"
                    ImportCMDBImportFrameLabel.place(x=10, y=10)

                    self.ImportCMDBImportFrameEntry = Entry(self.ImportCMDBFrame)
                    self.ImportCMDBImportFrameEntry['width']=50
                    self.ImportCMDBImportFrameEntry.place(x=150, y=10)
                    self.ImportCMDBImportFrameEntry['state'] = DISABLED
                    
                    ImportCMDBProcessedFrameLabel = Label(self.ImportCMDBFrame,text="Helvetica", font=("Helvetica", 10))
                    ImportCMDBProcessedFrameLabel["text"] = "Processed File Name:"
                    ImportCMDBProcessedFrameLabel.place(x=10, y=35)

                    self.ImportCMDBProcessedFrameEntry = Entry(self.ImportCMDBFrame)
                    self.ImportCMDBProcessedFrameEntry['width']=50
                    self.ImportCMDBProcessedFrameEntry.place(x=150, y=35)
                    self.ImportCMDBProcessedFrameEntry['state'] = DISABLED                    
                    
                    InstructionsListBox = Listbox(self.ImportCMDBFrame, width=120, height=10, font=("Helvetica", 8))        
                    InstructionsListBox.place(x=110, y=70)
                    InstructionsListBox.insert(END,"          INSTRUCTIONS ON HOW TO IMPORT NETWORK GEAR INVENTORY:")
                    InstructionsListBox.insert(END,"")
                    InstructionsListBox.insert(END,"   1.- Open the Excel File using the File Menu")
                    InstructionsListBox.insert(END,"   2.- Select the Excel File to be imported")
                    InstructionsListBox.insert(END,"   3.- If the Validation is Correct, the Import button will be enabled to import the File")
                    InstructionsListBox.insert(END,"   4.- The Order for Importing/updating the Database is as follows: A) Device Name, B) Serial No., C) IP Address and D) Network Address")
                    InstructionsListBox.insert(END,"   5.- Open the Processed File and all Cells wiht color and Bold are the Cells that were updated")
                    
                    # Setup Buttons
                    self.ButtonImportCMDBDXCValidate = Button(self.ImportCMDBFrame, text = 'Validate \n Inventory File', command = self.Call_Button_ImportCMDB_Validate_File, state=DISABLED)
                    self.ButtonImportCMDBDXCValidate.place(x = 500, y = 7, width=100, height=50)
                    self.ButtonImportCMDBDXCValidate['state'] = DISABLED

                    self.ButtonImportCMDBDXCImport = Button(self.ImportCMDBFrame, text = 'Import \n Inventory File', command = self.Call_Button_ImportCMDB_Import_File, state=DISABLED)
                    self.ButtonImportCMDBDXCImport.place(x = 620, y = 7, width=100, height=50)
                    self.ButtonImportCMDBDXCImport['state'] = DISABLED

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.ImportCMDBFrame, orient="horizontal",length=725, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=110, y=250)


                    #------------------ TREE VIEW For ImportCMDB Database <BEGIN> -----------------------------------
                    
                    # Create Tree and Scrollbars
                    self.ImportCMDBTreeviewDataColumns = ('Device ID','IP Address','Model','Serial No','Status','Action')
                    
                    self.ImportCMDBTreeview = ttk.Treeview(self.ImportCMDBFrame,columns=self.ImportCMDBTreeviewDataColumns, height=14) # <--- Make sure the frame is correct !!
                    self.ImportCMDBTreeviewysb = Scrollbar(self.ImportCMDBFrame,orient=VERTICAL, command=self.ImportCMDBTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.ImportCMDBTreeviewxsb = Scrollbar(self.ImportCMDBFrame,orient=HORIZONTAL, command=self.ImportCMDBTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.ImportCMDBTreeview['yscroll'] = self.ImportCMDBTreeviewysb.set
                    self.ImportCMDBTreeview['xscroll'] = self.ImportCMDBTreeviewxsb.set

                    # setup headings and column 
                    self.ImportCMDBTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.ImportCMDBTreeview.heading('#1', text='Device ID',   anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'Device ID', False)) # E for East and W for West
                    self.ImportCMDBTreeview.heading('#2', text='IP Address', anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'IP Address', False)) 
                    self.ImportCMDBTreeview.heading('#3', text='Model',     anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'Model', False)) 
                    self.ImportCMDBTreeview.heading('#4', text='Serial No',     anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'Serial No', False)) 
                    self.ImportCMDBTreeview.heading('#5', text='Status',     anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'Status', False)) 
                    self.ImportCMDBTreeview.heading('#6', text='Action',     anchor=W,command=lambda: self.treeview_sort_column(self.ImportCMDBTreeview, 'Action', False)) 
                    
                    self.ImportCMDBTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.ImportCMDBTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.ImportCMDBTreeview.column('#2', stretch=1, width=10)
                    self.ImportCMDBTreeview.column('#3', stretch=1, width=10)
                    self.ImportCMDBTreeview.column('#4', stretch=1, width=10)
                    self.ImportCMDBTreeview.column('#5', stretch=1, width=10)
                    self.ImportCMDBTreeview.column('#6', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.ImportCMDBTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.ImportCMDBTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.ImportCMDBTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.ImportCMDBTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    #self.ImportCMDBTreeview.bind('<ButtonRelease-1>', self.on_ImportCMDB_Tree_select_click) # When Select the Tree
                    #self.ImportCMDBTreeview.bind("<Double-1>", self.On_ImportCMDB_Tree_Refresh)

                    

                    '''
                    # SQL Querry to the ImportCMDB Table
                    sql = """
                                SELECT * FROM ImportCMDB
                                ORDER BY ImportCMDB_Name ASC
                              """
                    if (self.db.Execute(sql)):
                        i = 0
                        self.data_ready = True
                        self.results = []
                        while (i < len(self.db.results)):
                            num = i + 1
                            tags = "-" # To use in the futire                          
                            item = [self.db.results[i][0],
                                    self.db.results[i][1],
                                    self.db.results[i][2]]
                            self.results.append(item)
                            self.ImportCMDBTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.ImportCMDBFrame,title='Import CMDB',
                        message = 'No Records found')
                    '''
                    #------------------ TREE VIEW For ImportCMDB Database <END> -----------------------------------
                    #self.db.Disconnect()
                    self.ImportCMDBWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.ImportCMDBFrame,title='Import CMDB',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the ImportCMDB Class....:")
    ImportCMDB = Class_ImportCMDB("BV",Windows_Scaling)
    ImportCMDB.Display_ImportCMDB_Window()


if __name__ == '__main__':
    Main()

