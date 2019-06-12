#==========================================================================
# Program: BVAnalytics.py
# Author:  Jorge E. Rodriguez
# Date Created: Nov-16-2017
# Date Last Modified: Feb-18-2017
# Summary: This application is to access and Analize Netowork information
#==========================================================================

#***************************************************************
# ==================== Libraries Required <BEGIN> =============*
#***************************************************************

#*******************************************************
#============= READING VARIABLES TABLE ================*
#*******************************************************
try:
    from Utils import *
    Utils = Class_Utils()
    Utils.Get_Values()
    #------- DNS NAME ---------
    ODBC_DSN_name = Utils.Get_ODBC_Name()
    DatabaseVersion = Utils.Get_DB_Ver()
    BVAnalyticsDatabaseVersion = "1.11"
    Windows_Scaling = Utils.Get_Windows_Scaling()
    BackendServer = Utils.Get_BackendServer()
    Enviroment = Utils.Get_Enviroment()    
    #--------------------------
except:
    #------- DNS NAME ---------
    ODBC_DSN_name = "BV"
    DatabaseVersion = "1.1"
    BVAnalyticsDatabaseVersion = "1.11"
    Windows_Scaling = 1.0
    BackendServer = "NOT Available"
    Enviroment = "NOT Available"
    #--------------------------

#-------------------------- Validating the Variables Values <BEGIN> ----------------
#print (ODBC_DSN_name)
#print (DatabaseVersion)
#print (BVAnalyticsDatabaseVersion)
#print (Windows_Scaling)
#-------------------------- Validating the Variables Values <END> ----------------


try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
    Is_Excell_Available = True        
except:
    print ("********************************************************************************* \n")
    print ("*** No openpyxl library exist, please make sure you downlaod it and instal it *** \n")
    print ("********************************************************************************* \n")
    Is_Excell_Available = False    

try:
    from odbc_connector import *
    Is_ODBC_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO ODBC Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_ODBC_Available = False  

try:
    from SaveAs import *
    Is_SaveAs_Available = True
except:
    print ("********************************************************************************** \n")
    print ("*** NO SaveAs Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False

try:
    from Country import *
    Is_Country_Available = True
    Country = Class_Country(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************* \n")
    print ("*** NO Country Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************* \n")
    Is_Country_Available = False  

try:
    from Region import *
    Is_Region_Available = True
    Region = Class_Region(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************ \n")
    print ("*** NO Region Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************ \n")
    Is_Region_Available = False  

try:
    from Facility import *
    Is_Facility_Available = True
    Location = []
    Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("************************************************************************************** \n")
    print ("*** NO Facility Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************** \n")
    Is_Facility_Available = False  

try:
    from Sites import *
    Is_Sites_Available = True
    Sites = Class_Sites(ODBC_DSN_name,Windows_Scaling)
except:
    print ("*********************************************************************************** \n")
    print ("*** NO Sites Library Found, please download it in order to access the Databases *** \n")
    print ("*********************************************************************************** \n")
    Is_Sites_Available = False  

try:
    from Device import *
    Is_Device_Available = True
    Location = []
    Device = Class_Device(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("************************************************************************************* \n")
    print ("*** NO Devices Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************* \n")
    Is_Device_Available = False  

try:
    from Circuits import *
    Is_Circuit_Available = True
    Location = []
    Circuit = Class_Circuits(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("************************************************************************************** \n")
    print ("*** NO Circuits Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************** \n")
    Is_Circuit_Available = False  

try:
    from FindCircuitsEquipmentFacilities import *
    Is_Find_Available = True
    FindCEF = Class_FindCircuitsEquipmentFacilities(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************************************** \n")
    print ("*** NO  FindCircuitsEquipmentFacilities Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************************************** \n")
    Is_Find_Available = False  

try:
    from ICMP import *
    Is_ICMP_Available = True
    Location = []
    ICMP = Class_ICMP(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("************************************************************************************** \n")
    print ("*** NO Circuits Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************** \n")
    Is_ICMP_Available = False  

try:
    from LocalPointOfContacts import *
    Is_LocalPointOfContacts_Available = True
    Location = []
    LocalPointOfContacts = Class_LocalPointOfContacts(ODBC_DSN_name,Windows_Scaling,Location)
except:
    print ("***************************************************************************************************** \n")
    print ("*** NO Local Point Of Contacts Library Found, please download it in order to access the Databases *** \n")
    print ("***************************************************************************************************** \n")
    Is_LocalPointOfContacts_Available = False  


try:
    from FindLocalPointOfContacts import *
    Is_FindLocalPointOfContacts_Available = True
    FindLocalPointOfContacts = Class_FindLocalPointOfContacts(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************************************** \n")
    print ("*** NO Find Local Point Of Contacts Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************************************** \n")
    Is_FindLocalPointOfContacts_Available = False  

try:
    from FindNetworks import *
    Is_FindNetworks_Available = True
    FindNetworks = Class_FindNetworks(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************************************** \n")
    print ("*** NO Find Networks Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************************************** \n")
    Is_FindNetworks_Available = False  


try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['BVAnalytics Main Screen','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("************************************************************************************* \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************* \n")
    Is_logging_Available = False

try:
    from FindLogging import *
    Is_FindLogging_Available = True
    FindLogging = Class_FindLogging(ODBC_DSN_name,Windows_Scaling)
except:
    print ("****************************************************************************************** \n")
    print ("*** NO Find Logging Library Found, please download it in order to access the Databases *** \n")
    print ("****************************************************************************************** \n")
    Is_FindLogging_Available = False  

try:
    from DeviceType import *
    Is_DeviceType_Available = True
    DeviceType = Class_DeviceType(ODBC_DSN_name,Windows_Scaling)
except:
    print ("**************************************************************************************** \n")
    print ("*** NO DeviceType Library Found, please download it in order to access the Databases *** \n")
    print ("**************************************************************************************** \n")
    Is_DeviceType_Available = False  

try:
    from DeviceModel import *
    Is_DeviceModel_Available = True
    DeviceModel = Class_DeviceModel(ODBC_DSN_name,Windows_Scaling)
except:
    print ("***************************************************************************************** \n")
    print ("*** NO DeviceModel Library Found, please download it in order to access the Databases *** \n")
    print ("***************************************************************************************** \n")
    Is_DeviceModel_Available = False  

try:
    from ReportOutsourceCostPerCountry import *
    Is_ReportOutsourcePerCountry_Available = True
    ReportOutsourcePerCountry = Class_ReportOutsourcePerCountry(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************************** \n")
    print ("*** NO ReportOutsourcePerCountry Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************************** \n")
    Is_ReportOutsourcePerCountry = False  

try:
    from ReportOutsourceCostPerFacility import *
    Is_ReportOutsourcePerFacility_Available = True
    ReportOutsourcePerFacility = Class_ReportOutsourcePerFacility(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************************** \n")
    print ("*** NO ReportOutsourcePerFacility Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************************** \n")
    Is_ReportOutsourcePerFacility = False

try:
    from Vendor import *
    Is_Vendor_Available = True
    Vendor = Class_Vendor(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************ \n")
    print ("*** NO Vendor Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************ \n")
    Is_Vendor_Available = False  
    
try:
    from Carriers import *
    Is_Carrier_Available = True
    Carrier = Class_Carrier(ODBC_DSN_name,Windows_Scaling)
except:
    print ("************************************************************************************* \n")
    print ("*** NO Carrier Library Found, please download it in order to access the Databases *** \n")
    print ("************************************************************************************* \n")
    Is_Carrier_Available = False  

try:
    from CircuitTypes import *
    Is_CircuitType_Available = True
    CircuitType = Class_CircuitType(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Circuit Types Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_CircuitType_Available = False  

try:
    from PortSpeeds import *
    Is_PortSpeed_Available = True
    PortSpeed = Class_PortSpeed(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Port Speed Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_PortSpeed_Available = False  

try:
    from ImportCMDB import *
    Is_ImportCMDB_Available = True
    ImportCMDB = Class_ImportCMDB(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Import CMDB Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_ImportCMDB_Available = False  

try:
    from NetworkTemplates import *
    Is_NetworkTemplates_Available = True
    NetworkTemplates = Class_NetworkTemplates(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Network Templates Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_NetworkTemplates_Available = False  
    
try:
    from NetworkTemplatesBreakOut import *
    Is_NetworkTemplatesBreakOut_Available = True
    NetworkTemplatesBreakOut = Class_NetworkTemplatesBreakOut(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Network TemplatesBreakout Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_NetworkTemplatesBreakOut_Available = False  
 
try:
    from NetworkAllocations import *
    Is_NetworkAllocations_Available = True
    NetworkAllocations = Class_NetworkAllocations(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Network Allocations Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_NetworkAllocations_Available = False  

try:
    from NetworkProvision import *
    Is_NetworkProvisioning_Available = True
    NetworkProvisioning = Class_NetworkProvisioning(ODBC_DSN_name,Windows_Scaling)
except:
    print ("******************************************************************************************* \n")
    print ("*** NO Network Allocations Library Found, please download it in order to access the Databases *** \n")
    print ("******************************************************************************************* \n")
    Is_NetworkProvisioning_Available = False  


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

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************


#******************************************************
# ================= Global Variables <BEGIN>==========*
#******************************************************

BVAnalyticsVer = 'Version 6.0'
file_name = ""
file_saved_status = ""
#--------------- Windows Control State -----------
UpdateTablesWindowExist      = False
ICMPWindowExist              = False
ICMPSummaryWindowExist       = False
TracerouteSummaryWindowExist = False
PercentageLossWindowExist    = False
CountryWindowExist           = False

Username = os.getlogin()
#-------------------------------------------------
DateAndTime = datetime.datetime.now()
if Is_ODBC_Available:
    db = ODBC(ODBC_DSN_name)
    db2 = ODBC(ODBC_DSN_name)
    version = db.Get_Version()
else:
    db = "NONE"
    db2 = "NONE"
    version = "NONE"
    
#****************************************************
# ================= Global Variables <END>==========*
#****************************************************


#================ Menu Subrutines <BEGIN>=============

def Window_About():
    STRMessageAbout = 'Black & Veatch Analytics Tool for: \n'
    STRMessageAbout = STRMessageAbout + 'Configuation and Network Management \n'
    STRMessageAbout = STRMessageAbout + BVAnalyticsVer + '\n'                       # Sequence No. of Modifications / improvements
    STRMessageAbout = STRMessageAbout + 'August 29, 2018'                            # Date of Modifications 
    STRMessageAbout = STRMessageAbout + '\n\n\n'
    STRMessageAbout = STRMessageAbout + 'Last Compiled By Jorge Rodriguez'          # Who made last changes
    mbox.showinfo('About ',STRMessageAbout) 


def Database_Version():
    STRMessageAbout = 'Black & Veatch Analytics Tool for: \n'
    STRMessageAbout = STRMessageAbout + 'Configuation and Network Management \n'
    STRMessageAbout = STRMessageAbout + 'Database Version: [' + DatabaseVersion + ']\n'                       # Sequence No. of Modifications / improvements
    STRMessageAbout = STRMessageAbout + 'April 39th, 2018'                            # Date of Modifications 
    STRMessageAbout = STRMessageAbout + '\n\n\n'
    STRMessageAbout = STRMessageAbout + 'Last Modified By Jorge Rodriguez'          # Who made last changes
    mbox.showinfo('Database Version ',STRMessageAbout) 



def treeview_sort_column(tv, col, reverse):
    #print('sorting %s!' % col)
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)

    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        #print('Moving Index:%r, Value:%r, k:%r' % (index, val, k))
        tv.move(k, '', index)

    # reverse sort next time
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


def donothing():
    mbox.showerror(master=root,title='Under Construction',
                         message = '... Comming Soon...')               

def CloseMainWindow():
    global root

    if UpdateTablesWindowExist:
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Update Database Tables','CLOSE Window']    
            Logging.Log(Parameter)
        TablesWindow.destroy()

    if ICMPWindowExist:
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics ICMP Window','CLOSE Window']    
            Logging.Log(Parameter)
        ICMPWindow.destroy()
        
    if ICMPSummaryWindowExist:
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics ICMP Summary Window','CLOSE Window']    
            Logging.Log(Parameter)
        ICMPSummaryWindow.destroy()

    if TracerouteSummaryWindowExist:
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Traceroute Summary Window','CLOSE Window']    
            Logging.Log(Parameter)
        TracerouteSummaryWindow.destroy()

    if PercentageLossWindowExist:
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Percentage Loss Windows','CLOSE Window']    
            Logging.Log(Parameter)
        PercentageLossWindow.destroy()

        
    if (Is_Country_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Country Window','CLOSE Window']    
            Logging.Log(Parameter)
        Country.on_CountryWindow_quit()

    if (Is_Region_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Region Window','CLOSE Window']    
            Logging.Log(Parameter)
        Region.on_RegionWindow_quit()

    if (Is_Facility_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Facility Window','CLOSE Window']    
            Logging.Log(Parameter)        
        Facility.on_FacilityWindow_quit()

    if (Is_Sites_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Sites Window','CLOSE Window']    
            Logging.Log(Parameter)
        Sites.on_SitesWindow_quit()

    if (Is_Device_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Devices Window','CLOSE Window']    
            Logging.Log(Parameter)
        Device.on_DeviceWindow_quit()

    if (Is_DeviceType_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Devices Type Window','CLOSE Window']    
            Logging.Log(Parameter)
        DeviceType.on_DeviceTypeWindow_quit()

    if (Is_DeviceModel_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Devices Model Window','CLOSE Window']    
            Logging.Log(Parameter)
        DeviceModel.on_DeviceModelWindow_quit()

    if (Is_Circuit_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Circuits Window','CLOSE Window']    
            Logging.Log(Parameter)
        Circuit.on_CircuitsWindow_quit()

    if (Is_Find_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics FindCEF Window','CLOSE Window']    
            Logging.Log(Parameter)
        FindCEF.on_FindCircuitWindow_quit()
        FindCEF.on_FindDeviceWindow_quit()
        FindCEF.on_FindFacilityWindow_quit()

    if (Is_LocalPointOfContacts_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Local Point Of Contacts Window','CLOSE Window']    
            Logging.Log(Parameter)
        LocalPointOfContacts.on_LocalPointOfContactsWindow_quit()
        
    if (Is_ICMP_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics ICMP Window','CLOSE Window']    
            Logging.Log(Parameter)
        ICMP.on_ICMPWindow_quit()

    if (Is_FindLocalPointOfContacts_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Find Local POC Window','CLOSE Window']    
            Logging.Log(Parameter)
        FindLocalPointOfContacts.on_FindLocalPointOfContactsWindow_quit()

    if (Is_FindNetworks_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Find Networks Window','CLOSE Window']    
            Logging.Log(Parameter)
        FindNetworks.on_FindNetworksWindow_quit()

    if (Is_FindLogging_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Find Logging Window','CLOSE Window']    
            Logging.Log(Parameter)
        FindLogging.on_FindLoggingWindow_quit()

    if (Is_ReportOutsourcePerCountry_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Report Outsource Cost per Coutry Window','CLOSE Window']    
            Logging.Log(Parameter)
        ReportOutsourcePerCountry.on_ReportOutsourcePerCountryWindow_quit()

    if (Is_ReportOutsourcePerFacility_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Report Outsource Cost per Facility Window','CLOSE Window']    
            Logging.Log(Parameter)
        ReportOutsourcePerFacility.on_ReportOutsourcePerFacilityWindow_quit()

    if (Is_Vendor_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Vendor Window','CLOSE Window']    
            Logging.Log(Parameter)
        Vendor.on_VendorWindow_quit()

    if (Is_Carrier_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Carrier Window','CLOSE Window']    
            Logging.Log(Parameter)
        Carrier.on_CarrierWindow_quit()

    if (Is_CircuitType_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics CircuitType Window','CLOSE Window']    
            Logging.Log(Parameter)
        CircuitType.on_CircuitTypeWindow_quit()

    if (Is_PortSpeed_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics PortSpeed Window','CLOSE Window']    
            Logging.Log(Parameter)
        PortSpeed.on_PortSpeedWindow_quit()

    if (Is_ImportCMDB_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics ImportCMDB Window','CLOSE Window']    
            Logging.Log(Parameter)
        ImportCMDB.on_ImportCMDBWindow_quit()

    if (Is_NetworkTemplates_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics NetworkTemplates Window','CLOSE Window']    
            Logging.Log(Parameter)
        NetworkTemplates.on_NetworkTemplatesWindow_quit()

    if (Is_NetworkTemplatesBreakOut_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics NetworkTemplatesBreakOut Window','CLOSE Window']    
            Logging.Log(Parameter)
        NetworkTemplatesBreakOut.on_NetworkTemplatesBreakOutWindow_quit()

    if (Is_NetworkAllocations_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Is_NetworkAllocations_Available Window','CLOSE Window']    
            Logging.Log(Parameter)
        NetworkAllocations.on_NetworkAllocationsWindow_quit()

    if (Is_NetworkProvisioning_Available):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['BVAnalytics Is_NetworkProvisioning_Available Window','CLOSE Window']    
            Logging.Log(Parameter)
        NetworkProvisioning.on_NetworkProvisioningWindow_quit()

    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Main Screen','CLOSE Window']    
        Logging.Log(Parameter)

    root.destroy()

####################################################################################
#----------------------------- Database Menu <BEGIN> ------------------------------#
####################################################################################    
    
def Test_ODBC_Connection():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics ODBC Test Connection','OPEN Window']    
        Logging.Log(Parameter)
    
    db = ODBC(ODBC_DSN_name)
    timezero = time.time()
    if db.Connect():
        db.Disconnect()
        timeone = time.time()
        totaltime = timeone - timezero
        mbox.showinfo(master=root,title='TEST ODBC Connection',
                         message = 'THE ODBC Connection was Succesful \r The Connection Latency was ['+str(totaltime)+'] seconds')
    else:
        mbox.showerror(master=root,title='TEST ODBC Connection',
                         message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                   + 'Please make sure the ODBC DSN Name mathes: ['
                                   + ODBC_DSN_name + "]")

def on_TablesWindow_quit():
    global UpdateTablesWindowExist
    global TablesWindow
    
    UpdateTablesWindowExist = False
    db.Disconnect()
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Database Update Tables Window','CLOSE Window']    
        Logging.Log(Parameter)
    TablesWindow.destroy()


def Update_Tables_and_Fields():
    global UpdateTablesWindowExist
    global TablesWindow
    global TablesLabel
    global TablesListBox

    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Database Update Tables Window','OPEN Window']    
        Logging.Log(Parameter)

    if db.Connect():
        if not UpdateTablesWindowExist:
            UpdateTablesWindowExist = True
            TablesWindow = Tk()
            TablesWindow.geometry('1000x300+350+70')
            TablesWindow.title("Updating TABLES and FIELDS into the Database")
            TablesWindow.protocol("WM_DELETE_WINDOW", on_TablesWindow_quit)
            TablesWindow.call('tk', 'scaling', Windows_Scaling)
            TablesLabel = Label(TablesWindow,text="Helvetica", font=("Helvetica", 15))
            TablesLabel["text"] = "Updating TABLES and FIELDS into the Database"
            TablesLabel.pack()

            TablesFrame = Frame(TablesWindow)
            TablesFrame.pack()
            TablesListBox = Listbox(TablesFrame, width=120, height=30, font=("Helvetica", 9))        
            yscrollbar = Scrollbar(TablesFrame, orient="vertical")
            yscrollbar.config(command=TablesListBox.yview)
            yscrollbar.pack(side="right", fill="y")
            TablesListBox.config(yscrollcommand=yscrollbar.set)        
            xscrollbar = Scrollbar(TablesFrame, orient="horizontal")
            xscrollbar.config(command=TablesListBox.xview)
            xscrollbar.pack(side="bottom", fill="y")
            TablesListBox.config(xscrollcommand=xscrollbar.set)
            TablesListBox.pack(side="left", fill="y")
            #------------- GET All Table Names <BEGIN>------------------
            TablesListBox.insert(END,"Opening Database.....")
            tables_names = db.Get_Table_Names()
            #if (len(tables_names) > 0):
            #    print (len(tables_names))
            #    print (tables_names[0])
            #print (tables_names)
            if ((len(tables_names) == 0) or ((len(tables_names) == 2) and
                                             (tables_names[0] == 'trace_xe_action_map') and
                                             (tables_names[1] == 'trace_xe_event_map'))):
                TablesListBox.insert(END,"Empty Database...")
                #------------- CREATE TABLES VARIABLES <BEGIN>-----------------------
                sql = """CREATE TABLE VARIABLES (
                    Database_ver      CHAR(20) NOT NULL,
                    ODBC_DSN_Name     CHAR(20),
                    Created_Date      CHAR(40),  
                    Last_Time_Updated CHAR(40),
                    Updated_by_UserID CHAR(20),
                    Window_Scaling    FLOAT)"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table Variables was Succesful Created.")
                    #------------- INSERT data into a Table <BEGIN> ---------
                    date = datetime.datetime.now()
                    sql = """INSERT INTO VARIABLES(Database_ver,ODBC_DSN_Name, Created_Date, Last_Time_Updated, \
                             Updated_by_UserID,Window_Scaling)
                                VALUES ('1.6', 'BV','%s', '%s', '%s',%f)""" % (date,date,Username,1.0)
                    if (db.Add_Move_Change_Data(sql)):
                        TablesListBox.insert(END,"Record Added to the Variables Table.")
                    else:
                        TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                    #------------- INSERT data into a Table <END> ---------
                else:
                    TablesListBox.insert(END,"Table Variables already Exists")
                #------------- CREATE TABLES VARIABLES <END>-----------------------

                #------------- CREATE TABLE ICMP <BEGIN>-----------------------
                sql = """CREATE TABLE ICMP (
                        Device_IP_Date_Time_Size_of_Ping CHAR(100) NOT NULL PRIMARY KEY,
                        Device_IP                        CHAR(20) NOT NULL,
                        Date_String                      CHAR(20) NOT NULL,
                        Time_String                      CHAR(20) NOT NULL,
                        Day                              INT,
                        Month                            INT,
                        Year                             INT,
                        Hour                             INT,
                        Minute                           INT,
                        Second                           INT,
                        Size_of_Ping                     CHAR(10) NOT NULL,
                        Percentage_Loss                  INT,
                        Response_Time_Max                INT,
                        Response_Time_Min                INT,
                        Response_Time_Avg                INT,
                        Response_Status                  CHAR(10),
                        Executed_by_UserID               CHAR(20) )"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table ICMP was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table ICMP already Exists")
                #------------- CREATE TABLES ICMP <END>-----------------------
                    
                #------------- CREATE TABLE ICMP_SUMMARY <BEGIN>-----------------------
                sql = """CREATE TABLE ICMP_SUMMARY (
                    Date_Time           CHAR(100) NOT NULL PRIMARY KEY,
                    Date_String         CHAR(20) NOT NULL,
                    Time_String         CHAR(20) NOT NULL,
                    Day                 INT,
                    Month               INT,
                    Year                INT,
                    Hour                INT,
                    Minute              INT,
                    Second              INT,
                    Tota_no_of_Pings    INT,
                    Normal_Pings        INT,
                    Time_Out_Pings      INT,
                    TTL_Expired_Pings   INT,
                    None_Pings          INT,
                    Percentage_Loss_0   INT,  
                    Percentage_Loss_25  INT,  
                    Percentage_Loss_50  INT,  
                    Percentage_Loss_75  INT,  
                    Percentage_Loss_100 INT,  
                    Executed_by_UserID   CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table ICMP SUMMARY was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table ICMP SUMMARY already Exists")
                #------------- CREATE TABLES ICMP_SUMMARY <END>-----------------------

                #------------- CREATE TABLE TRACEROUTE <BEGIN>-----------------------
                sql = """CREATE TABLE TRACEROUTE (
                    Device_IP_Date_Time_Hop_No CHAR(100) NOT NULL PRIMARY KEY,
                    Device_IP                  CHAR(20) NOT NULL,
                    Date_String                CHAR(20) NOT NULL,
                    Time_String                CHAR(20) NOT NULL,
                    Hop_No                     INT,
                    Day                        INT,
                    Month                      INT,
                    Year                       INT,
                    Hour                       INT,
                    Minute                     INT,
                    Second                     INT,
                    Next_Hop                   CHAR(20),
                    Response_Time_Max          CHAR(10),
                    Response_Time_Min          CHAR(10),
                    Response_Time_Avg          CHAR(10),
                    Executed_by_UserID         CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table TRACEROUTE was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table TRACEROUTE already Exists")
                #------------- CREATE TABLES TRACEROUTE <END>-----------------------

                #------------- CREATE TABLE TRACEROUTE SUMMARY <BEGIN>-----------------------
                sql = """CREATE TABLE TRACEROUTE_SUMMARY (
                    Date_Time                              CHAR(100) NOT NULL PRIMARY KEY,
                    Date_String                            CHAR(20) NOT NULL,
                    Time_String                            CHAR(20) NOT NULL,
                    Total_No_Of_Traceroute                 INT,
                    Day                                    INT,
                    Month                                  INT,
                    Year                                   INT,
                    Hour                                   INT,
                    Minute                                 INT,
                    Second                                 INT,
                    Total_Max_Response_Time                INT,
                    Total_Min_Response_Time                INT,
                    Total_Avg_Response_Time_lower_limit    INT,
                    Total_Avg_Response_Time_Upper_limit    INT,
                    Executed_by_UserID                     CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table TRACEROUTE SUMMARY was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table TRACEROUTE SUMMARY already Exists")
                #------------- CREATE TABLES TRACEROUTE SUMMARY<END>-----------------------

                #------------- CREATE TABLE CIRCUITS <BEGIN>-----------------------
                sql = """CREATE TABLE CIRCUITS (
                    Circuit_ID              CHAR(100) NOT NULL PRIMARY KEY,
                    Description             CHAR(200),
                    Country_ID              CHAR(20),                    
                    Region_ID               CHAR(20),
                    Facility_ID             CHAR(20),
                    Site_ID                 CHAR(20),
                    Carrier_ID              CHAR(20),
                    Circuit_Type            CHAR(40),
                    Port_Speed              CHAR(20),
                    Bandwidth               CHAR(20),
                    Status                  CHAR(20),
                    Dmarc_Info_1            CHAR(200),
                    Dmarc_Info_2            CHAR(200),
                    LEC1                    CHAR(50),
                    LEC2                    CHAR(50),
                    LEC3                    CHAR(50),
                    LEC4                    CHAR(50),
                    LEC5                    CHAR(50),
                    CE_ASN                  CHAR(20),
                    CE_IP_Address           CHAR(50),
                    PE_ASN                  CHAR(20),
                    PE_IP_Address           CHAR(50),
                    VLAN_ID                 CHAR(10),
                    PE_Switch               CHAR(100),
                    PE_Location             CHAR(100),
                    NPA_NXX                 CHAR(20),
                    Monthly_Cost            FLOAT,
                    Order_Number            CHAR(40),
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
                    Term_Day                CHAR(10),
                    Term_Time               CHAR(10),
                    ETF                     FLOAT,
                    Contract_No             CHAR(50),
                    Account_No              CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table CIRCUITS was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table CIRCUIT already Exists")
                #------------- CREATE TABLES CIRCUITS <END>-----------------------
                    
                #------------- CREATE TABLE COUNTRY <BEGIN>-----------------------
                sql = """CREATE TABLE COUNTRY (
                    Country_ID              CHAR(20) NOT NULL PRIMARY KEY,
                    Country_Name            CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table COUNTRY was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table COUNTRY already Exists")
                #------------- CREATE TABLES COUNTRY <END>-----------------------

                #------------- CREATE TABLE REGION <BEGIN>-----------------------
                sql = """CREATE TABLE REGION (
                    Country_ID_Region_ID    CHAR(45) NOT NULL PRIMARY KEY,
                    Country_ID              CHAR(20) NOT NULL,
                    Region_ID               CHAR(20) NOT NULL,
                    Region_Name             CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table REGION was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table REGION already Exists")
                #------------- CREATE TABLES REGION <END>-----------------------

                #------------- CREATE TABLE FACILITY <BEGIN>-----------------------
                sql = """CREATE TABLE FACILITY (
                    Country_ID_Region_ID_Facility_ID    CHAR(65) NOT NULL PRIMARY KEY,
                    Country_ID                          CHAR(20) NOT NULL,
                    Region_ID                           CHAR(20) NOT NULL,
                    Facility_ID                         CHAR(20) NOT NULL,
                    Facility_Name                       CHAR(50) NOT NULL,
                    Business_Unit_ID                    CHAR(50) NOT NULL,
                    Office_Type                         CHAR(20) NOT NULL,
                    Office_Size                         CHAR(20) NOT NULL,
                    Lease_End_Date                      CHAR(20),
                    Lease_End_Month                     INT,
                    Lease_End_Day                       INT,
                    Lease_End_Year                      INT,
                    Address1                            CHAR(100),
                    Address2                            CHAR(100),
                    City                                CHAR(50),
                    State                               CHAR(50),
                    Zip_Code                            CHAR(20),
                    Country                             CHAR(50),
                    Headcount                           INT,
                    HR_ID_Location                      CHAR(20),
                    HR_Office_ID                        CHAR(20),
                    Status                              CHAR(20),
                    Executed_by_UserID                  CHAR(20),
                    Strategy                            CHAR(250),
                    Action                              CHAR(250),
                    Notes                               CHAR(250),
                    New_Address1                        CHAR(100),
                    New_Address2                        CHAR(100),
                    New_City                            CHAR(50),
                    New_State                           CHAR(50),
                    New_Zip_Code                        CHAR(20),
                    New_Country                         CHAR(50))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table FACILITY was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table FACILITY already Exists")
                #------------- CREATE TABLES FACILITY <END>-----------------------

                #------------- CREATE TABLE SITES <BEGIN>-----------------------
                sql = """CREATE TABLE SITES (
                    Country_ID_Region_ID_Facility_ID_Site_ID    CHAR(85) NOT NULL PRIMARY KEY,
                    Country_ID                                  CHAR(20) NOT NULL,
                    Region_ID                                   CHAR(20) NOT NULL,
                    Facility_ID                                 CHAR(20) NOT NULL,
                    Site_ID                                     CHAR(20) NOT NULL,
                    Site_Name                                   CHAR(50) NOT NULL,
                    Description                                 CHAR(100),
                    Executed_by_UserID                          CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table SITES was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table SITES already Exists")
                #------------- CREATE TABLES SITES <END>-----------------------

                #------------- CREATE TABLE DEVICE_TYPE <BEGIN>-----------------------
                sql = """CREATE TABLE DEVICE_TYPE (
                    Device_Type_ID          CHAR(30) NOT NULL PRIMARY KEY,
                    Device_Type_Name        CHAR(50),
                    Vendor_ID               CHAR(20),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table DEVICE_TYPE was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table DEVICE_TYPE already Exists")
                #------------- CREATE TABLES DEVICE_TYPE <END>-----------------------

                #------------- CREATE TABLE DEVICE_MODEL <BEGIN>----------------------- Added on Ver 1.2
                sql = """CREATE TABLE DEVICE_MODEL (
                    Device_Model_ID         CHAR(30) NOT NULL PRIMARY KEY,
                    Device_Model_Name       CHAR(50),
                    Vendor_ID               CHAR(20),
                    Outsource_Cost          FLOAT,
                    Executed_by_UserID      CHAR(20),
                    Device_Type_ID          CHAR(30))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table DEVICE_MODEL was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table DEVICE_MODEL already Exists")
                #------------- CREATE TABLES DEVICE_MODEL <END>-----------------------

                #------------- CREATE TABLE DEVICE_MODEL_COUNTRY <BEGIN>-----------------------
                sql = """CREATE TABLE DEVICE_MODEL_COUNTRY (
                    Device_Model_ID_Country_ID      CHAR(60) NOT NULL PRIMARY KEY,
                    Device_Model_ID                 CHAR(30) NOT NULL,
                    Country_ID                      CHAR(20) NOT NULL,
                    Outsource_Cost                  FLOAT,
                    Executed_by_UserID              CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table DEVICE_MODEL_COUNTRY was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table DEVICE_MODEL_COUNTRY already Exists")
                #------------- CREATE TABLES DEVICE_MODEL_COUNTRY <END>-----------------------

                #------------- CREATE TABLE VENDOR <BEGIN>-----------------------
                sql = """CREATE TABLE VENDOR (
                    Vendor_ID                  CHAR(20) NOT NULL PRIMARY KEY,
                    Vendor_Name                CHAR(50),
                    Tel_Support                CHAR(50),
                    Contract_No                CHAR(50),
                    Executed_by_UserID         CHAR(20),
                    Point_of_Contact           CHAR(50),
                    Tel_Point_of_Contact       CHAR(50),
                    Email_Point_of_Contact     CHAR(50))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table VENDOR was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table VENDOR already Exists")
                #------------- CREATE TABLES DEVICE_TYPE <END>-----------------------

                #------------- CREATE TABLE DEVICES <BEGIN>-----------------------
                sql = """CREATE TABLE DEVICES (
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
                    MAC_Address             CHAR(20),
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
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table DEVICES was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table DEVICES already Exists")
                #------------- CREATE TABLES DEVICES <END>-----------------------

                #------------- CREATE TABLE CARRIER <BEGIN>-----------------------
                sql = """CREATE TABLE CARRIER (
                    Carrier_ID              CHAR(20) NOT NULL PRIMARY KEY,
                    Description             CHAR(100),
                    Point_Of_Contact        CHAR(50),                    
                    Telephone_NO            CHAR(50),
                    E_Mail_Address          CHAR(100),
                    Web_Site                CHAR(100),
                    Tech_Support_No         CHAR(50),
                    Master_Contract_No      CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table CARRIER was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table CARRIER already Exists")
                #------------- CREATE TABLES CARRIER <END>-----------------------
                    
                #------------- CREATE TABLE CIRCUIT_TYPE <BEGIN>-----------------------
                sql = """CREATE TABLE CIRCUIT_TYPE (
                    Circuit_Type            CHAR(40) NOT NULL PRIMARY KEY,
                    Description             CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table CIRCUIT_TYPE was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table CIRCUIT_TYPE already Exists")
                #------------- CREATE TABLES CIRCUIT_TYPE <END>-----------------------
                    
                #------------- CREATE TABLE PORT_SPEED <BEGIN>-----------------------
                sql = """CREATE TABLE PORT_SPEED (
                    Port_Speed_ID           CHAR(20) NOT NULL PRIMARY KEY,
                    Description             CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table PORT_SPEED was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table PORT_SPEED already Exists")
                #------------- CREATE TABLES PORT_SPEED <END>-----------------------
                    
                #------------- CREATE TABLE LOCAL_POINT_OF_CONTACTS <BEGIN>-----------------------
                sql = """CREATE TABLE LOCAL_POINT_OF_CONTACTS (
                    Country_ID_Region_ID_Facility_ID_SEQ_No     CHAR(85) NOT NULL PRIMARY KEY,
                    Country_ID                                  CHAR(20) NOT NULL,
                    Region_ID                                   CHAR(20) NOT NULL,
                    Facility_ID                                 CHAR(20) NOT NULL,
                    SEQ_No                                      INT NOT NULL,
                    Primary_POC_Name                            CHAR(100),
                    Primary_POC_Office_No                       CHAR(100),
                    Primary_POC_Alternate_No                    CHAR(100),
                    Primary_POC_2nd_Alternate_No                CHAR(100),
                    Secondary_POC_Name                          CHAR(100),
                    Secondary_POC_Office_No                     CHAR(100),
                    Secondary_POC_Alternate_No                  CHAR(100),
                    Secondary_POC_2nd_Alternate_No              CHAR(100),
                    Last_time_it_was_Updated                    CHAR(100),
                    Executed_by_UserID                          CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table LOCAL_POINT_OF_CONTACTS was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table LOCAL_POINT_OF_CONTACTS already Exists")
                #------------- CREATE TABLES LOCAL_POINT_OF_CONTACTS <END>-----------------------
                    
                #------------- CREATE TABLE LOGGING <BEGIN>-----------------------
                sql = """CREATE TABLE LOGGING (
                    User_ID              CHAR(20) NOT NULL,
                    Log_Date             CHAR(20) NOT NULL,
                    Log_Day              INT,
                    Log_Month            INT,
                    Log_Year             INT,
                    Log_Time             CHAR(20) NOT NULL,
                    Log_Hour             INT,
                    Log_Min              INT,
                    Log_Sec              INT,
                    Log_Module           CHAR(50) NOT NULL,
                    Log_Function         CHAR(50) NOT NULL,
                    Executed_by_UserID   CHAR(20))"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table LOGGING was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table LOGGING already Exists")
                #------------- CREATE TABLES LOGGING <END>-----------------------

                #------------- CREATE TABLE Network_Templates <BEGIN>-----------------------
                sql = """CREATE TABLE Network_Templates (
                        Template_ID             CHAR(100) NOT NULL PRIMARY KEY,
                        Template_Name           CHAR(100) NOT NULL,
                        Base_Network_Slash      INT,
                        Split_in_Slash          INT,
                        Executed_by_UserID      CHAR(20) )"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table Network_Templates was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table Network_Templates already Exists")
                #------------- CREATE TABLES Network_Templates <END>-----------------------

                #------------- CREATE TABLE Network_Templates_Break_Out <BEGIN>-----------------------
                sql = """CREATE TABLE Network_Templates_Break_Out (
                        Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                        Template_ID              CHAR(100) NOT NULL,
                        Sequence_1               INT,
                        Sequence_2               INT,
                        Split_in_Slash           INT,
                        VLAN_L3_Type             CHAR(20),
                        VLAN_L3_Name             CHAR(50),
                        Description              CHAR(100),
                        DHCP                     CHAR(10),
                        Split_More               CHAR(10),
                        Split_More_Slash         INT,
                        Executed_by_UserID       CHAR(20) )"""
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table Network_Templates Break Out was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table Network_Templates Break Out already Exists")
                #------------- CREATE TABLES Network_Templates_Break_Out <END>-----------------------

                #------------- CREATE TABLE Networks <BEGIN>-----------------------
                sql = """CREATE TABLE Networks (
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
                        IP_From_Num              FLOAT,
                        IP_To_Num                FLOAT,
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
                if (db.Create_Table(sql)):
                    TablesListBox.insert(END,"Table Networks was Succesful Created.")
                else:
                    TablesListBox.insert(END,"Table Networks already Exists")
                #------------- CREATE TABLES Networks <END>-----------------------

                #------------- CREATE INDEXES <BEGIN>-----------------------
                if (BackendServer == 'SQL'):
                    sql = "CREATE INDEX Idx_ICMP_Date_Time_Loss ON dbo.ICMP (Device_IP,Date_String,Time_String,Percentage_Loss);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 1 on ICMP was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 1 on ICMP was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_ICMP_YMD_HMS_PS ON dbo.ICMP (Year, Month, Day, Hour, Minute, Second, Size_of_Ping);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 2 on ICMP was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 2 on ICMP was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_ICMP_SUMMARY_YMD_HMS ON dbo.ICMP_SUMMARY (Year, Month, Day, Hour, Minute, Second);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 3 on ICMP_SUMMARY was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 3 on ICMP_SUMMARY was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_DEVICES_IP4 ON dbo.DEVICES (IP4_Address);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 4 on DEVICES was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 4 on DEVICES was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_DEVICES_C0_RE_FA_SI_DE_IP_ST ON dbo.DEVICES (Region_ID, Country_ID, Facility_ID, Site_ID,Device_ID, IP4_Address, Status);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                            TablesListBox.insert(END,"INDEX 5 on DEVICES was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 5 on DEVICES was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_TRACEROUTE_SUMMARY_YMD_HMS ON dbo.TRACEROUTE_SUMMARY (Year, Month, Day, Hour, Minute, Second);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 6 on TRACEROUTE_SUMMARY was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 6 on TRACEROUTE_SUMMARY was NOT Succesful Created.")
                    sql = "CREATE INDEX Idx_TRACEROUTE_Date_Time_IP ON dbo.TRACEROUTE (Date_String,Time_String, Device_IP );" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 7 on TRACEROUTE was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 7 on TRACEROUTE was NOT Succesful Created.")
                        
                    sql = "CREATE INDEX Idx_DEVICES_CRF ON dbo.DEVICES (Country_ID, Region_ID, Facility_ID);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 8 on DEVICES was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 8 on DEVICES was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_CIRCUITS_CRF ON dbo.CIRCUITS (Country_ID, Region_ID, Facility_ID);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 9 on CIRCUITS was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 9 on CIRCUITS was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_Logging_UserID ON dbo.LOGGING (User_ID);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 10 on LOGGING was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 10 on LOGGING was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_Logging_Date ON dbo.LOGGING (Log_Date);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 11 on LOGGING was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 11 on LOGGING was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_Logging_Time ON dbo.LOGGING (Log_Time);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 12 on LOGGING was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 12 on LOGGING was NOT Succesful Created.")
                        
                    sql = "CREATE INDEX Idx_Logging_UserID_Date_Time ON dbo.LOGGING (User_ID,Log_Date,Log_Time);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 13 on LOGGING was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 13 on LOGGING was NOT Succesful Created.")

                    sql = "CREATE INDEX Idx_Logging_UserID_Date_Time_Numbers ON dbo.LOGGING (User_ID,Log_Day,Log_Month,Log_year,Log_Hour,Log_Min,Log_Sec);" #<---------- Needs a ";" at the end
                    if (db2.Create_Index(sql)):
                        TablesListBox.insert(END,"INDEX 14 on LOGGING was Succesful Created.")
                    else:
                        TablesListBox.insert(END,"INDEX 14 on LOGGING was NOT Succesful Created.")

                #------------- CREATE INDEXES <END>-----------------------
            else:
                if (DatabaseVersion != BVAnalyticsDatabaseVersion):
                    if db2.Connect():
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.1  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.1"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.1")
                            if (db2.Alter_Table_Add_Field('VARIABLES','Window_Scaling','FLOAT')):
                                TablesListBox.insert(END,"TABLE:[VARIABLES] FIELD: [Window_Scaling] TYPE: [FLOAT]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[VARIABLES] FIELD: [Window_Scaling] TYPE: [*FLOAT*]")
                                All_Fields_Updated = False
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------    
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', Updated_by_UserID = '%s', \
                                       Window_Scaling = '%f'" % ('1.1',date,Username,1.0)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.2  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.2"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.2")
                            #------------- CREATE TABLE DEVICE_MODEL <BEGIN>-------------------
                            sql = """CREATE TABLE DEVICE_MODEL (
                                Device_Model_ID         CHAR(30) NOT NULL PRIMARY KEY,
                                Device_Model_Name       CHAR(50),
                                Vendor_ID               CHAR(20),
                                Outsource_Cost          FLOAT,
                                Executed_by_UserID      CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table DEVICE_MODEL was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table DEVICE_MODEL already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES DEVICE_MODEL <END>-----------------------                            
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------    
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.2',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.3  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.3"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.3")
                            if (db2.Alter_Table_Add_Field('DEVICE_MODEL','Device_Type_ID','CHAR(30)')):
                                TablesListBox.insert(END,"TABLE:[DEVICE_MODEL] FIELD: [Device_Type_ID] TYPE: [CHAR(30)]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICE_MODEL] FIELD: [Device_Type_ID] TYPE: [*CHAR(30)*]")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES DEVICE_MODEL <END>-----------------------                            
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------    
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.3',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.4  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.4"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.4")
                            All_Fields_Updated = False
                            #------------- CREATE TABLES DEVICE_MODEL <END>-----------------------
                            #sql = "DROP INDEX Index_ICMP_1 ON dbo.ICMP" # To DROP INDEXES
                            #if (db2.Drop_Index(sql)):
                            #    TablesListBox.insert(END,"INDEX on ICMP was Succesful Removed.")
                            #else:
                            #    TablesListBox.insert(END,"INDEX on ICMP was NOT Succesful Removed.")                            
                            sql = "CREATE INDEX Idx_ICMP_Date_Time_Loss ON dbo.ICMP (Device_IP,Date_String,Time_String,Percentage_Loss);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 1 on ICMP was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 1 on ICMP was NOT Succesful Created.")
                                All_Fields_Updated = False

                            sql = "CREATE INDEX Idx_ICMP_YMD_HMS_PS ON dbo.ICMP (Year, Month, Day, Hour, Minute, Second, Size_of_Ping);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 2 on ICMP was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 2 on ICMP was NOT Succesful Created.")
                                All_Fields_Updated = False

                            sql = "CREATE INDEX Idx_ICMP_SUMMARY_YMD_HMS ON dbo.ICMP_SUMMARY (Year, Month, Day, Hour, Minute, Second);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 3 on ICMP_SUMMARY was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 3 on ICMP_SUMMARY was NOT Succesful Created.")
                                All_Fields_Updated = False

                            sql = "CREATE INDEX Idx_DEVICES_IP4 ON dbo.DEVICES (IP4_Address);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 4 on DEVICES was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 4 on DEVICES was NOT Succesful Created.")
                                All_Fields_Updated = False

                            sql = "CREATE INDEX Idx_DEVICES_C0_RE_FA_SI_DE_IP_ST ON dbo.DEVICES (Region_ID, Country_ID, Facility_ID, Site_ID,Device_ID, IP4_Address, Status);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 5 on DEVICES was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 5 on DEVICES was NOT Succesful Created.")
                                All_Fields_Updated = False

                            sql = "CREATE INDEX Idx_TRACEROUTE_SUMMARY_YMD_HMS ON dbo.TRACEROUTE_SUMMARY (Year, Month, Day, Hour, Minute, Second);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 6 on TRACEROUTE_SUMMARY was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 6 on TRACEROUTE_SUMMARY was NOT Succesful Created.")
                                All_Fields_Updated = False
                            sql = "CREATE INDEX Idx_TRACEROUTE_Date_Time_IP ON dbo.TRACEROUTE (Date_String,Time_String, Device_IP );" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 7 on TRACEROUTE was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"INDEX 7 on TRACEROUTE was NOT Succesful Created.")
                                All_Fields_Updated = False
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.4',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.5  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.5"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.5")
                            All_Fields_Updated = False
                            #------------- DROP TABLES <BEGIN>-----------------------
                            table_name = ("CIRCUITS")
                            if (db.Drop_Table(table_name)):
                                TablesListBox.insert(END,"OLD Cicuits Table was Deleted.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"OLD Cicuits Table was Not Found.")
                            #------------- DROP TABLES <END>-----------------------
                            #------------- CREATE TABLE CIRCUITS <BEGIN>-----------------------
                            sql = """CREATE TABLE CIRCUITS (
                                Circuit_ID              CHAR(100) NOT NULL PRIMARY KEY,
                                Description             CHAR(200),
                                Country_ID              CHAR(20),                    
                                Region_ID               CHAR(20),
                                Facility_ID             CHAR(20),
                                Site_ID                 CHAR(20),
                                Carrier_ID              CHAR(20),
                                Circuit_Type            CHAR(40),
                                Port_Speed              CHAR(20),
                                Bandwidth               CHAR(20),
                                Status                  CHAR(20),
                                Dmarc_Info_1            CHAR(200),
                                Dmarc_Info_2            CHAR(200),
                                LEC1                    CHAR(50),
                                LEC2                    CHAR(50),
                                LEC3                    CHAR(50),
                                LEC4                    CHAR(50),
                                LEC5                    CHAR(50),
                                CE_ASN                  CHAR(20),
                                CE_IP_Address           CHAR(50),
                                PE_ASN                  CHAR(20),
                                PE_IP_Address           CHAR(50),
                                VLAN_ID                 CHAR(10),
                                PE_Switch               CHAR(100),
                                PE_Location             CHAR(100),
                                NPA_NXX                 CHAR(20),
                                Monthly_Cost            FLOAT,
                                Order_Number            CHAR(40),
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
                                Term_Day                CHAR(10),
                                Term_Time               CHAR(10),
                                ETF                     FLOAT,
                                Contract_No             CHAR(50),
                                Account_No              CHAR(50),                                
                                Executed_by_UserID      CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table CIRCUITS was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table CIRCUITS already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES CIRCUITS <END>-----------------------
                            #------------- CREATE TABLE CARRIER <BEGIN>-----------------------
                            sql = """CREATE TABLE CARRIER (
                                Carrier_ID              CHAR(20) NOT NULL PRIMARY KEY,
                                Description             CHAR(100),
                                Point_Of_Contact        CHAR(50),                    
                                Telephone_NO            CHAR(50),
                                E_Mail_Address          CHAR(100),
                                Web_Site                CHAR(100),
                                Tech_Support_No         CHAR(50),
                                Master_Contract_No      CHAR(50),
                                Executed_by_UserID      CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table CARRIER was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table CARRIER already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES CARRIER <END>-----------------------
                            #------------- CREATE TABLE CIRCUIT_TYPE <BEGIN>-----------------------
                            sql = """CREATE TABLE CIRCUIT_TYPE (
                                Circuit_Type            CHAR(40) NOT NULL PRIMARY KEY,
                                Description             CHAR(50),
                                Executed_by_UserID      CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table CIRCUIT_TYPE was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table CIRCUIT_TYPE already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES CIRCUIT_TYPE <END>-----------------------
                            #------------- CREATE TABLE PORT_SPEED <BEGIN>-----------------------
                            sql = """CREATE TABLE PORT_SPEED (
                                Port_Speed_ID           CHAR(20) NOT NULL PRIMARY KEY,
                                Description             CHAR(50),
                                Executed_by_UserID      CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table PORT_SPEED was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table PORT_SPEED already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES PORT_SPEED <END>-----------------------
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.5',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.5a  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.5a"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.6")
                            sql = "CREATE INDEX Idx_DEVICES_CRF ON dbo.DEVICES (Country_ID, Region_ID, Facility_ID);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 8 on DEVICES was Succesful Created.")
                            else:
                                TablesListBox.insert(END,"INDEX 8 on DEVICES was NOT Succesful Created.")

                            sql = "CREATE INDEX Idx_CIRCUITS_CRF ON dbo.CIRCUITS (Country_ID, Region_ID, Facility_ID);" #<---------- Needs a ";" at the end
                            if (db2.Create_Index(sql)):
                                TablesListBox.insert(END,"INDEX 9 on CIRCUITS was Succesful Created.")
                            else:
                                TablesListBox.insert(END,"INDEX 9 on CIRCUITS was NOT Succesful Created.")
                                
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.6  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.6"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.6")
                            #------------- CREATE TABLE LOCAL_POINT_OF_CONTACTS <BEGIN>-----------------------
                            sql = """CREATE TABLE LOCAL_POINT_OF_CONTACTS (
                                Country_ID_Region_ID_Facility_ID_SEQ_No     CHAR(85) NOT NULL PRIMARY KEY,
                                Country_ID                                  CHAR(20) NOT NULL,
                                Region_ID                                   CHAR(20) NOT NULL,
                                Facility_ID                                 CHAR(20) NOT NULL,
                                SEQ_No                                      INT NOT NULL,
                                Primary_POC_Name                            CHAR(100),
                                Primary_POC_Office_No                       CHAR(100),
                                Primary_POC_Alternate_No                    CHAR(100),
                                Primary_POC_2nd_Alternate_No                CHAR(100),
                                Secondary_POC_Name                          CHAR(100),
                                Secondary_POC_Office_No                     CHAR(100),
                                Secondary_POC_Alternate_No                  CHAR(100),
                                Secondary_POC_2nd_Alternate_No              CHAR(100),
                                Last_time_it_was_Updated                    CHAR(100),
                                Executed_by_UserID                          CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table LOCAL_POINT_OF_CONTACTS was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table LOCAL_POINT_OF_CONTACTS already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES LOCAL_POINT_OF_CONTACTS <END>-----------------------

                            if (db.Alter_Table_Add_Field('DEVICES','NAT_IP4_Address','CHAR(20)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [NAT_IP4_Address] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [NAT_IP4_Address] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Management_IP4_Address','CHAR(20)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Management_IP4_Address] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Management_IP4_Address] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_Success_ICMP','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Success_ICMP] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Success_ICMP] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_ICMP_Status','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_ICMP_Status] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_ICMP_Status] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.6',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------                                
                                
                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.7  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.7"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.7")
                            #------------- CREATE TABLE LOGGING <BEGIN>-----------------------
                            sql = """CREATE TABLE LOGGING (
                                User_ID              CHAR(20) NOT NULL,
                                Log_Date             CHAR(20) NOT NULL,
                                Log_Day              INT,
                                Log_Month            INT,
                                Log_Year             INT,
                                Log_Time             CHAR(20) NOT NULL,
                                Log_Hour             INT,
                                Log_Min              INT,
                                Log_Sec              INT,
                                Log_Module           CHAR(50) NOT NULL,
                                Log_Function         CHAR(50) NOT NULL,
                                Executed_by_UserID   CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table LOGGING was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table LOGGING already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES LOGGING <END>-----------------------

                            #------------- CREATE TABLES INDEXES <BEGIN>-----------------------
                            if (BackendServer == 'SQL'):
                                sql = "CREATE INDEX Idx_Logging_UserID ON dbo.LOGGING (User_ID);" #<---------- Needs a ";" at the end
                                if (db2.Create_Index(sql)):
                                    TablesListBox.insert(END,"INDEX 10 on LOGGING was Succesful Created.")
                                    All_Fields_Updated = True
                                else:
                                    TablesListBox.insert(END,"INDEX 10 on LOGGING was NOT Succesful Created.")
                                    All_Fields_Updated = False

                                sql = "CREATE INDEX Idx_Logging_Date ON dbo.LOGGING (Log_Date);" #<---------- Needs a ";" at the end
                                if (db2.Create_Index(sql)):
                                    TablesListBox.insert(END,"INDEX 11 on LOGGING was Succesful Created.")
                                    All_Fields_Updated = True
                                else:
                                    TablesListBox.insert(END,"INDEX 11 on LOGGING was NOT Succesful Created.")
                                    All_Fields_Updated = False

                                sql = "CREATE INDEX Idx_Logging_Time ON dbo.LOGGING (Log_Time);" #<---------- Needs a ";" at the end
                                if (db2.Create_Index(sql)):
                                    TablesListBox.insert(END,"INDEX 12 on LOGGING was Succesful Created.")
                                    All_Fields_Updated = True
                                else:
                                    TablesListBox.insert(END,"INDEX 12 on LOGGING was NOT Succesful Created.")
                                    All_Fields_Updated = False

                                sql = "CREATE INDEX Idx_Logging_UserID_Date_Time ON dbo.LOGGING (User_ID,Log_Date,Log_Time);" #<---------- Needs a ";" at the end
                                if (db2.Create_Index(sql)):
                                    TablesListBox.insert(END,"INDEX 13 on LOGGING was Succesful Created.")
                                    All_Fields_Updated = True
                                else:
                                    TablesListBox.insert(END,"INDEX 13 on LOGGING was NOT Succesful Created.")
                                    All_Fields_Updated = False

                                sql = "CREATE INDEX Idx_Logging_UserID_Date_Time_Numbers ON dbo.LOGGING (User_ID,Log_Day,Log_Month,Log_year,Log_Hour,Log_Min,Log_Sec);" #<---------- Needs a ";" at the end
                                if (db2.Create_Index(sql)):
                                    TablesListBox.insert(END,"INDEX 14 on LOGGING was Succesful Created.")
                                    All_Fields_Updated = True
                                else:
                                    TablesListBox.insert(END,"INDEX 14 on LOGGING was NOT Succesful Created.")
                                    All_Fields_Updated = False


                            #------------- CREATE TABLES INDEXES <END>-----------------------

                            #------------- ADD FIELDS TO TABLES <BEGIN>-----------------------

                            if (db.Alter_Table_Add_Field('DEVICES','ICMP_Capable','CHAR(10)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [ICMP_Capable] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [ICMP_Capable] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Monitor_via_ICMP','CHAR(10)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Monitor_via_ICMP] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Monitor_via_ICMP] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_Updated_CMDB_Date','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Date] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Date] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_Updated_CMDB_Day','INT')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Day] TYPE: [INT]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Day] TYPE: [*INT*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_Updated_CMDB_Month','INT')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Month] TYPE: [INT]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Month] TYPE: [*INT*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('DEVICES','Last_Updated_CMDB_Year','INT')):
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Year] TYPE: [INT]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[DEVICES] FIELD: [Last_Updated_CMDB_Year] TYPE: [*INT*]")
                                All_Fields_Updated = False

                            #------------- ADD FIELDS TO TABLES <BEGIN>-----------------------

                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.7',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------                                

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.8  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.8"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.8")
                            #------------- CREATE TABLE DEVICE_MODEL_COUNTRY <BEGIN>-----------------------
                            sql = """CREATE TABLE DEVICE_MODEL_COUNTRY (
                                Device_Model_ID_Country_ID      CHAR(60) NOT NULL PRIMARY KEY,
                                Device_Model_ID                 CHAR(30) NOT NULL,
                                Country_ID                      CHAR(20) NOT NULL,
                                Outsource_Cost                  FLOAT,
                                Executed_by_UserID              CHAR(20))"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table DEVICE_MODEL_COUNTRY was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table DEVICE_MODEL_COUNTRY already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES DEVICE_MODEL_COUNTRY <END>-----------------------

                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.8',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------                                

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.9  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.9"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.9")
                            
                            if (db.Alter_Table_Add_Field('VENDOR','Point_of_contact','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Point_of-contact] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Point_of_contact] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('VENDOR','Tel_Point_of_Contact','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Tel_Point_of_Contact] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Tel_Point_of_Contact] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('VENDOR','Email_Point_of_Contact','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Email_Point_of_Contact] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[VENDOR] FIELD: [Email_Point_of_Contact] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.9',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------                                

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.10  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.10"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.10")
                            
                            if (db.Alter_Table_Add_Field('FACILITY','Strategy','CHAR(250)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Strategy] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Strategy] TYPE: [*CHAR*]")
                                All_Fields_Updated = False
     
                            if (db.Alter_Table_Add_Field('FACILITY','Action','CHAR(250)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Action] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Action] TYPE: [*CHAR*]")
                                All_Fields_Updated = False
                                
                            if (db.Alter_Table_Add_Field('FACILITY','Notes','CHAR(250)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Notes] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [Notes] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_Address1','CHAR(100)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Address1] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Address1] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_Address2','CHAR(100)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Address2] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Address2] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_City','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_City] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_City] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_State','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_State] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_State] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_Zip_Code','CHAR(20)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Zip_Code] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Zip_Code] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            if (db.Alter_Table_Add_Field('FACILITY','New_Country','CHAR(50)')):
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Country] TYPE: [CHAR]")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"TABLE:[FACILITY] FIELD: [New_Country] TYPE: [*CHAR*]")
                                All_Fields_Updated = False

                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.10',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------                                

                        #*********************************************************************************
                        #-----------------------------  Database Ver 1.11  -------------------------------*
                        #*********************************************************************************
                        All_Fields_Updated = False
                        if (BVAnalyticsDatabaseVersion == "1.11"):
                            TablesListBox.insert(END,"NEW UPDATES With Version 1.11")
                            #------------- CREATE TABLE Network_Templates <BEGIN>-----------------------
                            sql = """CREATE TABLE Network_Templates (
                                    Template_ID             CHAR(100) NOT NULL PRIMARY KEY,
                                    Template_Name           CHAR(100) NOT NULL,
                                    Base_Network_Slash      INT,
                                    Split_in_Slash          INT,
                                    Executed_by_UserID      CHAR(20) )"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table Network_Templates was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table Network_Templates already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES Network_Templates <END>-----------------------

                            #------------- CREATE TABLE Network_Templates_Break_Out <BEGIN>-----------------------
                            sql = """CREATE TABLE Network_Templates_Break_Out (
                                    Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                                    Template_ID              CHAR(100) NOT NULL,
                                    Sequence_1               INT,
                                    Sequence_2               INT,
                                    Split_in_Slash           INT,
                                    VLAN_L3_Type             CHAR(20),
                                    VLAN_L3_Name             CHAR(50),
                                    Description              CHAR(100),
                                    DHCP                     CHAR(10),
                                    Split_More               CHAR(10),
                                    Split_More_Slash         INT,
                                    Executed_by_UserID       CHAR(20) )"""
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table Network_Templates Break Out was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table Network_Templates Break Out already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES Network_Templates_Break_Out <END>-----------------------
                                
                            #------------- CREATE TABLE Networks <BEGIN>-----------------------
                            sql = """CREATE TABLE Networks (
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
                                    IP_From_Num              FLOAT,
                                    IP_To_Num                FLOAT,
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
                            if (db.Create_Table(sql)):
                                TablesListBox.insert(END,"Table Networks was Succesful Created.")
                                All_Fields_Updated = True
                            else:
                                TablesListBox.insert(END,"Table Networks already Exists")
                                All_Fields_Updated = False
                            #------------- CREATE TABLES Networks <END>-----------------------
                        
                            #------------------------ UPDATE the VARIABLES SETTINGS ----------------------------
                            #All_Fields_Updated = False
                            if (All_Fields_Updated):
                                #------------- UPDATE VARIABLES data into a Table <BEGIN> ---------
                                date = datetime.datetime.now()
                                sql = "UPDATE VARIABLES SET Database_ver = '%s', Last_Time_Updated = '%s', \
                                      Updated_by_UserID = '%s' " % ('1.11',date,Username)
                                if (db.Add_Move_Change_Data(sql)):
                                    TablesListBox.insert(END,"Record Updated to the Variables Table.")
                                else:
                                    TablesListBox.insert(END,"Error adding the record to the Variables Table, posible dupliated it")
                                #------------- UPDATE VARIABLES data into a Table <END> ---------
                            

#######################             
                        db2.Disconnect()    
                else:
                    TablesListBox.insert(END,"NO NEEDED UPDATES")
            db.Disconnect()
            TablesListBox.select_set(0)
#            TablesListBox.bind("<Double-Button-1>",Display_Window) # Double Clcik
#            TablesListBox.bind("<<ListboxSelect>>", OnSelect_ICMP(TablesListBox))    # Single Click or Keys up and down too.
#           TablesListBox.bind("<<ListboxSelect>>", OnSelect('database'))    # Single Click or Keys up and down too.
            TablesWindow.mainloop()
        else:
            db.Disconnect()
            TablesLabel["text"] = "Updating TABLES and FIELDS into the Database"
            TablesLabel.pack()
            #TablesListBox.insert(11,"MORE......")
            #TablesListBox.pack()
    else:
        mbox.showerror(master=root,title='Update Tables and Fields',
                         message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                   + 'Please make sure the ODBC DSN Name mathes: ['
                                   + ODBC_DSN_name + "]")


####################################################################################
#----------------------------- Database Menu <END> ------------------------------#
####################################################################################
    

#================ Menu Subrutines <END>=============




#####################################################
#---------------- Find Menus <BEGIN> -----------#
#####################################################
def Display_Find_Circuits_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Find Circuits Window','OPEN Window']    
        Logging.Log(Parameter)
    FindCEF.Find_Circuits_Window()

def Display_Find_Devices_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Find Devices Window','OPEN Window']    
        Logging.Log(Parameter)
    FindCEF.Find_Devices_Window()

def Display_Find_Facility_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Find Facilities Window','OPEN Window']    
        Logging.Log(Parameter)
    FindCEF.Find_Facility_Window()

def Display_Find_LocalPointOfContacts_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Find Local POC Window','OPEN Window']    
        Logging.Log(Parameter)
    FindLocalPointOfContacts.Find_LocalPointOfContacts_Window()
    
def Display_Find_Networks_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Find Networks Window','OPEN Window']    
        Logging.Log(Parameter)
    FindNetworks.Find_Networks_Window()

###################################################
#---------------- Circuits Menus <END> -----------#
###################################################


#####################################################
#---------------- Location Menus <BEGIN> -----------#
#####################################################

def Display_Country_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Country Window','OPEN Window']    
        Logging.Log(Parameter)
    Country.Display_Country_Window()

def Display_Region_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Region Window','OPEN Window']    
        Logging.Log(Parameter)
    Region.Display_Region_Window()

def Display_Facility_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Facility Window','OPEN Window']    
        Logging.Log(Parameter)
    Facility.Display_Facility_Window()

def Display_Sites_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Sites Window','OPEN Window']    
        Logging.Log(Parameter)
    Sites.Display_Sites_Window()

def Display_LocalPointOfContacts_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Local POC Window','OPEN Window']    
        Logging.Log(Parameter)
    LocalPointOfContacts.Display_LocalPointOfContacts_Window()
    
###################################################
#---------------- Location Menus <END> -----------#
###################################################

def Display_Device_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Devices Window','OPEN Window']    
        Logging.Log(Parameter)
    Device.Display_Device_Window()

def Display_Vendor_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Vendor Window','OPEN Window']    
        Logging.Log(Parameter)
    Vendor.Display_Vendor_Window()


def Display_DeviceType_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Device Type Window','OPEN Window']    
        Logging.Log(Parameter)
    DeviceType.Display_DeviceType_Window()

def Display_DeviceModel_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Device Model Window','OPEN Window']    
        Logging.Log(Parameter)
    DeviceModel.Display_DeviceModel_Window()

def Display_ImportCMDB_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics ImportCMDB Window','OPEN Window']    
        Logging.Log(Parameter)
    ImportCMDB.Display_ImportCMDB_Window()

#####################################################
#---------------- Circuits Menus <BEGIN> -----------#
#####################################################
def Display_Carrier_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Carrier Window','OPEN Window']    
        Logging.Log(Parameter)
    Carrier.Display_Carrier_Window()

def Display_CircuitType_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics CircuitType Window','OPEN Window']    
        Logging.Log(Parameter)
    CircuitType.Display_CircuitType_Window()

def Display_PortSpeed_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics PortSpeed Window','OPEN Window']    
        Logging.Log(Parameter)
    PortSpeed.Display_PortSpeed_Window()


def Display_Circuit_Window():    
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Circuits Window','OPEN Window']    
        Logging.Log(Parameter)
    Circuit.Display_Circuits_Window()

###################################################
#---------------- Circuits Menus <END> -----------#
###################################################


##############################################
#---------------- Reports <BEGIN> -----------#
##############################################

def Report_Outsource_Per_Country():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Report Outsouirce Per Country Window','OPEN Window']    
        Logging.Log(Parameter)
    ReportOutsourcePerCountry.Find_ReportOutsourcePerCountry_Window()

def Report_Outsource_Per_Facility():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Report Outsource Per Facility Window','OPEN Window']    
        Logging.Log(Parameter)
    ReportOutsourcePerFacility.Find_ReportOutsourcePerFacility_Window()


############################################
#---------------- Reports <END> -----------#
############################################


##############################################
#---------------- Network <BEGIN> -----------#
##############################################

def Display_NetworkTemplates_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics NetworkTemplates Window','OPEN Window']    
        Logging.Log(Parameter)
    NetworkTemplates.Display_NetworkTemplates_Window()

def Display_NetworkTemplatesBreakOut_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics NetworkTemplatesBreakOut Window','OPEN Window']    
        Logging.Log(Parameter)
    NetworkTemplatesBreakOut.Display_NetworkTemplatesBreakOut_Window()

def Display_NetworkAllocations_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics NetworkAllocations Window','OPEN Window']    
        Logging.Log(Parameter)
    NetworkAllocations.Display_NetworkAllocations_Window()

def Display_NetworkProvisioning_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics NetworkProvisioning Window','OPEN Window']    
        Logging.Log(Parameter)
    NetworkProvisioning.Display_NetworkProvisioning_Window()
    

############################################
#---------------- Network <END> -----------#
############################################


##############################################
#---------------- Logging <BEGIN> -----------#
##############################################

def Find_Logging():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Loging Window','OPEN Window']    
        Logging.Log(Parameter)
    FindLogging.Find_Logging_Window()

############################################
#---------------- Logging <END> -----------#
############################################


###########################################################        
#---------------- Response Times Menus <BEGIN> -----------#
###########################################################

def on_ICMPSummaryWindow_quit():
    global ICMPSummaryWindowExist
    global ICMPSummaryWindow
    
    ICMPSummaryWindowExist = False
    db.Disconnect()
    db2.Disconnect()
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics ICMP Summary Window','CLOSE Window']    
        Logging.Log(Parameter)
    ICMPSummaryWindow.destroy()


def On_ICMPSummary_Tree_Double_ClickWindow_quit():
    global PercentageLossWindowExist
    global PercentageLossWindow
    
    PercentageLossWindowExist = False
    db.Disconnect()
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Percentage Loss Window','CLOSE Window']    
        Logging.Log(Parameter)
    PercentageLossWindow.destroy()

def on_Percentage_Loss_Treeview_double_click(event):
    #print ("double Click")
    global PercentageLossTreeview
    global PercentageLossFrame
    
    curItem = PercentageLossTreeview.focus()
    dic = PercentageLossTreeview.item(curItem)
    values = dic.get('values')
    if (len(values) > 0):
        #print (values)
        device_id = values[3].strip()
        if (db.Connect()):
            sql = """
                SELECT * FROM DEVICES
                WHERE Device_ID = '%s'
                """ % (device_id)
            if (db.Execute(sql)):
                    i = 0
                    while (i < len(db.results)):
                        if (db.results[i][39] == None): # The Dafault is YES
                            sql = "UPDATE DEVICES SET ICMP_Capable = 'NO' WHERE Device_ID = '%s'" %(device_id)
                            if (db.Add_Move_Change_Data(sql)):
                                    mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                    message = 'Device ['+device_id+'] Changed ICMP Capable to NO')            
                            else:
                                    mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                    message = 'Device ['+device_id+'] NOT Found')    
                        else:
                            if (db.results[i][39].strip() == 'YES'):
                                sql = "UPDATE DEVICES SET ICMP_Capable = 'NO' WHERE Device_ID = '%s'" %(device_id)
                                if (db.Add_Move_Change_Data(sql)):
                                        mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                        message = 'Device ['+device_id+'] Changed ICMP Capable to NO')            
                                else:
                                        mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                        message = 'Device ['+device_id+'] NOT Found')                            
                            else:
                                sql = "UPDATE DEVICES SET ICMP_Capable = 'YES' WHERE Device_ID = '%s'" %(device_id)
                                if (db.Add_Move_Change_Data(sql)):
                                        mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                        message = 'Device ['+device_id+'] Changed ICMP Capable to YES')            
                                else:
                                        mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                                        message = 'Device ['+device_id+'] NOT Found')
                        i = i + 1                                    
            else:
                mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                message = 'Device ['+device_id+'] NOT Found')                            
                
        else:
            mbox.showerror(master=root,title='Percentage Loss and TTL',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + ODBC_DSN_name + "]")



def ICMPSummary_Double_Click_Call_Save_As():
    global PercentageLossTreeviewDataColumns
    global ICMPSummaryArray

    data_ready = True
    if (data_ready):
        #print ("Save as")
        input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
        if input_file_name != "":
            #version = '4.0'
            file_name = input_file_name
            ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,file_name,version)
            Tab0 = "Percentage Loss and TTL"
            Tabs = [Tab0]
            ExcellFile.Call_Write_to_File(Tabs)
            Row = 2    # 3
            Column = 1 # A
            ExcellFile.Add_DataToWorksheet(PercentageLossTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
            Row = 3
            #Column = 1
            i = 0
            while (i < len(ICMPSummaryArray)):                                                                            #<------ Data for the Columns
                num = i + 1
                tags = "-" # To use in the futire
                item = ICMPSummaryArray[i]
                '''
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
                '''
                ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                i = i + 1
                Row = Row + 1
                #Column = 1
            if (ExcellFile.Save_File()):
                mbox.showinfo(master=PercentageLossFrame,title='Percentage Loss and TTL',
                        message = '!!! The File was saved !!!')
                PercentageLossWindow.title("Percentage Loss and TTL                    File: ["+file_name+"] SAVED")                  #<---- Window Name Change
                if Is_logging_Available:
                    Parameter = []
                    Parameter = ['Percentage Loss and TTL','SAVE AS']    
                    Logging.Log(Parameter)
            else:
                mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                        message = '*** The File was not saved, Perhaps It is already open ***')



def On_ICMPSummary_Tree_Double_Click(event):
    global ICMPSummaryTreeview              # The ICMP Summary Tree
    global PercentageLossWindow
    global PercentageLossWindowExist
    global PercentageLossTreeview
    global PercentageLossFrame
    global PercentageLossTreeviewDataColumns
    global ICMPSummaryArray

    if db.Connect():
        # Set up the Window
        PercentageLossWindowExist = False # to allow multiple windows but the user needs to close them all
        PercentageLossWindow = Tk()
        PercentageLossWindow.geometry('1350x700+350+70')
        #PercentageLossWindow.protocol("WM_DELETE_WINDOW", On_ICMPSummary_Tree_Double_ClickWindow_quit)
        PercentageLossWindow.call('tk', 'scaling', Windows_Scaling)
        PercentageLossWindow.title("Percentage Loss and TTL")
        PercentageLossLabel = Label(PercentageLossWindow,text="Helvetica", font=("Helvetica", 13))
        PercentageLossLabel["text"] = "Percentage Loss and TTL"
        PercentageLossLabel.pack()

        data_ready = False

        #------------------ MENU <BEGIN> ----------------------------------------------------------
        menubar = Menu(PercentageLossWindow)
        filemenu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=filemenu)
        if Is_SaveAs_Available:
            filemenu.add_command(label="Save As", command=ICMPSummary_Double_Click_Call_Save_As)
        filemenu.add_separator()
        #filemenu.add_command(label="Exit", command=On_ICMPSummary_Tree_Double_ClickWindow_quit)
        PercentageLossWindow.config(menu=menubar)
        #------------------ MENU <END> ----------------------------------------------------------

            
        # Setup Frame
        PercentageLossFrame = Frame(PercentageLossWindow)
        PercentageLossFrame.pack(side=TOP, fill=BOTH, expand=Y)

        # set frame resizing priorities
        PercentageLossFrame.rowconfigure(0, weight=1)
        PercentageLossFrame.columnconfigure(0, weight=1)

        #------------------ TREE VIEW For Traceroute Database <BEGIN> -----------------------------------
        # Create Tree and Scrollbars
        PercentageLossTreeviewDataColumns = ('Date',
                           'Time',
                           'IP Address',
                           'Device Name',
                           'Device Type',
                           'Percentage Loss',
                           'Status',
                           'Package Size',
                           'Country ID',
                           'Region ID',
                           'Facility ID')
            
        PercentageLossTreeview = ttk.Treeview(PercentageLossFrame,columns=PercentageLossTreeviewDataColumns, height=16)     # <--- Make sure the frame is correct !!
        PercentageLossTreeviewysb = Scrollbar(PercentageLossFrame,orient=VERTICAL, command=PercentageLossTreeview.yview)    # <--- Make sure the frame is correct !!
        PercentageLossTreeviewxsb = Scrollbar(PercentageLossFrame,orient=HORIZONTAL, command=PercentageLossTreeview.xview)        # <--- Make sure the frame is correct !!
        PercentageLossTreeview['yscroll'] = PercentageLossTreeviewysb.set
        PercentageLossTreeview['xscroll'] = PercentageLossTreeviewxsb.set

        # setup headings and column 
        PercentageLossTreeview.heading('#0', text='Item No.',           anchor=W) # E for East and W for West, CENTER
        PercentageLossTreeview.heading('#1', text='Date',               anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Date', False)) # E for East and W for West
        PercentageLossTreeview.heading('#2', text='Time',               anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Time', False)) 
        PercentageLossTreeview.heading('#3', text='IP Address',         anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'IP Address', False))
        PercentageLossTreeview.heading('#4', text='Device Name',        anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Device Name', False))        
        PercentageLossTreeview.heading('#5', text='Device Type',        anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Device Type', False))
        PercentageLossTreeview.heading('#6', text='Percentage Loss',    anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Percentage Loss', False))
        PercentageLossTreeview.heading('#7', text='Status',             anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Status', False))
        PercentageLossTreeview.heading('#8', text='Package Size',       anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Package Size', False))
        PercentageLossTreeview.heading('#9', text='Country ID',         anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Country ID', False))
        PercentageLossTreeview.heading('#10', text='Region ID',          anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Region ID', False))
        PercentageLossTreeview.heading('#11', text='Facility ID',        anchor=W,command=lambda: treeview_sort_column(PercentageLossTreeview, 'Facility ID', False))
            
        PercentageLossTreeview.column('#0', stretch=1, width=3 , anchor=W)
        PercentageLossTreeview.column('#1', stretch=1, width=10, anchor=W)
        PercentageLossTreeview.column('#2', stretch=1, width=10)
        PercentageLossTreeview.column('#3', stretch=1, width=20, anchor=W)
        PercentageLossTreeview.column('#4', stretch=1, width=3)
        PercentageLossTreeview.column('#5', stretch=1, width=3)
        PercentageLossTreeview.column('#6', stretch=1, width=20)
        PercentageLossTreeview.column('#7', stretch=1, width=20)
        PercentageLossTreeview.column('#8', stretch=1, width=20)
        PercentageLossTreeview.column('#9', stretch=1, width=20)
        PercentageLossTreeview.column('#10', stretch=1, width=20)
        PercentageLossTreeview.column('#11', stretch=1, width=20)

        # add tree and scrollbars to frame
        PercentageLossTreeview.grid(row=0, column=0, sticky=NSEW)
        PercentageLossTreeviewysb.grid(row=0, column=1, sticky=NS)
        PercentageLossTreeviewxsb.grid(row=1, column=0, sticky=EW)

        # create fonts and tags for future use to compare Region and delay
        PercentageLossTreeview.tag_configure('25', font=('Helvetica', 8), background='LightBlue1')
        PercentageLossTreeview.tag_configure('50', font=('Helvetica', 8), background='coral1')
        PercentageLossTreeview.tag_configure('75', font=('Helvetica', 8), background='pink')
        PercentageLossTreeview.tag_configure('TTL', font=('Helvetica', 8), background='red')
            
        # Bind the double Click
        #TracerouteTreeview.bind('<ButtonRelease-1>', on_ICMP_Tree_select_click)
        PercentageLossTreeview.bind("<Triple-Button-1>", on_Percentage_Loss_Treeview_double_click) # Triple Right 
        PercentageLossTreeview.bind("<Double-3>", on_Percentage_Loss_Treeview_double_click) # Double Left

        curItem = ICMPSummaryTreeview.focus()
        dic = ICMPSummaryTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            #print (values)
            date = values[0].strip()
            time = values[1].strip()
            # SQL Querry to the ICMP Table
            if (db.Connect() and db2.Connect()):
                sql = """
                        SELECT * FROM ICMP
                        WHERE Date_String = '%s' AND Time_String = '%s' AND (Percentage_Loss > 0 AND Percentage_Loss <= 100) \
                        ORDER BY Percentage_Loss desc, Device_IP ASC
                      """ % (date,time)
                #print (sql)
                if (db.Execute(sql)):
                    i = 0
                    ICMPSummaryArray = []
                    while (i < len(db.results)):
                        num = i + 1
                        tags = "regions"
                        if (db.results[i][11] == 25):
                            tags = "25"
                        if (db.results[i][11] == 50):
                            tags = "50"
                        if (db.results[i][11] == 75):
                            tags = "75"
                        if ((db.results[i][15].find("TTL expire",0) != -1)): 
                            tags = "TTL"
                        DateStr = db.results[i][2]
                        TimeStr = db.results[i][3]
                        Device_IP = db.results[i][1].strip()
                        PercentageLoss = db.results[i][11]
                        Status = db.results[i][15]
                        PackageSize = db.results[i][10]
                        CountryID = '*'
                        RegionID = '*'
                        FacilityID = '*'
                        DeviceName = " "
                        DeviceType = " "
                        
                        sql = """
                                SELECT * FROM DEVICES 
                                WHERE IP4_Address = '%s'
                              """ % (Device_IP)
                        #print (sql)

                        if (db2.Execute(sql)):
                            if (len(db.results) > 0):
                                #print ('found')
                                CountryID = db2.results[0][2]
                                RegionID = db2.results[0][3]
                                FacilityID = db2.results[0][4]
                                DeviceName = db2.results[0][0]
                                DeviceType = db2.results[0][6]
                                
                        item = [DateStr,
                                TimeStr,
                                Device_IP,
                                DeviceName,
                                DeviceType,
                                PercentageLoss,
                                Status,
                                PackageSize,
                                CountryID,
                                RegionID,
                                FacilityID]
                        PercentageLossTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                        i = i + 1
                        ICMPSummaryArray.append(item)
                else:
                    mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                    message = 'No Records found')
                #db.Disconnect()
            else:
                mbox.showerror(master=PercentageLossFrame,title='Percentage Loss and TTL',
                message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                        + 'Please make sure the ODBC DSN Name mathes: ['
                        + ODBC_DSN_name + "]")            
            #db.Disconnect()
            PercentageLossWindow.mainloop()
        #else:
            #db.Disconnect()
    else:
        mbox.showerror(master=root,title='Percentage Loss and TTL',
                         message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                   + 'Please make sure the ODBC DSN Name mathes: ['
                                   + ODBC_DSN_name + "]")


def ICMPSummary_Call_Save_As():
    global data_ready
    global ICMPSummaryTreeviewDataColumns
    global ICMPSummaryWindowExist
    global ICMPSummaryWindow
    global ICMPSummaryLabel
    global ICMPSummaryTreeview
    global ICMPSummaryFrameEntry
    global ICMPSummaryFrame
    global ICMPSummaryTreeviewDataColumns
    
    data_ready = True
    if (data_ready):
        #print ("Save as")
        input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
        if input_file_name != "":
            file_name = input_file_name
            #version = "4.0"
            ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,file_name,version)
            Tab0 = "ICMP Summary Report"
            Tabs = [Tab0]
            ExcellFile.Call_Write_to_File(Tabs)
            Row = 2    # 3
            Column = 1 # A
            ExcellFile.Add_DataToWorksheet(ICMPSummaryTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
            Row = 3
            #Column = 1
            i = 0
            # SQL Querry to the COUNTRY Table
            sql = """
                        SELECT * FROM ICMP_SUMMARY
                        ORDER BY Year desc, Month desc, Day desc, Hour desc, Minute desc, Second desc """
            if (db.Execute(sql)):
                i = 0
                while (i < len(db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [db.results[i][1],
                            db.results[i][2],
                            db.results[i][9],
                            db.results[i][10],
                            db.results[i][11],
                            db.results[i][12],
                            db.results[i][14],
                            db.results[i][15],
                            db.results[i][16],
                            db.results[i][17],
                            db.results[i][18]]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=ICMPSummaryFrame,title='ICMP Summary Report',
                            message = '!!! The File was saved !!!')
                    ICMPSummaryWindow.title("ICMP Summary Report                    File: ["+file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['ICMP Summary Report','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=ICMPSummaryFrame,title='ICMP Summary Report',
                            message = '*** The File was not saved, Perhaps It is already open ***')
            else:
                mbox.showerror(master=ICMPSummaryFrame,title='ICMP Summary Report',
                message = 'No Records found')

    
def Display_ICMP_Summary_Window():
    global ICMPSummaryWindowExist
    global ICMPSummaryWindow
    global ICMPSummaryLabel
    global ICMPSummaryTreeview
    global ICMPSummaryFrameEntry
    global ICMPSummaryFrame
    global ICMPSummaryTreeviewDataColumns
    
    if db.Connect():
        if not ICMPSummaryWindowExist:
            # Set up the Window
            ICMPSummaryWindowExist = True
            ICMPSummaryWindow = Tk()
            ICMPSummaryWindow.geometry('1350x700+350+70')
            ICMPSummaryWindow.title("ICMP Summary Report")
            ICMPSummaryWindow.protocol("WM_DELETE_WINDOW", on_ICMPSummaryWindow_quit)
            ICMPSummaryWindow.call('tk', 'scaling', Windows_Scaling)
            ICMPSummaryLabel = Label(ICMPSummaryWindow,text="Helvetica", font=("Helvetica", 13))
            ICMPSummaryLabel["text"] = "ICMP Summary Report"
            ICMPSummaryLabel.pack()

            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(ICMPSummaryWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=ICMPSummary_Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=on_ICMPSummaryWindow_quit)
            ICMPSummaryWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------

            
            # Setup Frame
            ICMPSummaryFrame = Frame(ICMPSummaryWindow)
            ICMPSummaryFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            ICMPSummaryFrame.rowconfigure(0, weight=1)
            ICMPSummaryFrame.columnconfigure(0, weight=1)

            #------------------ TREE VIEW For ICMP Summary Database <BEGIN> -----------------------------------
            # Create Tree and Scrollbars

            ICMPSummaryTreeviewDataColumns = ('Date','Time', 'Total No. Of PINGs','Normal PINGs','Time Out PINGs','TTL Expired PINGs',
                                              '0% Loss', '25% Loss','50% Loss','75% Loss','100% Loss')
            
            ICMPSummaryTreeview = ttk.Treeview(ICMPSummaryFrame,columns=ICMPSummaryTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
            ICMPSummaryTreeviewysb = Scrollbar(ICMPSummaryFrame,orient=VERTICAL, command=ICMPSummaryTreeview.yview)           # <--- Make sure the frame is correct !!
            ICMPSummaryTreeviewxsb = Scrollbar(ICMPSummaryFrame,orient=HORIZONTAL, command=ICMPSummaryTreeview.xview)         # <--- Make sure the frame is correct !!
            ICMPSummaryTreeview['yscroll'] = ICMPSummaryTreeviewysb.set
            ICMPSummaryTreeview['xscroll'] = ICMPSummaryTreeviewxsb.set

            # setup headings and column 
            ICMPSummaryTreeview.heading('#0', text='Item No.',          anchor=W) # E for East and W for West
            ICMPSummaryTreeview.heading('#1', text='Date',              anchor=W,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'Date', False)) # E for East and W for West
            ICMPSummaryTreeview.heading('#2', text='Time',              anchor=W,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'Time', False)) 
            ICMPSummaryTreeview.heading('#3', text='Total No. Of PINGs',anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'Total No. Of PINGs', False)) 
            ICMPSummaryTreeview.heading('#4', text='Normal PINGs',      anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'Normal PINGs', False)) 
            ICMPSummaryTreeview.heading('#5', text='Time Out PINGs',    anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'Time Out PINGs', False)) 
            ICMPSummaryTreeview.heading('#6', text='TTL Expired PINGs', anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, 'TTL Expired PINGs', False))
            ICMPSummaryTreeview.heading('#7', text='0% Loss',   anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, '0% Loss', False)) 
            ICMPSummaryTreeview.heading('#8', text='25% Loss',   anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, '25% Loss', False)) 
            ICMPSummaryTreeview.heading('#9', text='50% Loss',   anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, '50% Loss', False)) 
            ICMPSummaryTreeview.heading('#10', text='75% Loss', anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, '75% Loss', False)) 
            ICMPSummaryTreeview.heading('#11', text='100% Loss', anchor=CENTER,command=lambda: treeview_sort_column(ICMPSummaryTreeview, '100% Loss', False)) 
            
            ICMPSummaryTreeview.column('#0', stretch=1, width=3 , anchor=W)
            ICMPSummaryTreeview.column('#1', stretch=1, width=10, anchor=W)
            ICMPSummaryTreeview.column('#2', stretch=1, width=10, anchor=W)
            ICMPSummaryTreeview.column('#3', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#4', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#5', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#6', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#7', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#8', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#9', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#10', stretch=1, width=10, anchor=CENTER)
            ICMPSummaryTreeview.column('#11', stretch=1, width=10, anchor=CENTER)

            # add tree and scrollbars to frame
            ICMPSummaryTreeview.grid(row=0, column=0, sticky=NSEW)
            ICMPSummaryTreeviewysb.grid(row=0, column=1, sticky=NS)
            ICMPSummaryTreeviewxsb.grid(row=1, column=0, sticky=EW)


            # create fonts and tags
            ICMPSummaryTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
           
            # Bind the double Click
            #ICMPSummaryTreeview.bind('<ButtonRelease-1>', on_ICMPSummary_Tree_select_click) # When Select the Tree
            ICMPSummaryTreeview.bind("<Double-1>", On_ICMPSummary_Tree_Double_Click)    

            # SQL Querry to the COUNTRY Table
            sql = """
                        SELECT * FROM ICMP_SUMMARY
                        ORDER BY Year desc, Month desc, Day desc, Hour desc, Minute desc, Second desc """
            if (db.Execute(sql)):
                i = 0
                while (i < len(db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [db.results[i][1],
                            db.results[i][2],
                            db.results[i][9],
                            db.results[i][10],
                            db.results[i][11],
                            db.results[i][12],
                            db.results[i][14],
                            db.results[i][15],
                            db.results[i][16],
                            db.results[i][17],
                            db.results[i][18]]
                    ICMPSummaryTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=ICMPSummaryFrame,title='ICMP Summary Report',
                message = 'No Records found')
            #------------------ TREE VIEW For ICMP Summary Database <END> -----------------------------------
            #db.Disconnect()
            ICMPSummaryWindow.mainloop()
        #else:
            #db.Disconnect()
    else:
        mbox.showerror(master=root,title='ICMP Summary Report',
                         message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                   + 'Please make sure the ODBC DSN Name mathes: ['
                                   + ODBC_DSN_name + "]")

def on_TracerouteSummaryWindow_quit():
    global TracerouteSummaryWindowExist
    global TracerouteSummaryWindow
    
    TracerouteSummaryWindowExist = False
    db.Disconnect()
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics Traceroute Summary Window','CLOSE Window']    
        Logging.Log(Parameter)
    TracerouteSummaryWindow.destroy()

def Traceroute_Call_Save_As():
    global TracerouteSummaryTreeviewDataColumns
    global TracerouteSummaryArray
    global TracerouteSummaryFrame

    data_ready = True
    if (data_ready):
        #print ("Save as")
        input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
        if input_file_name != "":
            file_name = input_file_name
            #version = '4.0'
            ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,file_name,version)
            Tab0 = "Traceroute Summary Report"
            Tabs = [Tab0]
            ExcellFile.Call_Write_to_File(Tabs)
            Row = 2    # 3
            Column = 1 # A
            ExcellFile.Add_DataToWorksheet(TracerouteSummaryTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
            Row = 3
            i = 0
            while (i < len(TracerouteSummaryArray)):                                                                            #<------ Data for the Columns
                item = TracerouteSummaryArray[i]
                ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                i = i + 1
                Row = Row + 1
            if (ExcellFile.Save_File()):
                mbox.showinfo(master=TracerouteSummaryFrame,title='Traceroute Summary Report in ms.',
                        message = '!!! The File was saved !!!')
                TracerouteSummaryWindow.title("Traceroute Summary Report in ms.                    File: ["+file_name+"] SAVED")                  #<---- Window Name Change
                if Is_logging_Available:
                    Parameter = []
                    Parameter = ['Vendor','SAVE AS']    
                    Logging.Log(Parameter)
            else:
                mbox.showerror(master=TracerouteSummaryFrame,title='Traceroute Summary Report in ms.',
                        message = '*** The File was not saved, Perhaps It is already open ***')



def Display_Traceroute_Summary_Window():
    global TracerouteSummaryWindowExist
    global TracerouteSummaryWindow
    global TracerouteSummaryLabel
    global TracerouteSummaryTreeview
    global TracerouteSummaryFrameEntry
    global TracerouteSummaryTreeviewDataColumns
    global TracerouteSummaryArray
    global TracerouteSummaryFrame
    
    if db.Connect():
        if not TracerouteSummaryWindowExist:
            # Set up the Window
            TracerouteSummaryWindowExist = True
            TracerouteSummaryWindow = Tk()
            TracerouteSummaryWindow.geometry('1350x700+350+70')
            TracerouteSummaryWindow.title("Traceroute Summary Report in ms.")
            TracerouteSummaryWindow.protocol("WM_DELETE_WINDOW", on_TracerouteSummaryWindow_quit)
            TracerouteSummaryWindow.call('tk', 'scaling', Windows_Scaling)
            TracerouteSummaryLabel = Label(TracerouteSummaryWindow,text="Helvetica", font=("Helvetica", 13))
            TracerouteSummaryLabel["text"] = "Traceroute Summary Report in ms."
            TracerouteSummaryLabel.pack()

            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(TracerouteSummaryWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=Traceroute_Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=on_TracerouteSummaryWindow_quit)
            TracerouteSummaryWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
            
            # Setup Frame
            TracerouteSummaryFrame = Frame(TracerouteSummaryWindow)
            TracerouteSummaryFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            TracerouteSummaryFrame.rowconfigure(0, weight=1)
            TracerouteSummaryFrame.columnconfigure(0, weight=1)

            #------------------ TREE VIEW For Traceroute Summary Database <BEGIN> -----------------------------------
            # Create Tree and Scrollbars
            TracerouteSummaryTreeviewDataColumns = ('Date','Time', 'Total No. Of Traceroutes','Min. Respopnse Time','Max. Response Time',
                                                    'Lower Avg. Response Time','Higer Avg. Response Time')
            
            TracerouteSummaryTreeview = ttk.Treeview(TracerouteSummaryFrame,columns=TracerouteSummaryTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
            TracerouteSummaryTreeviewysb = Scrollbar(TracerouteSummaryFrame,orient=VERTICAL, command=TracerouteSummaryTreeview.yview)           # <--- Make sure the frame is correct !!
            TracerouteSummaryTreeviewxsb = Scrollbar(TracerouteSummaryFrame,orient=HORIZONTAL, command=TracerouteSummaryTreeview.xview)         # <--- Make sure the frame is correct !!
            TracerouteSummaryTreeview['yscroll'] = TracerouteSummaryTreeviewysb.set
            TracerouteSummaryTreeview['xscroll'] = TracerouteSummaryTreeviewxsb.set

            # setup headings and column 
            TracerouteSummaryTreeview.heading('#0', text='Item No.',          anchor=W) # E for East and W for West
            TracerouteSummaryTreeview.heading('#1', text='Date',              anchor=W,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Date', False)) # E for East and W for West
            TracerouteSummaryTreeview.heading('#2', text='Time',              anchor=W,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Time', False)) 
            TracerouteSummaryTreeview.heading('#3', text='Total No. Of Traceroutes',  anchor=CENTER,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Total No. Of Traceroutes', False)) 
            TracerouteSummaryTreeview.heading('#4', text='Min. Respopnse Time',       anchor=CENTER,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Min. Respopnse Time', False)) 
            TracerouteSummaryTreeview.heading('#5', text='Max. Response Time',        anchor=CENTER,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Max. Response Time', False)) 
            TracerouteSummaryTreeview.heading('#6', text='Lower Avg. Response Time',  anchor=CENTER,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Lower Avg. Response Time', False)) 
            TracerouteSummaryTreeview.heading('#7', text='Higer Avg. Response Time',  anchor=CENTER,command=lambda: treeview_sort_column(TracerouteSummaryTreeview, 'Higer Avg. Response Time', False)) 
            
            TracerouteSummaryTreeview.column('#0', stretch=1, width=3 , anchor=W)
            TracerouteSummaryTreeview.column('#1', stretch=1, width=10, anchor=W)
            TracerouteSummaryTreeview.column('#2', stretch=1, width=10, anchor=W)
            TracerouteSummaryTreeview.column('#3', stretch=1, width=10, anchor=CENTER)
            TracerouteSummaryTreeview.column('#4', stretch=1, width=10, anchor=CENTER)
            TracerouteSummaryTreeview.column('#5', stretch=1, width=10, anchor=CENTER)
            TracerouteSummaryTreeview.column('#6', stretch=1, width=10, anchor=CENTER)
            TracerouteSummaryTreeview.column('#7', stretch=1, width=10, anchor=CENTER)

            # add tree and scrollbars to frame
            TracerouteSummaryTreeview.grid(row=0, column=0, sticky=NSEW)
            TracerouteSummaryTreeviewysb.grid(row=0, column=1, sticky=NS)
            TracerouteSummaryTreeviewxsb.grid(row=1, column=0, sticky=EW)


            # create fonts and tags
            TracerouteSummaryTreeview.tag_configure('Normal', font=('Helvetica', 8), background='gray1')
            
            # Bind the double Click
            #TracerouteSummaryTreeview.bind('<ButtonRelease-1>', on_TracerouteSummary_Tree_select_click) # When Select the Tree
            #TracerouteSummaryTreeview.bind("<Double-1>", On_TracerouteSummary_Tree_Edit)         

            # SQL Querry to the COUNTRY Table
            sql = """
                        SELECT * FROM TRACEROUTE_SUMMARY
                        ORDER BY Year desc, Month desc, Day desc, Hour desc, Minute desc, Second desc """
            if (db.Execute(sql)):
                i = 0
                TracerouteSummaryArray = []
                while (i < len(db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [db.results[i][1],
                            db.results[i][2],
                            db.results[i][3],
                            db.results[i][11],
                            db.results[i][10],
                            db.results[i][12],
                            db.results[i][13]]                           
                    TracerouteSummaryTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                    TracerouteSummaryArray.append(item)
            else:
                mbox.showerror(master=TracerouteSummaryFrame,title='Traceroute Summary Report',
                message = 'No Records found')
            #------------------ TREE VIEW For COUNTRY Database <END> -----------------------------------
            #db.Disconnect()
            TracerouteSummaryWindow.mainloop()
        #else:
            #db.Disconnect()
    else:
        mbox.showerror(master=root,title='ICMP Summary Report',
                         message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                   + 'Please make sure the ODBC DSN Name mathes: ['
                                   + ODBC_DSN_name + "]")

def OnSelect_ICMP(event): # no working :(
    #widget = event.widget
    widget = event
    print (event)
    print (widget)
    selection = widget.curselection()
    SelectionValue = selection[0]
    valuestr = widget.get(selection[0])
    print (SelectionValue)
    print (valuestr)

def OnDouble_Click_ICMP(event):
    #widget = event.widget
    
    selection = ICMPListBox.curselection()
    SelectionValue = selection[0]
    valuestr = ICMPListBox.get(selection[0])
    print (SelectionValue)
    print (valuestr)

def OnSelect(listboxname):
    global TablesListBox
    global ICMPListBox

    print (listboxname)
    if (listboxname == 'database'):
        selection = TablesListBox.curselection()
        SelectionValue = selection[0]
        valuestr = TablesListBox.get(selection[0])
        print (selection)
        print (SelectionValue)
        print (valuestr)

    if (listboxname == 'icmp'):
        selection = ICMPListBox.curselection()
        SelectionValue = selection[0]
        valuestr = ICMPListBox.get(selection[0])
        #print (SelectionStr)
        print (selection)
        print (SelectionValue)
        print (valuestr)
   
def Display_Window(event):
    selection = TablesListBox.curselection()
    SelectionValue = selection[0]
    valuestr = TablesListBox.get(selection[0])
    print (SelectionValue)
    print (valuestr)

def on_ICMP_Tree_double_click(event):   # <- goood for testig new functions !!!!!
        Display_Traceroute_Window()
        global ICMPtreeview

        curItem = ICMPtreeview.focus()
        dic = ICMPtreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            print (values)
            print ("Date => " + values[0])
            print ("Time => " + values[1])
            print ("IP Address => " + values[2])
            
            '''
            ipaddress = values[7]
            # SQL Querry to the ICMP Table
            if db.Connect():
                sql = """SELECT * FROM ICMP 
                        WHERE Device_IP = '%s' ORDER BY Year, Month, Day,
                        Hour, Minute, Second, Size_of_Ping DESC """ % (ipaddress)
                if (db.Execute(sql)):
                    #------------------------------- Deleting Begin --------
                    x = ICMPtreeview.get_children()
                    if x != '()': # checks if there is something in the first row
                        for child in x:
                            #print (child)
                            ICMPtreeview.delete(child)
                    #------------------------------- Deleting End --------
                    i = 0
                    while (i < len(db.results)):
                        num = i + 1
                        tags = "regions" # more work on in for RTTD > 250 ms.
                        item = [db.results[i][2],
                                db.results[i][3],
                                db.results[i][1],
                                db.results[i][10],
                                db.results[i][11],
                                db.results[i][13],
                                db.results[i][14],
                                db.results[i][12]]
                        ICMPtreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                        i = i + 1
                else:
                    mbox.showerror(title='ICMP Response Time History',
                    message = 'No Records found')
                db.Disconnect()
            else:
                mbox.showerror(title='ICMP Response Time History',
                message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                        + 'Please make sure the ODBC DSN Name mathes: ['
                        + ODBC_DSN_name + "]")
            '''

def Build_fake_data():
    # create a dict with a number as key, and randomized contents matching
    # the column layout of the table

    data = {}

    for n in range(20):
        num = n + 1
        ipaddress = '192.168.164.2%.2d' % num
        if random.randrange(2) == 0:
            deviceName = "Name_"+str(n)
            deviceDescription = "Description_"+str(n)
            country = "USA"
            region = 'Central'
            facility = 'Overlan Park'
            site = 'P Building'
            deviceType = 'Router'
            status = 'Disconnected'
        else:
            deviceName = "Name_"+str(n)
            deviceDescription = "Description_"+str(n)
            country = "UK"
            region = 'EMEA'
            facility = 'Red Hill'
            site = 'Suite 300'
            deviceType = 'Switched'
            status = 'Connected'
            
        data[num] = [deviceName,
                     deviceDescription,
                     country,
                     region,
                     facility,
                     site,
                     deviceType,
                     ipaddress,
                     status]
    return data

def Populate_tree_wiht_Fake_Data(tree,data):
    for n in range(len(data)):
        num = n+1
        item = data[num]

        if item[8] == 'Disconnected': # use highlight if status is 'Disconnected'
            tags = ('Disconnected')
        else:
            tags = ('Connected')

        #if '5' in item[0]: # override styles if there's a 5 in the ipaddress
        #    tags = ['whacky']

        tree.insert('', END, text='%3d'%num, values=item, tags=tags)


def Display_ICMP_Window():
    if Is_logging_Available:
        Parameter = []
        Parameter = ['BVAnalytics ICMP and Traceroute Window','OPEN Window']    
        Logging.Log(Parameter)
    ICMP.Display_ICMP_Window()

#########################################################    
#---------------- Response Times Menus <END> -----------#
#########################################################


    
# ============== Menu <BEGIN>======
def Window_Menu():
    menubar = Menu(root)
    
    file_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    #file_menu.add_command(label="New", command=donothing)
    #file_menu.add_command(label="Open", command=donothing)
    #file_menu.add_command(label="Save", command=donothing)
    #file_menu.add_command(label="Save As", command=donothing)
    #file_menu.add_separator()
    file_menu.add_command(label="Exit", command=CloseMainWindow)

    if (Is_ODBC_Available == True):
        find_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Find", menu=find_menu)
        find_menu.add_command(label="Find Circuits", command=Display_Find_Circuits_Window)
        find_menu.add_command(label="Find Equipment", command=Display_Find_Devices_Window)
        find_menu.add_command(label="Find Facility", command=Display_Find_Facility_Window)
        find_menu.add_command(label="Find Local Contacts", command=Display_Find_LocalPointOfContacts_Window)
        find_menu.add_command(label="Find Networks", command=Display_Find_Networks_Window)        
        #circuitmenu.add_separator()

        monitoring_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Response Times", menu=monitoring_menu)
        monitoring_menu.add_command(label="ICMP and Traceroute", command=Display_ICMP_Window)
        monitoring_menu.add_command(label="ICMP Summary", command=Display_ICMP_Summary_Window)
        monitoring_menu.add_command(label="Traceroute Summary", command=Display_Traceroute_Summary_Window)
        #circuitmenu.add_separator()


        location_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Locations", menu=location_menu)
        if (Is_Country_Available):
            location_menu.add_command(label="Country", command=Display_Country_Window)
        if (Is_Region_Available):
            location_menu.add_command(label="Region", command=Display_Region_Window)
        if (Is_Facility_Available):
            location_menu.add_command(label="Facility / Office", command=Display_Facility_Window)
        if (Is_Sites_Available):
            location_menu.add_command(label="Site", command=Display_Sites_Window)
        #if (Is_LocalPOC_Available):
            location_menu.add_command(label="Local Point Of Contacts", command=Display_LocalPointOfContacts_Window)

        #circuitmenu.add_separator()


        equipment_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Equipment", menu=equipment_menu)
        if (Is_Device_Available):
            equipment_menu.add_command(label="Device", command=Display_Device_Window)
        if (Is_Vendor_Available):
            equipment_menu.add_command(label="Vendors", command=Display_Vendor_Window)
        equipment_menu.add_command(label="Device Type", command=Display_DeviceType_Window)
        equipment_menu.add_command(label="Device Model", command=Display_DeviceModel_Window)
        equipment_menu.add_separator()
        equipment_menu.add_command(label="Report Outsource Cost Per Country", command=Report_Outsource_Per_Country)
        equipment_menu.add_command(label="Report Outsource Cost Per Facility", command=Report_Outsource_Per_Facility)
        equipment_menu.add_separator()
        if (Is_ImportCMDB_Available):
            equipment_menu.add_command(label="Import Network Inventory", command=Display_ImportCMDB_Window)


        circuits_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Circuits", menu=circuits_menu)
        if (Is_Circuit_Available):
            circuits_menu.add_command(label="Circuit", command=Display_Circuit_Window)
        if (Is_Carrier_Available):
            circuits_menu.add_command(label="Carriers", command=Display_Carrier_Window)
        if (Is_CircuitType_Available):
            circuits_menu.add_command(label="Circuit Types", command=Display_CircuitType_Window)
        if (Is_PortSpeed_Available):
            circuits_menu.add_command(label="Port Speeds", command=Display_PortSpeed_Window)

        network_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Network", menu=network_menu)
        if (Is_NetworkTemplates_Available):
            network_menu.add_command(label="Network Templates", command=Display_NetworkTemplates_Window)
        if (Is_NetworkTemplatesBreakOut_Available):
            network_menu.add_command(label="Network Templates Provisioning", command=Display_NetworkTemplatesBreakOut_Window)
        network_menu.add_separator()
        if (Is_Region_Available):
            network_menu.add_command(label="Network Regions", command=Display_Region_Window)
        if (Is_NetworkAllocations_Available):
            network_menu.add_command(label="Network Allocations", command=Display_NetworkAllocations_Window)
        if (Is_NetworkProvisioning_Available):
            network_menu.add_command(label="Network Provisioning", command=Display_NetworkProvisioning_Window)


        odbc_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Database Tools", menu=odbc_menu)
        odbc_menu.add_command(label="Test ODBC Connection", command=Test_ODBC_Connection)
        odbc_menu.add_command(label="Get Database Version", command=Database_Version)
        odbc_menu.add_separator()
        odbc_menu.add_command(label="Update Tables and Fields", command=Update_Tables_and_Fields)
        #odbcmenu.add_separator()

        log_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Logging", menu=log_menu)
        log_menu.add_command(label="User Activity", command=Find_Logging)
        #log_menu.add_command(label="...", command=donothing)
        #log_menu.add_separator()
        #log_menu.add_command(label="...s", command=Update_Tables_and_Fields)
        #odbcmenu.add_separator()

    help_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Naming Conventions", command=donothing)
    help_menu.add_command(label="IP Address Schema", command=donothing)
    help_menu.add_command(label="Offices Templates/Standards", command=donothing)
    help_menu.add_command(label="VLANs Conventions", command=donothing)
    help_menu.add_command(label="About", command=Window_About)

    root.config(menu=menubar)
# ============== Menu <END>======


def Main():
    if sys.argv[1:]:
#        print ("there is parameters")
        Parameters = int(sys.argv[1])
#        print ("==== this is what you typed => ",sys.argv[0:])
    else:
        Parameters = 10
#        print ("There are NO parameters")

#    print ('The Parameters is => ', Parameters )
#    WelcomneScreen()




# =========================>>>>>>>> MAIN PROGRAM START HERE <<<<<<===================
Main()

# ------ Gloval Variables ---------
root = Tk()
# root.iconbitmap(default='C:\Windows\SysWOW64\bvlogo.ico') <- Need to work on this
#root.geometry('1325x660+350+70')
root.geometry('800x125+350+70')
#root.title("Welcome to the Black & Veatch Analytics Tool     [ {} - {} ] Database in use: [ {} ]".format(os.path.basename(file_name),file_saved_status,ODBC_DSN_name))
root.title("Welcome to the Black & Veatch Analytics Tool     [ {} - {} ] Database in use: [ {} ]".format(BackendServer,Enviroment,ODBC_DSN_name))
root.protocol("WM_DELETE_WINDOW", CloseMainWindow)
root.call('tk', 'scaling', Windows_Scaling)

root_frame = Frame(root)
root_frame.pack()

rootLabel1 = Label(root_frame,text="Helvetica", font=("Helvetica", 15))
rootLabel1["text"] = "______________________________"
rootLabel1.pack()

rootLabel2 = Label(root_frame,text="Helvetica", font=("Helvetica", 15))
rootLabel2["text"] = "Black & Veatch Analytics Tool"
rootLabel2.pack()

rootLabel3 = Label(root_frame,text="Helvetica", font=("Helvetica", 15))
rootLabel3["text"] = "______________________________"
rootLabel3.pack()


#----------------- Research on how to add Scroll Bars to the Main Window
'''
MListBox = Listbox(root_frame, width=120, height=30, font=("Helvetica", 8))
yscrollbar = Scrollbar(root_frame, orient="vertical")
yscrollbar.config(command=MListBox.yview)
yscrollbar.pack(side="right", fill="y")
MListBox.config(yscrollcommand=yscrollbar.set)        
xscrollbar = Scrollbar(root_frame, orient="horizontal")
xscrollbar.config(command=MListBox.xview)
xscrollbar.pack(side="bottom", fill="y")
MListBox.config(xscrollcommand=xscrollbar.set)
MListBox.pack(side="left", fill="y")
'''
#---------------- Display the Menu -------------
Window_Menu()
root.mainloop()


