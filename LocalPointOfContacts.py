#==========================================================================
# Program: LocalPointOfContacts.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Feb-25-2018
# Date Last Modified: Feb-25-2018
# Summary: This is Class to for the LocalPointOfContacts
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
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Local Point Of Contatcs','OPEN Window']    
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

class Class_LocalPointOfContacts:

    def __init__(self,DSN_Name,Windows_Scaling,Location):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.LocalPointOfContactsWindowExist = False
        self.LocalPointOfContactsCalendarExist = False
        self.Username = os.getlogin()
        self.DateAndTime = datetime.datetime.now()                 
        self.date = ""
        self.Windows_Scaling = Windows_Scaling
        self.Selection = 'none'
        self.LocalPointOfContactsTableSeqNo = 0
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

        
#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************
        
    def Clean_Screen(self,option,option2):
        # Setup Buttons

        #self.LocalPointOfContactsBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False

        if (option == 'country'): ## The option are country,region and LocalPointOfContacts
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

        if (option == 'region'):
            self.ComboBoxFacilityID.set("")
            self.ComboBoxFacilityID['state'] = DISABLED
            if (option2 != 'region-combo'):
                self.ComboBoxRegionID.set("")
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = DISABLED
                self.ButtonFacilityRefresh['state'] = DISABLED

        if (option2 == 'facility-combo'):
            self.ButtonLocalPointOfContactsAdd['state'] = ACTIVE
        else:
            self.ButtonLocalPointOfContactsAdd['state'] = DISABLED
            
        self.ButtonLocalPointOfContactsEdit['state'] = DISABLED
        self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
        self.ButtonLocalPointOfContactsOK['state'] = DISABLED
        self.ButtonLocalPointOfContactsCancel['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry        
        self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
        self.LocalPointOfContacts1FrameEntry.delete(0,END)
        self.LocalPointOfContacts1FrameEntry['state'] = 'readonly'

        self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
        self.LocalPointOfContactsOfficeNo1FrameEntry.delete(0,END)
        self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'readonly'

        self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'normal'
        self.LocalPointOfContactsAlternateNo1FrameEntry.delete(0,END)
        self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'readonly'

        self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'normal'
        self.LocalPointOfContacts2ndAlternateNo1FrameEntry.delete(0,END)
        self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'readonly'

        self.LocalPointOfContacts2FrameEntry['state'] = 'normal'
        self.LocalPointOfContacts2FrameEntry.delete(0,END)
        self.LocalPointOfContacts2FrameEntry['state'] = 'readonly'
                    
        self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'normal'
        self.LocalPointOfContactsOfficeNo2FrameEntry.delete(0,END)
        self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'readonly'

        self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'normal'
        self.LocalPointOfContactsAlternateNo2FrameEntry.delete(0,END)
        self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'readonly'

        self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'normal'
        self.LocalPointOfContacts2ndAlternateNo2FrameEntry.delete(0,END)
        self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'readonly'

        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.LocalPointOfContactsTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.LocalPointOfContactsTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------

    def Display_Screen(self,curItem):      
            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.LocalPointOfContactsPrimaryKey = self.LocalPointOfContactsTablePrimaryKeyArray[curItem]
            self.LocalPointOfContactsTableSeqNo = self.LocalPointOfContactsTableSeqNoArray[curItem]
            
            self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts1FrameEntry.delete(0,END)
            self.LocalPointOfContacts1FrameEntry.insert(0,self.LocalPointOfContactsTablePrimaryPOCNameArray[curItem])
            self.LocalPointOfContacts1FrameEntry['state'] = 'readonly'

            self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo1FrameEntry.delete(0,END)
            self.LocalPointOfContactsOfficeNo1FrameEntry.insert(0,self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray[curItem])
            self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'readonly'

            self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo1FrameEntry.delete(0,END)
            self.LocalPointOfContactsAlternateNo1FrameEntry.insert(0,self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray[curItem])
            self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'readonly'

            self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo1FrameEntry.delete(0,END)
            self.LocalPointOfContacts2ndAlternateNo1FrameEntry.insert(0,self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray[curItem])
            self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'readonly'

            self.LocalPointOfContacts2FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2FrameEntry.delete(0,END)
            self.LocalPointOfContacts2FrameEntry.insert(0,self.LocalPointOfContactsTableSecondaryPOCNameArray[curItem])
            self.LocalPointOfContacts2FrameEntry['state'] = 'readonly'
                        
            self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo2FrameEntry.delete(0,END)
            self.LocalPointOfContactsOfficeNo2FrameEntry.insert(0,self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray[curItem])
            self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'readonly'

            self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo2FrameEntry.delete(0,END)
            self.LocalPointOfContactsAlternateNo2FrameEntry.insert(0,self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray[curItem])
            self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'readonly'

            self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo2FrameEntry.delete(0,END)
            self.LocalPointOfContacts2ndAlternateNo2FrameEntry.insert(0,self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray[curItem])
            self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'readonly'



    def Enable_Screen(self,option):
        # This function is used when the ADD button is selected

        #self.LocalPointOfContactsBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

        self.ComboBoxCoutryID['state'] = DISABLED
        self.ComboBoxRegionID['state'] = DISABLED
        self.ComboBoxFacilityID['state'] = DISABLED
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = DISABLED
            self.ButtonCountryRefresh['state'] = DISABLED
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = DISABLED
            self.ButtonRegionRefresh['state'] = DISABLED
        if (Is_Facility_Available):
            self.ButtonFacilityAdd['state'] = DISABLED
            self.ButtonFacilityRefresh['state'] = DISABLED
            
        self.ButtonLocalPointOfContactsAdd['state'] = DISABLED
        self.ButtonLocalPointOfContactsEdit['state'] = DISABLED
        self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
        self.ButtonLocalPointOfContactsOK['state'] = ACTIVE
        self.ButtonLocalPointOfContactsCancel['state'] = ACTIVE

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        if (option == 'add'): #<----------------------------------- ADD Button
            self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts1FrameEntry.delete(0,END)

            self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo1FrameEntry.delete(0,END)

            self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo1FrameEntry.delete(0,END)

            self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo1FrameEntry.delete(0,END)

            self.LocalPointOfContacts2FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2FrameEntry.delete(0,END)
                        
            self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo2FrameEntry.delete(0,END)

            self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo2FrameEntry.delete(0,END)

            self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo2FrameEntry.delete(0,END)

        if (option == 'edit'): #<----------------------------------- EDIT Button
            self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'normal'
            self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'normal'
                
       
    def Disable_Screen(self):
        # This function is used when the entry was added.modified to the Database

        #self.LocalPointOfContactsBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

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
            
        self.ButtonLocalPointOfContactsAdd['state'] = ACTIVE
        self.ButtonLocalPointOfContactsEdit['state'] = DISABLED
        self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
        self.ButtonLocalPointOfContactsOK['state'] = DISABLED
        self.ButtonLocalPointOfContactsCancel['state'] = ACTIVE

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.LocalPointOfContacts1FrameEntry['state'] = 'readonly'
        self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'readonly'
        self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'readonly'
        self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'readonly'
        self.LocalPointOfContacts2FrameEntry['state'] = 'readonly'
        self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'readonly'
        self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'readonly'
        self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'readonly'


    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.LocalPointOfContactsTablePrimaryPOCName = self.LocalPointOfContacts1FrameEntry.get()
        self.LocalPointOfContactsTablePrimaryPOCOfficeNo = self.LocalPointOfContactsOfficeNo1FrameEntry.get()
        self.LocalPointOfContactsTablePrimaryPOCAlternateNo = self.LocalPointOfContactsAlternateNo1FrameEntry.get()
        self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo = self.LocalPointOfContacts2ndAlternateNo1FrameEntry.get()
        self.LocalPointOfContactsTableSecondaryPOCName = self.LocalPointOfContacts2FrameEntry.get()
        self.LocalPointOfContactsTableSecondaryPOCOfficeNo = self.LocalPointOfContactsOfficeNo2FrameEntry.get()
        self.LocalPointOfContactsTableSecondaryPOCAlternateNo = self.LocalPointOfContactsAlternateNo2FrameEntry.get()
        self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo = self.LocalPointOfContacts2ndAlternateNo2FrameEntry.get()

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <END> ------------------------*
#****************************************************************************************


#****************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************

    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # SQL Querry to the LocalPointOfContacts Table
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
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='LocalPointOfContacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the LocalPointOfContacts Table
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
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                message = 'No Country Records found')
                self.sql_querry = False
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
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
            Parameter = ['Local Point Of Contacts','Region Window']    
            Logging.Log(Parameter)
        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        self.Clean_Screen('region','region-combo')
        if self.db.Connect():
            # SQL Querry to the LocalPointOfContacts Table
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
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','Region Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the LocalPointOfContacts Table
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
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                message = 'No Records found')
                self.ComboBoxRegionID['state'] = DISABLED
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonLocalPointOfContactsAdd['state']    = DISABLED
                self.ButtonLocalPointOfContactsEdit['state']   = DISABLED
                self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
                self.ButtonLocalPointOfContactsOK['state']     = DISABLED
                self.ButtonLocalPointOfContactsCancel['state'] = DISABLED

                #self.LocalPointOfContactsIDFrameEntry['state'] = 'normal'
                #self.LocalPointOfContactsIDFrameEntry.delete(0,END)
                #self.LocalPointOfContactsIDFrameEntry['state'] = 'readonly'
                
                #self.LocalPointOfContactsNameFrameEntry['state'] = 'normal'
                #self.LocalPointOfContactsNameFrameEntry.delete(0,END)
                #self.LocalPointOfContactsNameFrameEntry['state'] = 'readonly'

                self.sql_querry = False

        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
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
            Parameter = ['Local Point Of Contacts','Facility Window']    
            Logging.Log(Parameter)
        Facility.Display_Facility_Window()

    def on_facility_combo_changed(self,event):
        self.Clean_Screen('facility','facility-combo')
        if self.db.Connect():
            # SQL Querry to the LocalPointOfContacts Table
            sql = """
                    SELECT * FROM Local_Point_Of_Contacts
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                    ORDER BY Primary_POC_Name ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                         self.FacilityIDArray[self.ComboBoxFacilityID.current()])
            #print (sql)
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.LocalPointOfContactsTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.LocalPointOfContactsTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                #-------------- Initializing Arrays <BEGIN> ----------------------
                self.LocalPointOfContactsTablePrimaryKeyArray =  []
                self.LocalPointOfContactsTableCountryIDArray =  [] 
                self.LocalPointOfContactsTableRegionIDArray =  []
                self.LocalPointOfContactsTableFacilityIDArray =  []
                self.LocalPointOfContactsTableSeqNoArray =  []
                self.LocalPointOfContactsTablePrimaryPOCNameArray =  []
                self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray =  []
                self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray =  []
                self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray =  []
                self.LocalPointOfContactsTableSecondaryPOCNameArray =  []
                self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray =  []
                self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray =  []
                self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray =  []
                self.LocalPointOfContactsTableLastTimeItWasUpdatedArray = []
                self.LocalPointOfContactsTableExecutedByUserArray = []
                self.LocalPointOfContactsTableArrayColumns = []
                self.LocalPointOfContactsTableArrayTemp = []
                self.results = []

                self.LocalPointOfContactsTableArrayColumns = (
                                'Country ID',
                                'Region ID',
                                'Facility ID',
                                'SEQ No',
                                'Primary POC Name',
                                'Primary POC Office No',
                                'Primary POC Alternate No',
                                'Primary POC 2nd Alternate No',
                                'Secondary POC Name',
                                'Secondary POC Office No',
                                'Secondary POC Alternate No',
                                'Secondary POC 2nd Alternate No',
                                'Last Time it was Updated',
                                'Executed by UserID'
                                )
                
                #-------------- Initializing Arrays <END> ----------------------
                self.data_ready = True
                i = 0
                while (i < len(self.db.results)):
                    self.LocalPointOfContactsTablePrimaryKeyArray.append(self.db.results[i][0].strip())
                    self.LocalPointOfContactsTableCountryIDArray.append(self.db.results[i][1].strip())
                    self.LocalPointOfContactsTableRegionIDArray.append(self.db.results[i][2].strip())
                    self.LocalPointOfContactsTableFacilityIDArray.append(self.db.results[i][3].strip())
                    self.LocalPointOfContactsTableSeqNoArray.append(self.db.results[i][4])
                    if (self.db.results[i][5] == None):
                        self.LocalPointOfContactsTablePrimaryPOCNameArray.append("")
                    else:
                        self.LocalPointOfContactsTablePrimaryPOCNameArray.append(self.db.results[i][5].strip())
                    if (self.db.results[i][6] == None):
                        self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray.append("")
                    else:
                        self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray.append(self.db.results[i][6].strip())
                    if (self.db.results[i][7] == None):
                        self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray.append("")
                    else: 
                        self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray.append(self.db.results[i][7].strip())
                    if (self.db.results[i][8] == None):
                        self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray.append("")
                    else:
                        self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray.append(self.db.results[i][8].strip())
                    if (self.db.results[i][9] == None):
                        self.LocalPointOfContactsTableSecondaryPOCNameArray.append("")
                    else:
                        self.LocalPointOfContactsTableSecondaryPOCNameArray.append(self.db.results[i][9].strip())
                    if (self.db.results[i][10] == None):
                        self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray.append("")
                    else:
                        self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray.append(self.db.results[i][10].strip())
                    if (self.db.results[i][11] == None):
                        self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray.append("")
                    else:
                        self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray.append(self.db.results[i][11].strip())
                    if (self.db.results[i][12] == None):
                        self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray.append("")
                    else:
                        self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray.append(self.db.results[i][12].strip())
                    self.LocalPointOfContactsTableLastTimeItWasUpdatedArray.append(self.db.results[i][13].strip())
                    self.LocalPointOfContactsTableExecutedByUserArray.append(self.db.results[i][14].strip())
                    i = i + 1
                i = 0
                while (i < len(self.LocalPointOfContactsTablePrimaryKeyArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            self.LocalPointOfContactsTablePrimaryPOCNameArray[i],
                            self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray[i],
                            self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray[i],
                            self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray[i],
                            self.LocalPointOfContactsTableSecondaryPOCNameArray[i],
                            self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray[i],
                            self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray[i],
                            self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray[i],
                            self.LocalPointOfContactsTableExecutedByUserArray[i],
                            self.LocalPointOfContactsTableLastTimeItWasUpdatedArray[i],
                            ]
                    self.LocalPointOfContactsTableArrayTemp = [
                        self.LocalPointOfContactsTableCountryIDArray[i],
                        self.LocalPointOfContactsTableRegionIDArray[i],
                        self.LocalPointOfContactsTableFacilityIDArray[i],
                        self.LocalPointOfContactsTableSeqNoArray[i],
                        self.LocalPointOfContactsTablePrimaryPOCNameArray[i],
                        self.LocalPointOfContactsTablePrimaryPOCOfficeNoArray[i],
                        self.LocalPointOfContactsTablePrimaryPOCAlternateNoArray[i],
                        self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNoArray[i],
                        self.LocalPointOfContactsTableSecondaryPOCNameArray[i],
                        self.LocalPointOfContactsTableSecondaryPOCOfficeNoArray[i],
                        self.LocalPointOfContactsTableSecondaryPOCAlternateNoArray[i],
                        self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNoArray[i],
                        self.LocalPointOfContactsTableLastTimeItWasUpdatedArray[i],
                        self.LocalPointOfContactsTableExecutedByUserArray[i]
                        ]
                    self.results.append(self.LocalPointOfContactsTableArrayTemp)
                    self.LocalPointOfContactsTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonLocalPointOfContactsAdd['state']    = ACTIVE
                self.ButtonLocalPointOfContactsEdit['state']   = DISABLED
                self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
                self.ButtonLocalPointOfContactsOK['state']     = DISABLED
                self.ButtonLocalPointOfContactsCancel['state'] = DISABLED
            else:
                if (self.Go_To_Location):
                    mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                    message = 'No Records found for Local Point of Contact, But Adding a BLANK Record to be filled later')
                    self.Selection = 'add'
                    self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
                    self.LocalPointOfContacts1FrameEntry.delete(0,END)
                    self.LocalPointOfContacts1FrameEntry.insert(0,"Missing POC")
                    self.LocalPointOfContacts1FrameEntry['state'] = 'readonly'

                    self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
                    self.LocalPointOfContactsOfficeNo1FrameEntry.delete(0,END)
                    self.LocalPointOfContactsOfficeNo1FrameEntry.insert(0,"Missing Phone")
                    self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'readonly'
                    self.Call_Button_LocalPointOfContacts_OK()
                else:
                    mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                    message = 'No Records found for Local Point of Contact')
                
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','Facility Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('facility','all')
        if self.db.Connect():
            # SQL Querry to the LocalPointOfContacts Table
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
                else:
                    self.ComboBoxFacilityID['state'] = 'readonly'
                    self.ComboBoxFacilityID.set("")
                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd['state'] = 'active'
                        self.ButtonFacilityRefresh['state'] = 'active'
            else:
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                message = 'No Records found')
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonLocalPointOfContactsAdd['state']    = DISABLED
                self.ButtonLocalPointOfContactsEdit['state']   = DISABLED
                self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
                self.ButtonLocalPointOfContactsOK['state']     = DISABLED
                self.ButtonLocalPointOfContactsCancel['state'] = DISABLED


                # Setup Labels and Entry        
                self.LocalPointOfContacts1FrameEntry['state'] = 'normal'
                self.LocalPointOfContacts1FrameEntry.delete(0,END)
                self.LocalPointOfContacts1FrameEntry['state'] = 'readonly'

                self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'normal'
                self.LocalPointOfContactsOfficeNo1FrameEntry.delete(0,END)
                self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = 'readonly'

                self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'normal'
                self.LocalPointOfContactsAlternateNo1FrameEntry.delete(0,END)
                self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = 'readonly'

                self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'normal'
                self.LocalPointOfContacts2ndAlternateNo1FrameEntry.delete(0,END)
                self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = 'readonly'

                self.LocalPointOfContacts2FrameEntry['state'] = 'normal'
                self.LocalPointOfContacts2FrameEntry.delete(0,END)
                self.LocalPointOfContacts2FrameEntry['state'] = 'readonly'
                            
                self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'normal'
                self.LocalPointOfContactsOfficeNo2FrameEntry.delete(0,END)
                self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = 'readonly'

                self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'normal'
                self.LocalPointOfContactsAlternateNo2FrameEntry.delete(0,END)
                self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = 'readonly'

                self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'normal'
                self.LocalPointOfContacts2ndAlternateNo2FrameEntry.delete(0,END)
                self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = 'readonly'

                self.sql_querry = False

        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <END> ------------------------*
#*************************************************************************************

    def Find_Next_Sequence_ID_on_LocalPointOfContacts(self):
        i = 0
        
    def on_LocalPointOfContactsWindow_quit(self):
        if (self.LocalPointOfContactsWindowExist):
            self.LocalPointOfContactsWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Local Point Of Contacts','CLOSE Winodw']    
                Logging.Log(Parameter)            
            self.LocalPointOfContactsWindow.destroy()


    def on_LocalPointOfContacts_Table_Refresh(self): # I need to do more research on this call.
        self.on_country_combo_changed("event")
       
    def Call_Button_LocalPointOfContacts_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','ADD Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.Enable_Screen('add')
        self.Selection = 'add'

    def Call_Button_LocalPointOfContacts_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        self.Enable_Screen('edit')

    def Call_Button_LocalPointOfContacts_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
            self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
            self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
            #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID+"-"+str(self.LocalPointOfContactsPrimaryKey))
            PrimaryKey = self.LocalPointOfContactsPrimaryKey     
            if (mbox.askyesnocancel(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',message = 'Are you Sure you want to Remove it?')):
                #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.LocalPointOfContactsID)
                #print (PrimaryKey)
                if (self.LocalPointOfContactsTableSeqNo == 0 ):
                    mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                        message = '*** The LocalPointOfContacts ZERO CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in LocalPointOfContacts --------
                    sql = """
                            SELECT * FROM DEVICES
                            WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                          """ % (self.CountryID,self.RegionID,self.FacilityID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                    message = '*** The Local Point Of Contacts you entered CANNOT be Removed Becuase there still are some Devices in that Facility***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                    else:
                        sql = """
                                SELECT * FROM Local_Point_Of_Contacts
                                WHERE Country_ID_Region_ID_Facility_ID_SEQ_No = '%s'
                              """ % (PrimaryKey)
                        #print (sql)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Local_Point_Of_Contacts WHERE Country_ID_Region_ID_Facility_ID_Seq_No = '%s'" % (PrimaryKey)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                        message = '*** The Local Point Of Contacts you entered were Removed ***')
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                        message = '*** The Local Point Of Contacts you entered were NOT Removed ***')               
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                            message = '*** The Local Point Of Contacts you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_LocalPointOfContacts_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ButtonLocalPointOfContactsAdd['state']    = DISABLED
            self.ButtonLocalPointOfContactsEdit['state']   = DISABLED
            self.ButtonLocalPointOfContactsRemove['state'] = DISABLED
            self.ButtonLocalPointOfContactsOK['state']     = ACTIVE
            self.ButtonLocalPointOfContactsCancel['state'] = ACTIVE
            self.Collect_Screen()
            PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID+"-"+str(self.LocalPointOfContactsTableSeqNo))
            #PrimaryKey = self.LocalPointOfContactsPrimaryKey
            #self.Selection = 'notyet'
            #-------------- ADD ----------------------
            '''
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
                                Executed_by_UserID                          CHAR(20))"""

            '''
            if (self.Selection == 'add'):
                if ((len(self.LocalPointOfContactsTablePrimaryPOCName) > 0) and (len(self.LocalPointOfContactsTablePrimaryPOCOfficeNo) > 0)):
                    sql = """
                            SELECT * FROM Local_Point_Of_Contacts
                            WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                            ORDER BY SEQ_No desc
                          """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                                 self.FacilityIDArray[self.ComboBoxFacilityID.current()])
                    if (self.db.Execute(sql)):
                        #self.LocalPointOfContactsTableSeqNo = len(self.db.results) + 1
                        #print (self.db.results[0][4])
                        #print (self.db.results[0][4] + 1)
                        self.LocalPointOfContactsTableSeqNo = self.db.results[0][4] + 1
                    else:
                        self.LocalPointOfContactsTableSeqNo = 1
                    PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID+"-"+str(self.LocalPointOfContactsTableSeqNo))
                    #print (PrimaryKey)
                    sql = """
                            SELECT * FROM Local_Point_Of_Contacts
                            WHERE Country_ID_Region_ID_Facility_ID_SEQ_No = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                        message = '*** The Local Point Of Contacts you entered already exist ***')
                    else:
                        if (len(self.LocalPointOfContactsTablePrimaryPOCName) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCName = self.LocalPointOfContactsTablePrimaryPOCName[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOCOfficeNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCOfficeNo = self.LocalPointOfContactsTablePrimaryPOCOfficeNo[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOCAlternateNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCAlternateNo = self.LocalPointOfContactsTablePrimaryPOCAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo = self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCName) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCName = self.LocalPointOfContactsTableSecondaryPOCName[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCOfficeNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCOfficeNo = self.LocalPointOfContactsTableSecondaryPOCOfficeNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCAlternateNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCAlternateNo = self.LocalPointOfContactsTableSecondaryPOCAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo = self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo[:100]
                        
                        sql = "INSERT INTO Local_Point_Of_Contacts(Country_ID_Region_ID_Facility_ID_SEQ_No, Country_ID, Region_ID, Facility_ID, SEQ_No, \
                               Primary_POC_Name, Primary_POC_Office_No, Primary_POC_Alternate_No, Primary_POC_2nd_Alternate_No, \
                               Secondary_POC_Name, Secondary_POC_Office_No, Secondary_POC_Alternate_No, Secondary_POC_2nd_Alternate_No, \
                               Last_Time_It_Was_Updated, Executed_by_UserID) \
                                   VALUES ('%s','%s','%s','%s','%d','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.CountryID,self.RegionID,
                                            self.FacilityID, self.LocalPointOfContactsTableSeqNo,
                                            self.LocalPointOfContactsTablePrimaryPOCName,self.LocalPointOfContactsTablePrimaryPOCOfficeNo,
                                            self.LocalPointOfContactsTablePrimaryPOCAlternateNo,self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo,                                                                        
                                            self.LocalPointOfContactsTableSecondaryPOCName,self.LocalPointOfContactsTableSecondaryPOCOfficeNo,
                                            self.LocalPointOfContactsTableSecondaryPOCAlternateNo,self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo,                                                                        
                                            self.DateAndTime,self.Username)
                        #print (sql)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                message = '*** The LocalPointOfContacts ID you entered was Added ***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                message = '*** The LocalPointOfContacts ID you entered was NOT Added ***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                else:
                     mbox.showerror(master=self.LocalPointOfContactsFrame,title='LocalPointOfContacts',
                        message = '*** The Local Point Of Contacts Primary Name and Office No. Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                #print (PrimaryKey+"....")
                if ((len(self.LocalPointOfContactsTablePrimaryPOCName) == 0) or (len(self.LocalPointOfContactsTablePrimaryPOCOfficeNo) ==  0)):
                    mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                    message = '*** The Local Point Of Contacts Primary Name and Office No. Cannot be BLANK ***')
                else:
                    sql = """
                            SELECT * FROM Local_Point_Of_Contacts
                            WHERE Country_ID_Region_ID_Facility_ID_SEQ_No = '%s'
                          """ % (PrimaryKey)
                    #print (sql)
                    if (self.db.Execute(sql)):   
                        #sql = "UPDATE LocalPointOfContacts SET LocalPointOfContacts_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID_LocalPointOfContacts_ID = '%s'" % (LocalPointOfContactsName,self.Username,PrimaryKey)
                        if (len(self.LocalPointOfContactsTablePrimaryPOCName) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCName = self.LocalPointOfContactsTablePrimaryPOCName[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOCOfficeNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCOfficeNo = self.LocalPointOfContactsTablePrimaryPOCOfficeNo[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOCAlternateNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOCAlternateNo = self.LocalPointOfContactsTablePrimaryPOCAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo) > 100):
                            self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo = self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCName) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCName = self.LocalPointOfContactsTableSecondaryPOCName[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCOfficeNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCOfficeNo = self.LocalPointOfContactsTableSecondaryPOCOfficeNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOCAlternateNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOCAlternateNo = self.LocalPointOfContactsTableSecondaryPOCAlternateNo[:100]
                        if (len(self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo) > 100):
                            self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo = self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo[:100]

                        sql = "UPDATE Local_Point_Of_Contacts SET Primary_POC_Name  = '%s', Primary_POC_Office_No = '%s', Primary_POC_Alternate_No = '%s', Primary_POC_2nd_Alternate_No = '%s', \
                               Secondary_POC_Name = '%s', Secondary_POC_Office_No = '%s', Secondary_POC_Alternate_No = '%s', Secondary_POC_2nd_Alternate_No = '%s', \
                               Last_Time_It_Was_Updated = '%s', Executed_by_UserID = '%s' \
                               WHERE Country_ID_Region_ID_Facility_ID_SEQ_No = '%s'" % (self.LocalPointOfContactsTablePrimaryPOCName,self.LocalPointOfContactsTablePrimaryPOCOfficeNo,
                                            self.LocalPointOfContactsTablePrimaryPOCAlternateNo,self.LocalPointOfContactsTablePrimaryPOC2ndAlternateNo,                                                                        
                                            self.LocalPointOfContactsTableSecondaryPOCName,self.LocalPointOfContactsTableSecondaryPOCOfficeNo,
                                            self.LocalPointOfContactsTableSecondaryPOCAlternateNo,self.LocalPointOfContactsTableSecondaryPOC2ndAlternateNo,                                                                        
                                            self.DateAndTime,self.Username,PrimaryKey)
                             
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                message = '*** The Local Point Of Contacts ID you entered were Updated ***')
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                message = '*** The Local Point Of Contacts ID you entered were NOT Upadted ***')
                        self.on_facility_combo_changed("event")
                        self.Disable_Screen()
                    else:
                        mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                        message = '*** The Local Point Of Contacts ID you try to Edit Does not exist Anymore ***')
                        self.on_facility_combo_changed("event")
                        self.Disable_Screen()                        
        else:
            mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_LocalPointOfContacts_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Local Point Of Contacts','CANCEL Button']    
            Logging.Log(Parameter)
        self.Clean_Screen('LocalPointOfContacts','all')
        self.on_facility_combo_changed("test")
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
        self.Selection = 'cancel'

    def on_LocalPointOfContacts_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.LocalPointOfContactsTreeview.focus()  
        dic = self.LocalPointOfContactsTreeview.item(curItem)
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
            curItem = int(dic.get('text')) - 1
            self.ButtonLocalPointOfContactsAdd['state']    = ACTIVE
            self.ButtonLocalPointOfContactsEdit['state']   = ACTIVE
            self.ButtonLocalPointOfContactsRemove['state'] = ACTIVE
            self.ButtonLocalPointOfContactsOK['state']     = DISABLED
            self.ButtonLocalPointOfContactsCancel['state'] = DISABLED
            self.Display_Screen(curItem)

    def On_LocalPointOfContacts_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_LocalPointOfContacts_Table_Refresh()

    def on_LocalPointOfContactsCalendarWindow_quit(self):
        self.data = self.cal.kill_and_save()
        self.LocalPointOfContactsCalendarExist = False
        self.LocalPointOfContactsCalendarWindow.destroy()
        if (len(self.data) > 0):
            self.date = str(self.data['month_selected']) + '/' + str(self.data['day_selected']) + '/' + str(self.data['year_selected'])
            #print (self.data['day_selected'])
            #print (self.data['month_selected'])
            #print (self.data['year_selected'])
            #print (self.data['day_name'])
            #print (self.data['month_name'])
        else:
            #print ("no date was selected")
            self.date = ""
        self.LocalPointOfContactsLeaseEndNameFrameEntry['state'] = 'normal'
        self.LocalPointOfContactsLeaseEndNameFrameEntry.delete(0,END)
        self.LocalPointOfContactsLeaseEndNameFrameEntry.insert(0,self.date)
        self.LocalPointOfContactsLeaseEndNameFrameEntry['state'] = 'readonly'
        
    def Call_Button_LocalPointOfContacts_Calendar_Lease_End(self):
        if not (self.LocalPointOfContactsCalendarExist):
            self.LocalPointOfContactsCalendarExist = True
            self.LocalPointOfContactsCalendarWindow = Tk()
            self.LocalPointOfContactsCalendarWindow.title("Lease End Calendar")
            self.LocalPointOfContactsCalendarWindow.protocol("WM_DELETE_WINDOW", self.on_LocalPointOfContactsCalendarWindow_quit)
            self.LocalPointOfContactsCalendarWindow.call('tk', 'scaling', self.Windows_Scaling)
            self.data = {}
            if (self.Selection == 'edit'):
                self.LocalPointOfContactsLeaseEndName = self.LocalPointOfContactsLeaseEndNameFrameEntry.get()
                curItem = self.LocalPointOfContactsTreeview.focus()  
                dic = self.LocalPointOfContactsTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.LocalPointOfContactsLeaseEndName) > 0):
                    if (self.LocalPointOfContactsTableLocalPointOfContactsLeaseEndDayArray[curItem] != 0):
                        self.data['day_selected'] = self.LocalPointOfContactsTableLocalPointOfContactsLeaseEndDayArray[curItem]
                        self.data['month_selected'] = self.LocalPointOfContactsTableLocalPointOfContactsLeaseEndMonthArray[curItem]
                        self.data['year_selected'] = self.LocalPointOfContactsTableLocalPointOfContactsLeaseEndYearArray[curItem]
            self.cal = Class_Calendar(self.LocalPointOfContactsCalendarWindow, self.data)
            self.cal.setup()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Local Point Of Contacts"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.LocalPointOfContactsTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts List',
                            message = '!!! The File was saved !!!')
                    self.LocalPointOfContactsWindow.title("Local Point Of Contacts                   File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Local Point Of Contacts','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_LocalPointOfContacts_Window(self): 
        if not self.LocalPointOfContactsWindowExist:
            # Set up the Window
            self.LocalPointOfContactsWindowExist = True
            self.LocalPointOfContactsWindow = Tk()
            self.LocalPointOfContactsWindow.geometry('1300x610+350+70')
            self.LocalPointOfContactsWindow.title("Local Point Of Contacts")
            self.LocalPointOfContactsWindow.protocol("WM_DELETE_WINDOW", self.on_LocalPointOfContactsWindow_quit)
            self.LocalPointOfContactsWindow.call('tk', 'scaling', self.Windows_Scaling)
            LocalPointOfContactsLabel = Label(self.LocalPointOfContactsWindow,text="Helvetica", font=("Helvetica", 19))
            LocalPointOfContactsLabel["text"] = "Local Point Of Contacts"
            LocalPointOfContactsLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.LocalPointOfContactsWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_LocalPointOfContactsWindow_quit)
            self.LocalPointOfContactsWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.LocalPointOfContactsFrame = Frame(self.LocalPointOfContactsWindow)
            self.LocalPointOfContactsFrame.pack(side=TOP, fill=BOTH, expand=Y)

            if self.db.Connect():

                    # set frame resizing priorities
                    self.LocalPointOfContactsFrame.rowconfigure(0, weight=1)
                    self.LocalPointOfContactsFrame.columnconfigure(0, weight=1)

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.LocalPointOfContactsFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 8, width=75, height=24)

                        self.ButtonCountryRefresh = Button(self.LocalPointOfContactsFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 8, width=75, height=24)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.LocalPointOfContactsFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 38, width=75, height=24)

                        self.ButtonRegionRefresh = Button(self.LocalPointOfContactsFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 38, width=75, height=24)

                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd = Button(self.LocalPointOfContactsFrame, text = '+ Facility', command = self.Display_Facility_Window, state=DISABLED)
                        self.ButtonFacilityAdd.place(x = 450, y = 68, width=75, height=24)

                        self.ButtonFacilityRefresh = Button(self.LocalPointOfContactsFrame, text = 'Refresh', command = self.on_Facility_Table_Refresh, state=DISABLED)
                        self.ButtonFacilityRefresh.place(x = 550, y = 68, width=75, height=24)


                    self.ButtonLocalPointOfContactsAdd = Button(self.LocalPointOfContactsFrame, text = 'Add', command = self.Call_Button_LocalPointOfContacts_Add, state=DISABLED)
                    self.ButtonLocalPointOfContactsAdd.place(x = 450, y = 98, width=75, height=25)

                    self.ButtonLocalPointOfContactsEdit = Button(self.LocalPointOfContactsFrame, text = 'Edit', command = self.Call_Button_LocalPointOfContacts_Edit, state=DISABLED)
                    self.ButtonLocalPointOfContactsEdit.place(x = 550, y = 98, width=75, height=25)

                    self.ButtonLocalPointOfContactsRemove = Button(self.LocalPointOfContactsFrame, text = 'Remove', command = self.Call_Button_LocalPointOfContacts_Remove, state=DISABLED)
                    self.ButtonLocalPointOfContactsRemove.place(x = 650, y = 98, width=75, height=25)

                    self.ButtonLocalPointOfContactsOK = Button(self.LocalPointOfContactsFrame, text = 'OK / UPDATE', command = self.Call_Button_LocalPointOfContacts_OK, state=DISABLED)
                    self.ButtonLocalPointOfContactsOK.place(x = 750, y = 98, width=100, height=25)

                    self.ButtonLocalPointOfContactsCancel = Button(self.LocalPointOfContactsFrame, text = 'Cancel', command = self.Call_Button_LocalPointOfContacts_Cancel, state=DISABLED)
                    self.ButtonLocalPointOfContactsCancel.place(x = 875, y = 98, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.LocalPointOfContactsFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=128)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=40)

                    RegionIDFrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Facility Name:"
                    RegionIDFrameLabel.place(x=10, y=70)

                    # Setup Labels and Entry
                    LocalPointOfContacts1FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    LocalPointOfContacts1FrameLabel["text"] = "Primary Point Of Contact:"
                    LocalPointOfContacts1FrameLabel.place(x=10, y=100)
                    self.LocalPointOfContacts1FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContacts1FrameEntry['width']=42
                    self.LocalPointOfContacts1FrameEntry.place(x=175, y=100)
                    self.LocalPointOfContacts1FrameEntry['state'] = DISABLED
                    
                    LocalPointOfContactsOfficeNo1FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    LocalPointOfContactsOfficeNo1FrameLabel["text"] = "Primary Office Number:"
                    LocalPointOfContactsOfficeNo1FrameLabel.place(x=10, y=130)
                    self.LocalPointOfContactsOfficeNo1FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContactsOfficeNo1FrameEntry['width']=42
                    self.LocalPointOfContactsOfficeNo1FrameEntry.place(x=175, y=130)
                    self.LocalPointOfContactsOfficeNo1FrameEntry['state'] = DISABLED               

                    self.LocalPointOfContactsAlternateNo1FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.LocalPointOfContactsAlternateNo1FrameLabel["text"] = "Alternate Number:"
                    self.LocalPointOfContactsAlternateNo1FrameLabel.place(x=10, y=160)
                    self.LocalPointOfContactsAlternateNo1FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContactsAlternateNo1FrameEntry['width']=42
                    self.LocalPointOfContactsAlternateNo1FrameEntry.place(x=175, y=160)
                    self.LocalPointOfContactsAlternateNo1FrameEntry['state'] = DISABLED               

                    self.LocalPointOfContacts2ndAlternateNo1FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.LocalPointOfContacts2ndAlternateNo1FrameLabel["text"] = "2nd Alternate Number:"
                    self.LocalPointOfContacts2ndAlternateNo1FrameLabel.place(x=450, y=160)
                    self.LocalPointOfContacts2ndAlternateNo1FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContacts2ndAlternateNo1FrameEntry['width']=42
                    self.LocalPointOfContacts2ndAlternateNo1FrameEntry.place(x=615, y=160)
                    self.LocalPointOfContacts2ndAlternateNo1FrameEntry['state'] = DISABLED               


###################33
                    LocalPointOfContacts2FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    LocalPointOfContacts2FrameLabel["text"] = "Secondary Point Of Contact:"
                    LocalPointOfContacts2FrameLabel.place(x=10, y=190)
                    self.LocalPointOfContacts2FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContacts2FrameEntry['width']=42
                    self.LocalPointOfContacts2FrameEntry.place(x=190, y=190)
                    self.LocalPointOfContacts2FrameEntry['state'] = DISABLED
                    
                    LocalPointOfContactsOfficeNo2FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    LocalPointOfContactsOfficeNo2FrameLabel["text"] = "Secondary Office Number:"
                    LocalPointOfContactsOfficeNo2FrameLabel.place(x=10, y=220)
                    self.LocalPointOfContactsOfficeNo2FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContactsOfficeNo2FrameEntry['width']=42
                    self.LocalPointOfContactsOfficeNo2FrameEntry.place(x=190, y=220)
                    self.LocalPointOfContactsOfficeNo2FrameEntry['state'] = DISABLED               

                    self.LocalPointOfContactsAlternateNo2FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.LocalPointOfContactsAlternateNo2FrameLabel["text"] = "Alternate Number:"
                    self.LocalPointOfContactsAlternateNo2FrameLabel.place(x=10, y=250)
                    self.LocalPointOfContactsAlternateNo2FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContactsAlternateNo2FrameEntry['width']=42
                    self.LocalPointOfContactsAlternateNo2FrameEntry.place(x=190, y=250)
                    self.LocalPointOfContactsAlternateNo2FrameEntry['state'] = DISABLED               

                    self.LocalPointOfContacts2ndAlternateNo2FrameLabel = Label(self.LocalPointOfContactsFrame,text="Helvetica", font=("Helvetica", 10))
                    self.LocalPointOfContacts2ndAlternateNo2FrameLabel["text"] = "2nd Alternate Number:"
                    self.LocalPointOfContacts2ndAlternateNo2FrameLabel.place(x=465, y=250)
                    self.LocalPointOfContacts2ndAlternateNo2FrameEntry = Entry(self.LocalPointOfContactsFrame)
                    self.LocalPointOfContacts2ndAlternateNo2FrameEntry['width']=42
                    self.LocalPointOfContacts2ndAlternateNo2FrameEntry.place(x=630, y=250)
                    self.LocalPointOfContacts2ndAlternateNo2FrameEntry['state'] = DISABLED               


####################


                    
                    #------------------ TREE VIEW For LocalPointOfContacts Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.LocalPointOfContactsTreeviewDataColumns = ('Primay POC','Office No','Alternate','2nd Alternate',
                                                                    'Secondary POC','Office No','Alternate','2nd Alternate','Added By','Date & Time')
                    
                    self.LocalPointOfContactsTreeview = ttk.Treeview(self.LocalPointOfContactsFrame,columns=self.LocalPointOfContactsTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.LocalPointOfContactsTreeviewysb = Scrollbar(self.LocalPointOfContactsFrame,orient=VERTICAL, command=self.LocalPointOfContactsTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.LocalPointOfContactsTreeviewxsb = Scrollbar(self.LocalPointOfContactsFrame,orient=HORIZONTAL, command=self.LocalPointOfContactsTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.LocalPointOfContactsTreeview['yscroll'] = self.LocalPointOfContactsTreeviewysb.set
                    self.LocalPointOfContactsTreeview['xscroll'] = self.LocalPointOfContactsTreeviewxsb.set

                    # setup headings and column 
                    self.LocalPointOfContactsTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                    self.LocalPointOfContactsTreeview.heading('#1', text='Primay POC',    anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Primay POC', False)) # E for East and W for West
                    self.LocalPointOfContactsTreeview.heading('#2', text='Office No',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Office No', False)) # E for East and W for West
                    self.LocalPointOfContactsTreeview.heading('#3', text='Alternate',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Alternate', False)) 
                    self.LocalPointOfContactsTreeview.heading('#4', text='2nd Alternate', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, '2nd Alternate', False)) 
                    self.LocalPointOfContactsTreeview.heading('#5', text='Secondary POC', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Secondary POC', False)) # E for East and W for West
                    self.LocalPointOfContactsTreeview.heading('#6', text='Office No',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Office No', False)) # E for East and W for West
                    self.LocalPointOfContactsTreeview.heading('#7', text='Alternate',     anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Alternate', False)) 
                    self.LocalPointOfContactsTreeview.heading('#8', text='2nd Alternate', anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, '2nd Alternate', False)) 
                    self.LocalPointOfContactsTreeview.heading('#9', text='Added By',      anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Added By', False))
                    self.LocalPointOfContactsTreeview.heading('#10', text='Date & Time',  anchor=W,command=lambda: self.treeview_sort_column(self.LocalPointOfContactsTreeview, 'Date & Time', False))
                    
                    
                    self.LocalPointOfContactsTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.LocalPointOfContactsTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.LocalPointOfContactsTreeview.column('#2', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#3', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#4', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#5', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#6', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#7', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#8', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#9', stretch=1, width=10)
                    self.LocalPointOfContactsTreeview.column('#10', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.LocalPointOfContactsTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.LocalPointOfContactsTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.LocalPointOfContactsTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    # Use later to mark Business Units per color.
                    self.LocalPointOfContactsTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.LocalPointOfContactsTreeview.bind('<ButtonRelease-1>', self.on_LocalPointOfContacts_Tree_select_click) # When Select the Tree
                    #self.LocalPointOfContactsTreeview.bind("<Double-1>", self.On_LocalPointOfContacts_Tree_Refresh)         

                    #------------------ TREE VIEW For LocalPointOfContacts Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.LocalPointOfContactsFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.LocalPointOfContactsFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 40)

                    self.ComboBoxFacilityID = ttk.Combobox(self.LocalPointOfContactsFrame, state='disabled', width = 50)
                    self.ComboBoxFacilityID.bind("<<ComboboxSelected>>", self.on_facility_combo_changed)
                    self.ComboBoxFacilityID.place(x = 110, y = 70)

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
                                    self.on_facility_combo_changed("event")
                                '''
                                    self.on_Sites_Table_Refresh()
                                    if (self.sql_querry):
                                        self.ComboBoxSitesID.current(0)
                                        self.on_sites_combo_changed("event")
                                '''
                        self.Selection = 'edit_ok'
                    else:                        
                        self.on_Country_Table_Refresh()
                    #self.on_Country_Table_Refresh()
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.LocalPointOfContactsWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.LocalPointOfContactsFrame,title='Local Point Of Contacts',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the LocalPointOfContacts Class....:")
    Location = []
    #Location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']    
    LocalPointOfContacts = Class_LocalPointOfContacts("BV",Windows_Scaling,Location)
    LocalPointOfContacts.Display_LocalPointOfContacts_Window()


if __name__ == '__main__':
    Main()

