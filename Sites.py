#==========================================================================
# Program: Sites.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-14-2018
# Date Last Modified: Jan-14-2018
# Summary: This is Class to for the Sites
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
    Parameter = ['Sites','OPEN Window']    
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

class Class_Sites:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.SitesWindowExist = False
        self.SitesCalendarExist = False
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

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************
        
    def Clean_Screen(self,option,option2):
        # Setup Buttons

        #self.SitesBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False
        
        if (option == 'country'): ## The option are country,region and Sites
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
            self.ButtonSitesAdd['state'] = ACTIVE
        else:
            self.ButtonSitesAdd['state'] = DISABLED
            
        self.ButtonSitesEdit['state'] = DISABLED
        self.ButtonSitesRemove['state'] = DISABLED
        self.ButtonSitesOK['state'] = DISABLED
        self.ButtonSitesCancel['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.SitesIDFrameEntry['state'] = 'normal'
        self.SitesIDFrameEntry.delete(0,END)
        self.SitesIDFrameEntry['state'] = 'readonly'

        self.SitesNameFrameEntry['state'] = 'normal'
        self.SitesNameFrameEntry.delete(0,END)
        self.SitesNameFrameEntry['state'] = 'readonly'

        self.SitesDescriptionFrameEntry['state'] = 'normal'
        self.SitesDescriptionFrameEntry.delete(0,END)
        self.SitesDescriptionFrameEntry['state'] = 'readonly'

        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.SitesTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.SitesTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------

    def Display_Screen(self,curItem):      
            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.SitesIDFrameEntry['state'] = 'normal'
            self.SitesIDFrameEntry.delete(0,END)
            self.SitesIDFrameEntry.insert(0,self.SitesTableSiteIDArray[curItem])
            self.SitesIDFrameEntry['state'] = 'readonly'

            self.SitesNameFrameEntry['state'] = 'normal'
            self.SitesNameFrameEntry.delete(0,END)
            self.SitesNameFrameEntry.insert(0,self.SitesTableSiteNameArray[curItem])
            self.SitesNameFrameEntry['state'] = 'readonly'

            self.SitesDescriptionFrameEntry['state'] = 'normal'
            self.SitesDescriptionFrameEntry.delete(0,END)
            self.SitesDescriptionFrameEntry.insert(0,self.SitesTableDescriptionArray[curItem])
            self.SitesDescriptionFrameEntry['state'] = 'readonly'


    def Enable_Screen(self,option):
        # This function is used when the ADD button is selected

        #self.SitesBusinessUnitPowerCheckbutton.select()
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
            
        self.ButtonSitesAdd['state'] = DISABLED
        self.ButtonSitesEdit['state'] = DISABLED
        self.ButtonSitesRemove['state'] = DISABLED
        self.ButtonSitesOK['state'] = ACTIVE
        self.ButtonSitesCancel['state'] = ACTIVE

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        if (option == 'add'): #<----------------------------------- ADD Button
            self.SitesIDFrameEntry['state'] = 'normal'
            self.SitesIDFrameEntry.delete(0,END)

            self.SitesNameFrameEntry['state'] = 'normal'
            self.SitesNameFrameEntry.delete(0,END)

            self.SitesDescriptionFrameEntry['state'] = 'normal'
            self.SitesDescriptionFrameEntry.delete(0,END)

        if (option == 'edit'): #<----------------------------------- EDIT Button
            self.SitesIDFrameEntry['state'] = 'readonly'
            self.SitesNameFrameEntry['state'] = 'normal'
            self.SitesDescriptionFrameEntry['state'] = 'normal'
                
       
    def Disable_Screen(self):
        # This function is used when the entry was added.modified to the Database

        #self.SitesBusinessUnitPowerCheckbutton.select()
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
            
        self.ButtonSitesAdd['state'] = ACTIVE
        self.ButtonSitesEdit['state'] = DISABLED
        self.ButtonSitesRemove['state'] = DISABLED
        self.ButtonSitesOK['state'] = DISABLED
        self.ButtonSitesCancel['state'] = ACTIVE

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.SitesIDFrameEntry['state'] = 'readonly'
        self.SitesNameFrameEntry['state'] = 'readonly'
        self.SitesDescriptionFrameEntry['state'] = 'readonly'


    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
        self.SitesID = self.SitesIDFrameEntry.get()
        self.SitesName = self.SitesNameFrameEntry.get()
        self.SitesDescription = self.SitesDescriptionFrameEntry.get()

#****************************************************************************************
#---------------------------- SCREEN SELECTION SECTION <END> ------------------------*
#****************************************************************************************


#****************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************

    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # SQL Querry to the Sites Table
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
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the Sites Table
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
                    #self.ComboBoxCoutryID['state'] = 'readonly'
                    #self.ComboBoxCoutryID.set("")
                    self.Clean_Screen('country','all')
            else:
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Country Records found')
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
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
            Parameter = ['Sites','Region Window']    
            Logging.Log(Parameter)
        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        self.Clean_Screen('region','region-combo')
        if self.db.Connect():
            # SQL Querry to the Sites Table
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
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','Region Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the Sites Table
            sql = """
                    SELECT * FROM Region
                    WHERE Country_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()])
            #print (sql)
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
                    if (Is_Region_Available):
                        self.ButtonRegionAdd['state'] = 'active'
                        self.ButtonRegionRefresh['state'] = 'active'
            else:
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Records found')
                self.ComboBoxRegionID['state'] = DISABLED
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonSitesAdd['state']    = DISABLED
                self.ButtonSitesEdit['state']   = DISABLED
                self.ButtonSitesRemove['state'] = DISABLED
                self.ButtonSitesOK['state']     = DISABLED
                self.ButtonSitesCancel['state'] = DISABLED

                self.SitesIDFrameEntry['state'] = 'normal'
                self.SitesIDFrameEntry.delete(0,END)
                self.SitesIDFrameEntry['state'] = 'readonly'
                
                self.SitesNameFrameEntry['state'] = 'normal'
                self.SitesNameFrameEntry.delete(0,END)
                self.SitesNameFrameEntry['state'] = 'readonly'

        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
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
            Parameter = ['Sites','Facility Window']    
            Logging.Log(Parameter)
        Facility.Display_Facility_Window()

    def on_facility_combo_changed(self,event):
        self.Clean_Screen('facility','facility-combo')
        if self.db.Connect():
            # SQL Querry to the Sites Table
            sql = """
                    SELECT * FROM Sites
                    WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()],
                         self.FacilityIDArray[self.ComboBoxFacilityID.current()])
            #print (sql)
            if (Is_Facility_Available):
                self.ButtonFacilityAdd['state'] = ACTIVE
                self.ButtonFacilityRefresh['state'] = ACTIVE
            if (self.db.Execute(sql)):
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.SitesTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.SitesTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                #-------------- Initializing Arrays <BEGIN> ----------------------
                self.SitesTablePriaryKeyArray =  []
                self.SitesTableCountryIDArray =  [] 
                self.SitesTableRegionIDArray =  []
                self.SitesTableFacilityIDArray =  []
                self.SitesTableSiteIDArray =  []
                self.SitesTableSiteNameArray =  []
                self.SitesTableDescriptionArray =  []
                self.SitesTableExecutedByUserArray = []
                self.results = []
                #-------------- Initializing Arrays <END> ----------------------
                self.data_ready = True
                i = 0
                while (i < len(self.db.results)):
                    self.SitesTablePriaryKeyArray.append(self.db.results[i][0].strip())
                    self.SitesTableCountryIDArray.append(self.db.results[i][1].strip()) 
                    self.SitesTableRegionIDArray.append(self.db.results[i][2].strip())
                    self.SitesTableFacilityIDArray.append(self.db.results[i][3].strip())
                    self.SitesTableSiteIDArray.append(self.db.results[i][4].strip())
                    self.SitesTableSiteNameArray.append(self.db.results[i][5].strip())
                    self.SitesTableDescriptionArray.append(self.db.results[i][6].strip())
                    self.SitesTableExecutedByUserArray.append(self.db.results[i][7].strip())
                    i = i + 1
                i = 0
                while (i < len(self.SitesTablePriaryKeyArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            #self.SitesTablePriaryKeyArray[i],
                            #self.SitesTableCountryIDArray[i], 
                            #self.SitesTableRegionIDArray[i],
                            #self.SitesTableFacilityIDArray[i],
                            self.SitesTableSiteIDArray[i],
                            self.SitesTableSiteNameArray[i],
                            self.SitesTableDescriptionArray[i],
                            self.SitesTableExecutedByUserArray[i]]
                    self.results.append(item)
                    self.SitesTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonSitesAdd['state']    = ACTIVE
                self.ButtonSitesEdit['state']   = DISABLED
                self.ButtonSitesRemove['state'] = DISABLED
                self.ButtonSitesOK['state']     = DISABLED
                self.ButtonSitesCancel['state'] = DISABLED
            else:
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Records found')
                self.data_ready = False
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','Facility Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('facility','all')
        if self.db.Connect():
            # SQL Querry to the Sites Table
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            #print (sql)
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
                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd['state'] = 'active'
                        self.ButtonFacilityRefresh['state'] = 'active'
            else:
                mbox.showerror(master=self.SitesFrame,title='Sites',
                message = 'No Records found')
                self.ComboBoxFacilityID['state'] = DISABLED
                
                self.ButtonSitesAdd['state']    = DISABLED
                self.ButtonSitesEdit['state']   = DISABLED
                self.ButtonSitesRemove['state'] = DISABLED
                self.ButtonSitesOK['state']     = DISABLED
                self.ButtonSitesCancel['state'] = DISABLED

                self.SitesIDFrameEntry['state'] = 'normal'
                self.SitesIDFrameEntry.delete(0,END)
                self.SitesIDFrameEntry['state'] = 'readonly'
                
                self.SitesNameFrameEntry['state'] = 'normal'
                self.SitesNameFrameEntry.delete(0,END)
                self.SitesNameFrameEntry['state'] = 'readonly'

        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- FACILITY SELECTION SECTION <END> ------------------------*
#*************************************************************************************

        
    def on_SitesWindow_quit(self):
        if (self.SitesWindowExist):
            self.SitesWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Sites','CLOSE Window']    
                Logging.Log(Parameter)
            self.SitesWindow.destroy()


    def on_Sites_Table_Refresh(self): # I need to do more research on this call.
        self.on_country_combo_changed("event")
       
    def Call_Button_Sites_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','ADD Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.Enable_Screen('add')
        self.Selection = 'add'

    def Call_Button_Sites_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        self.Enable_Screen('edit')

    def Call_Button_Sites_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
            self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
            self.FacilityID = self.FacilityIDArray[self.ComboBoxFacilityID.current()]
            self.SitesID = self.SitesIDFrameEntry.get()
            PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID+"-"+self.SitesID)
            if (mbox.askyesnocancel(master=self.SitesFrame,title='Sites',message = 'Are you Sure you want to Remove it?')):
                #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.SitesID)
                #print (PrimaryKey)
                if (self.SitesID == 'UNKNOWN'):
                    mbox.showerror(master=self.SitesFrame,title='Sites',
                        message = '*** The Sites ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Sites --------
                    sql = """
                            SELECT * FROM DEVICES
                            WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' AND Site_ID = '%s'
                          """ % (self.CountryID,self.RegionID,self.FacilityID,self.SitesID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.SitesFrame,title='Sites',
                                    message = '*** The Sites ID you entered CANNOT be Removed Becuase is used in Devices ***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                    else:
                        sql = """
                                SELECT * FROM Sites
                                WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'
                              """ % (PrimaryKey)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Sites WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'" % (PrimaryKey)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.SitesFrame,title='Sites',
                                        message = '*** The Site ID you entered was Removed ***')
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.SitesFrame,title='Sites',
                                        message = '*** The Site ID you entered was NOT Removed ***')               
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            mbox.showerror(master=self.SitesFrame,title='Sites',
                            message = '*** The Site ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Sites_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ButtonSitesAdd['state']    = DISABLED
            self.ButtonSitesEdit['state']   = DISABLED
            self.ButtonSitesRemove['state'] = DISABLED
            self.ButtonSitesOK['state']     = ACTIVE
            self.ButtonSitesCancel['state'] = ACTIVE
            self.Collect_Screen()
            PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID+"-"+self.SitesID)
            #self.Selection = 'notyet'
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(self.SitesID) > 0) and (len(self.SitesName) > 0)):
                    sql = """
                            SELECT * FROM Sites
                            WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.SitesFrame,title='Sites',
                        message = '*** The Sites ID you entered already exist ***')
                    else:
                        if (len(self.SitesDescription) > 100):
                            self.SitesDescription = self.SitesDescription[:100]
                        sql = "INSERT INTO Sites(Country_ID_Region_ID_Facility_ID_Site_ID, Country_ID, Region_ID, Facility_ID, Site_ID, Site_Name, \
                                   Description, Executed_by_UserID) \
                                   VALUES ('%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.CountryID,self.RegionID,self.FacilityID,self.SitesID,
                                                                                        self.SitesName,self.SitesDescription,self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.SitesFrame,title='Sites',
                                message = '*** The Sites ID you entered was Added ***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.SitesFrame,title='Sites',
                                message = '*** The Sites ID you entered was NOT Added ***')
                            self.on_facility_combo_changed("event")
                            self.Disable_Screen()
                else:
                     mbox.showerror(master=self.SitesFrame,title='Sites',
                        message = '*** The Sites ID and Sites Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                #print (PrimaryKey+"....")
                if (len(self.SitesName) == 0):
                    mbox.showerror(master=self.SitesFrame,title='Sites',
                    message = '*** The Sites Name Cannot be BLANK ***')
                else:
                    sql = """
                            SELECT * FROM Sites
                            WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'
                          """ % (PrimaryKey)                   
                    if (self.db.Execute(sql)):   
                        #sql = "UPDATE Sites SET Sites_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID_Sites_ID = '%s'" % (SitesName,self.Username,PrimaryKey)
                        self.SitesHROfficeID = "-"
                        if (len(self.SitesDescription) > 100):
                            self.SitesDescription = self.SitesDescription[:100]
                        sql = "UPDATE Sites SET Site_Name  = '%s', Description = '%s', Executed_by_UserID = '%s' \
                               WHERE Country_ID_Region_ID_Facility_ID_Site_ID = '%s'" %(self.SitesName,self.SitesDescription,self.Username,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.SitesFrame,title='Sites',
                                message = '*** The Sites ID you entered was Updated ***')
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.SitesFrame,title='Sites',
                                message = '*** The Sites ID you entered was NOT Upadted ***')
                        self.on_facility_combo_changed("event")
                        self.Disable_Screen()
                    else:
                        mbox.showerror(master=self.SitesFrame,title='Sites',
                        message = '*** The Sites ID you try to Edit Does not exist Anymore ***')
                        self.on_facility_combo_changed("event")
                        self.Disable_Screen()                        
        else:
            mbox.showerror(master=self.SitesFrame,title='Sites',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Sites_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Sites','CANCEL Button']    
            Logging.Log(Parameter)
        self.Clean_Screen('Sites','all')
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

    def on_Sites_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.SitesTreeview.focus()  
        dic = self.SitesTreeview.item(curItem)
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
            self.ButtonSitesAdd['state']    = ACTIVE
            self.ButtonSitesEdit['state']   = ACTIVE
            self.ButtonSitesRemove['state'] = ACTIVE
            self.ButtonSitesOK['state']     = DISABLED
            self.ButtonSitesCancel['state'] = DISABLED
            self.Display_Screen(curItem)

    def On_Sites_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_Sites_Table_Refresh()

    def on_SitesCalendarWindow_quit(self):
        self.data = self.cal.kill_and_save()
        self.SitesCalendarExist = False
        self.SitesCalendarWindow.destroy()
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
        self.SitesLeaseEndNameFrameEntry['state'] = 'normal'
        self.SitesLeaseEndNameFrameEntry.delete(0,END)
        self.SitesLeaseEndNameFrameEntry.insert(0,self.date)
        self.SitesLeaseEndNameFrameEntry['state'] = 'readonly'
        
    def Call_Button_Sites_Calendar_Lease_End(self):
        if not (self.SitesCalendarExist):
            self.SitesCalendarExist = True
            self.SitesCalendarWindow = Tk()
            self.SitesCalendarWindow.title("Lease End Calendar")
            self.SitesCalendarWindow.protocol("WM_DELETE_WINDOW", self.on_SitesCalendarWindow_quit)
            self.SitesCalendarWindow.call('tk', 'scaling', self.Windows_Scaling)
            self.data = {}
            if (self.Selection == 'edit'):
                self.SitesLeaseEndName = self.SitesLeaseEndNameFrameEntry.get()
                curItem = self.SitesTreeview.focus()  
                dic = self.SitesTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.SitesLeaseEndName) > 0):
                    if (self.SitesTableSitesLeaseEndDayArray[curItem] != 0):
                        self.data['day_selected'] = self.SitesTableSitesLeaseEndDayArray[curItem]
                        self.data['month_selected'] = self.SitesTableSitesLeaseEndMonthArray[curItem]
                        self.data['year_selected'] = self.SitesTableSitesLeaseEndYearArray[curItem]
            self.cal = Class_Calendar(self.SitesCalendarWindow, self.data)
            self.cal.setup()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Sites List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.SitesTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.SitesFrame,title='Sites List',
                            message = '!!! The File was saved !!!')
                    self.SitesWindow.title("Sites                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Sites','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.SitesFrame,title='Sites List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Display_Sites_Window(self): 
        if not self.SitesWindowExist:
            # Set up the Window
            self.SitesWindowExist = True
            self.SitesWindow = Tk()
            self.SitesWindow.geometry('1300x550+350+70')
            self.SitesWindow.title("Sites")
            self.SitesWindow.protocol("WM_DELETE_WINDOW", self.on_SitesWindow_quit)
            self.SitesWindow.call('tk', 'scaling', self.Windows_Scaling)
            SitesLabel = Label(self.SitesWindow,text="Helvetica", font=("Helvetica", 19))
            SitesLabel["text"] = "Site"
            SitesLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.SitesWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_SitesWindow_quit)
            self.SitesWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.SitesFrame = Frame(self.SitesWindow)
            self.SitesFrame.pack(side=TOP, fill=BOTH, expand=Y)

            if self.db.Connect():

                    # set frame resizing priorities
                    self.SitesFrame.rowconfigure(0, weight=1)
                    self.SitesFrame.columnconfigure(0, weight=1)

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.SitesFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 8, width=75, height=24)

                        self.ButtonCountryRefresh = Button(self.SitesFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 8, width=75, height=24)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.SitesFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 38, width=75, height=24)

                        self.ButtonRegionRefresh = Button(self.SitesFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 38, width=75, height=24)

                    if (Is_Facility_Available):
                        self.ButtonFacilityAdd = Button(self.SitesFrame, text = '+ Facility', command = self.Display_Facility_Window, state=DISABLED)
                        self.ButtonFacilityAdd.place(x = 450, y = 68, width=75, height=24)

                        self.ButtonFacilityRefresh = Button(self.SitesFrame, text = 'Refresh', command = self.on_Facility_Table_Refresh, state=DISABLED)
                        self.ButtonFacilityRefresh.place(x = 550, y = 68, width=75, height=24)


                    self.ButtonSitesAdd = Button(self.SitesFrame, text = 'Add', command = self.Call_Button_Sites_Add, state=DISABLED)
                    self.ButtonSitesAdd.place(x = 450, y = 98, width=75, height=25)

                    self.ButtonSitesEdit = Button(self.SitesFrame, text = 'Edit', command = self.Call_Button_Sites_Edit, state=DISABLED)
                    self.ButtonSitesEdit.place(x = 550, y = 98, width=75, height=25)

                    self.ButtonSitesRemove = Button(self.SitesFrame, text = 'Remove', command = self.Call_Button_Sites_Remove, state=DISABLED)
                    self.ButtonSitesRemove.place(x = 650, y = 98, width=75, height=25)

                    self.ButtonSitesOK = Button(self.SitesFrame, text = 'OK / UPDATE', command = self.Call_Button_Sites_OK, state=DISABLED)
                    self.ButtonSitesOK.place(x = 750, y = 98, width=100, height=25)

                    self.ButtonSitesCancel = Button(self.SitesFrame, text = 'Cancel', command = self.Call_Button_Sites_Cancel, state=DISABLED)
                    self.ButtonSitesCancel.place(x = 875, y = 98, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.SitesFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=128)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=40)

                    RegionIDFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Facility Name:"
                    RegionIDFrameLabel.place(x=10, y=70)

                    # Setup Labels and Entry
                    SitesIDFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    SitesIDFrameLabel["text"] = "Site ID:"
                    SitesIDFrameLabel.place(x=10, y=100)
                    self.SitesIDFrameEntry = Entry(self.SitesFrame)
                    self.SitesIDFrameEntry['width']=50
                    self.SitesIDFrameEntry.place(x=110, y=100)
                    self.SitesIDFrameEntry['state'] = DISABLED

                    
                    SitesNameFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    SitesNameFrameLabel["text"] = "Site Name:"
                    SitesNameFrameLabel.place(x=10, y=130)
                    self.SitesNameFrameEntry = Entry(self.SitesFrame)
                    self.SitesNameFrameEntry['width']=50
                    self.SitesNameFrameEntry.place(x=110, y=130)
                    self.SitesNameFrameEntry['state'] = DISABLED               

                    self.SitesDescriptionFrameLabel = Label(self.SitesFrame,text="Helvetica", font=("Helvetica", 10))
                    self.SitesDescriptionFrameLabel["text"] = "Description:"
                    self.SitesDescriptionFrameLabel.place(x=10, y=160)
                    self.SitesDescriptionFrameEntry = Entry(self.SitesFrame)
                    self.SitesDescriptionFrameEntry['width']=100
                    self.SitesDescriptionFrameEntry.place(x=110, y=160)
                    self.SitesDescriptionFrameEntry['state'] = DISABLED               
                    
                    #------------------ TREE VIEW For Sites Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.SitesTreeviewDataColumns = ('Site ID','Site Name','Description','Added By')
                    
                    self.SitesTreeview = ttk.Treeview(self.SitesFrame,columns=self.SitesTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.SitesTreeviewysb = Scrollbar(self.SitesFrame,orient=VERTICAL, command=self.SitesTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.SitesTreeviewxsb = Scrollbar(self.SitesFrame,orient=HORIZONTAL, command=self.SitesTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.SitesTreeview['yscroll'] = self.SitesTreeviewysb.set
                    self.SitesTreeview['xscroll'] = self.SitesTreeviewxsb.set

                    # setup headings and column 
                    self.SitesTreeview.heading('#0', text='Item No.',    anchor=W) # E for East and W for West and CENTER
                    self.SitesTreeview.heading('#1', text='Site ID',     anchor=W,command=lambda: self.treeview_sort_column(self.SitesTreeview, 'Site ID', False)) # E for East and W for West
                    self.SitesTreeview.heading('#2', text='Site Name',   anchor=W,command=lambda: self.treeview_sort_column(self.SitesTreeview, 'Site Name', False)) # E for East and W for West
                    self.SitesTreeview.heading('#3', text='Description', anchor=W,command=lambda: self.treeview_sort_column(self.SitesTreeview, 'Description', False)) 
                    self.SitesTreeview.heading('#4', text='Added By',    anchor=W,command=lambda: self.treeview_sort_column(self.SitesTreeview, 'Added By', False)) 
                    
                    self.SitesTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.SitesTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.SitesTreeview.column('#2', stretch=1, width=10)
                    self.SitesTreeview.column('#3', stretch=1, width=10)
                    self.SitesTreeview.column('#4', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.SitesTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.SitesTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.SitesTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    # Use later to mark Business Units per color.
                    self.SitesTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.SitesTreeview.bind('<ButtonRelease-1>', self.on_Sites_Tree_select_click) # When Select the Tree
                    #self.SitesTreeview.bind("<Double-1>", self.On_Sites_Tree_Refresh)         

                    #------------------ TREE VIEW For Sites Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.SitesFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.SitesFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 40)

                    self.ComboBoxFacilityID = ttk.Combobox(self.SitesFrame, state='disabled', width = 50)
                    self.ComboBoxFacilityID.bind("<<ComboboxSelected>>", self.on_facility_combo_changed)
                    self.ComboBoxFacilityID.place(x = 110, y = 70)

                    self.on_Country_Table_Refresh()
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.SitesWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.SitesFrame,title='Sites',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Sites Class....:")
    Sites = Class_Sites("BV",Windows_Scaling)
    Sites.Display_Sites_Window()


if __name__ == '__main__':
    Main()

