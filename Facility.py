#==========================================================================
# Program: Facility.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-30-2017
# Date Last Modified: Feb-18-2017
# Summary: This is Class to for the Facility
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
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Facility','OPEN Window']    
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

class Class_Facility:

    def __init__(self,DSN_Name,Windows_Scaling,Location):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.FacilityWindowExist = False
        self.FacilityCalendarExist = False
        self.Username = os.getlogin()
        self.date = ""
        self.Windows_Scaling = Windows_Scaling
        self.Selection = 'none'
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
#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------*
#****************************************************************************************
        
    def Clean_Screen(self,option,option2):
        # Setup Buttons

        #self.FacilityBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())
        self.data_ready = False

        if (option == 'country'): ## The option are country,region and facility
            if (option2 != 'country-combo'):
                self.ComboBoxCoutryID.set("")
            self.ComboBoxRegionID.set("")
            self.ComboBoxRegionID['state'] = DISABLED
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = DISABLED
                self.ButtonRegionRefresh['state'] = DISABLED

        if ((option == 'country') or (option == 'region') or (option == 'facility')): ## The option are country,region and facility
            if (option == 'region'):
                self.ComboBoxRegionID.set("") 
            
            self.ButtonFacilityAdd['state'] = DISABLED
            self.ButtonFacilityEdit['state'] = DISABLED
            self.ButtonFacilityRemove['state'] = DISABLED
            self.ButtonFacilityOK['state'] = DISABLED
            self.ButtonFacilityCancel['state'] = DISABLED
            self.ButtonFacilityValidate['state'] = DISABLED
            self.ButtonFacilityValidate['state'] = DISABLED

            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.FacilityIDFrameEntry['state'] = 'normal'
            self.FacilityIDFrameEntry.delete(0,END)
            self.FacilityIDFrameEntry['state'] = 'readonly'

            self.FacilityNameFrameEntry['state'] = 'normal'
            self.FacilityNameFrameEntry.delete(0,END)
            self.FacilityNameFrameEntry['state'] = 'readonly'
                    
            self.ComboBoxOfficeType['state'] = 'disabled'
            self.ComboBoxOfficeSize['state'] = 'disabled'

            self.FacilityLeaseEndNameFrameEntry['state'] = 'normal'
            self.FacilityLeaseEndNameFrameEntry.delete(0,END)
            self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
            self.ButtonFacilityLeaseEnd['state'] = 'disabled'

            self.FacilityOfficeAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeAddress1FrameEntry['state'] = 'readonly'
            self.FacilityOfficeAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress2FrameEntry.delete(0,END)
            self.FacilityOfficeAddress2FrameEntry['state'] = 'readonly'

            self.FacilityOfficeCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCityFrameEntry.delete(0,END)               
            self.FacilityOfficeCityFrameEntry['state'] = 'readonly'

            self.FacilityOfficeStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeStateFrameEntry.delete(0,END)               
            self.FacilityOfficeStateFrameEntry['state'] = 'readonly'

            self.FacilityOfficeCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCountryFrameEntry.delete(0,END)               
            self.FacilityOfficeCountryFrameEntry['state'] = 'readonly'

            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeZIPCodeFrameEntry.delete(0,END)               
            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'readonly'

            self.FacilityOfficeHeadcountFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHeadcountFrameEntry.delete(0,END)               
            self.FacilityOfficeHeadcountFrameEntry['state'] = 'readonly'

            self.FacilityOfficeHRIDFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHRIDFrameEntry.delete(0,END)               
            self.FacilityOfficeHRIDFrameEntry['state'] = 'readonly'

            self.ComboBoxOfficeStatus['state'] = 'disabled'

            self.FacilityStrategyFrameEntry['state'] = 'normal'
            self.FacilityStrategyFrameEntry.delete(0,END)
            self.FacilityStrategyFrameEntry['state'] = 'readonly'

            self.FacilityActionFrameEntry['state'] = 'normal'
            self.FacilityActionFrameEntry.delete(0,END)
            self.FacilityActionFrameEntry['state'] = 'readonly'

            self.FacilityNotesFrameEntry['state'] = 'normal'
            self.FacilityNotesFrameEntry.delete(0,END)
            self.FacilityNotesFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'readonly'
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress2FrameEntry.delete(0,END)
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCityFrameEntry.delete(0,END)               
            self.FacilityOfficeNewCityFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewStateFrameEntry.delete(0,END)               
            self.FacilityOfficeNewStateFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCountryFrameEntry.delete(0,END)               
            self.FacilityOfficeNewCountryFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewZIPCodeFrameEntry.delete(0,END)               
            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'readonly'


            # CheckButton
            self.FacilityBusinessUnitPowerCheckbutton.deselect()
            self.FacilityBusinessUnitWaterCheckbutton.deselect()
            self.FacilityBusinessUnitTelecomCheckbutton.deselect()
            self.FacilityBusinessUnitSPCheckbutton.deselect()
            self.FacilityBusinessUnitCorpCheckbutton.deselect()
            self.FacilityBusinessUnitDCCheckbutton.deselect()
            self.FacilityBusinessUnitWHCheckbutton.deselect()
            
            self.FacilityBusinessUnitPowerCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitWaterCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitTelecomCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitSPCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitCorpCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitDCCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitWHCheckbutton['state'] = 'disable'

            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.FacilityTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.FacilityTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------

    def Display_Screen(self,curItem):
            #self.FacilityTablePriaryKeyArray[curItem],
            #self.FacilityTableCountryIDArray[curItem], 
            #self.FacilityTableRegionIDArray[curItem],
            #self.FacilityTableFacilityIDArray[curItem],
            #self.FacilityTableFacilityNameArray[curItem],
            #self.FacilityTableBusinessUnitArray[curItem],
            #self.FacilityTableOfficeTypeArray[curItem],
            #self.FacilityTableOfficeSizeArray[curItem],
            #self.FacilityTableFacilityLeaseEndNameArray[curItem],
            #self.FacilityTableFacilityLeaseEndMonthArray[curItem],
            #self.FacilityTableFacilityLeaseEndDayArray[curItem],
            #self.FacilityTableFacilityLeaseEndYearArray[curItem],
            #self.FacilityTableFacilityOfficeAddress1Array[curItem],
            #self.FacilityTableFacilityOfficeAddress2Array[curItem],
            #self.FacilityTableFacilityOfficeCityArray[curItem],
            #self.FacilityTableFacilityOfficeStateArray[curItem],
            #self.FacilityTableFacilityOfficeZIPCodeArray[curItem],
            #self.FacilityTableFacilityOfficeCountryArray[curItem],  
            #self.FacilityTableFacilityOfficeHeadcountArray[curItem],
            #self.FacilityTableFacilityOfficeHRIDLocationArray[curItem],
            #self.FacilityTableFacilityOfficeHRIDArray[curItem],
            #self.FacilityTableOfficeStatusArray[curItem],        
            #self.FacilityTableExecutedByUserArray[curItem]]
      
            # Create Progress Bar
            self.progress['maximum'] = 100
            self.progress['value'] = 0

            # Setup Labels and Entry
            self.FacilityIDFrameEntry['state'] = 'normal'
            self.FacilityIDFrameEntry.delete(0,END)
            self.FacilityIDFrameEntry.insert(0,self.FacilityTableFacilityIDArray[curItem])
            self.FacilityIDFrameEntry['state'] = 'readonly'

            self.FacilityNameFrameEntry['state'] = 'normal'
            self.FacilityNameFrameEntry.delete(0,END)
            self.FacilityNameFrameEntry.insert(0,self.FacilityTableFacilityNameArray[curItem])
            self.FacilityNameFrameEntry['state'] = 'readonly'

            #self.OfficeTypeValues = ['Trailer','Fixed','HR','Corporate','Job Site']
            self.ComboBoxOfficeType['state'] = 'disabled'
            if (self.FacilityTableOfficeTypeArray[curItem] == 'Trailer'):
                self.ComboBoxOfficeType.current(0)
            if (self.FacilityTableOfficeTypeArray[curItem] == 'Fixed'):
                self.ComboBoxOfficeType.current(1)
            if (self.FacilityTableOfficeTypeArray[curItem] == 'HR'):
                self.ComboBoxOfficeType.current(2)
            if (self.FacilityTableOfficeTypeArray[curItem] == 'Corporate'):
                self.ComboBoxOfficeType.current(3)
            if (self.FacilityTableOfficeTypeArray[curItem] == 'Job Site'):
                self.ComboBoxOfficeType.current(4)


            #self.OfficeSizeValues = ['Small','Medium','Large','Corporate','Datacenter']
            self.ComboBoxOfficeSize['state'] = 'disabled'
            if (self.FacilityTableOfficeSizeArray[curItem] == 'Small'):
                self.ComboBoxOfficeSize.current(0)
            if (self.FacilityTableOfficeSizeArray[curItem] == 'Medium'):
                self.ComboBoxOfficeSize.current(1)
            if (self.FacilityTableOfficeSizeArray[curItem] == 'Large'):
                self.ComboBoxOfficeSize.current(2)
            if (self.FacilityTableOfficeSizeArray[curItem] == 'Corporate'):
                self.ComboBoxOfficeSize.current(3)
            if (self.FacilityTableOfficeSizeArray[curItem] == 'Datacenter'):
                self.ComboBoxOfficeSize.current(4)

            self.FacilityLeaseEndNameFrameEntry['state'] = 'normal'
            self.FacilityLeaseEndNameFrameEntry.delete(0,END)
            self.FacilityLeaseEndNameFrameEntry.insert(0,self.FacilityTableFacilityLeaseEndNameArray[curItem])
            self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
            self.ButtonFacilityLeaseEnd['state'] = 'disabled'

            self.FacilityLeaseEndName = self.FacilityLeaseEndNameFrameEntry.get()
            #curItem = self.FacilityTreeview.focus()  
            #dic = self.FacilityTreeview.item(curItem)
            #curItem = int(dic.get('text')) - 1
            self.data = {}
            if (len(self.FacilityLeaseEndName) > 0):
                if (self.FacilityTableFacilityLeaseEndDayArray[curItem] != 0):
                    self.data['day_selected'] = self.FacilityTableFacilityLeaseEndDayArray[curItem]
                    self.data['month_selected'] = self.FacilityTableFacilityLeaseEndMonthArray[curItem]
                    self.data['year_selected'] = self.FacilityTableFacilityLeaseEndYearArray[curItem]

            self.FacilityOfficeAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeAddress1FrameEntry.insert(0,self.FacilityTableFacilityOfficeAddress1Array[curItem])
            self.FacilityOfficeAddress1FrameEntry['state'] = 'readonly'
            self.FacilityOfficeAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress2FrameEntry.delete(0,END)
            self.FacilityOfficeAddress2FrameEntry.insert(0,self.FacilityTableFacilityOfficeAddress2Array[curItem])
            self.FacilityOfficeAddress2FrameEntry['state'] = 'readonly'

            self.FacilityOfficeCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCityFrameEntry.delete(0,END)               
            self.FacilityOfficeCityFrameEntry.insert(0,self.FacilityTableFacilityOfficeCityArray[curItem])
            self.FacilityOfficeCityFrameEntry['state'] = 'readonly'

            self.FacilityOfficeStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeStateFrameEntry.delete(0,END)               
            self.FacilityOfficeStateFrameEntry.insert(0,self.FacilityTableFacilityOfficeStateArray[curItem])
            self.FacilityOfficeStateFrameEntry['state'] = 'readonly'


            self.FacilityOfficeCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCountryFrameEntry.delete(0,END)               
            self.FacilityOfficeCountryFrameEntry.insert(0,self.FacilityTableFacilityOfficeCountryArray[curItem])
            self.FacilityOfficeCountryFrameEntry['state'] = 'readonly'

            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeZIPCodeFrameEntry.delete(0,END)               
            self.FacilityOfficeZIPCodeFrameEntry.insert(0,self.FacilityTableFacilityOfficeZIPCodeArray[curItem])
            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'readonly'

            self.FacilityOfficeHeadcountFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHeadcountFrameEntry.delete(0,END)
            self.FacilityOfficeHeadcountFrameEntry.insert(0,self.FacilityTableFacilityOfficeHeadcountArray[curItem])
            self.FacilityOfficeHeadcountFrameEntry['state'] = 'readonly'

            self.FacilityOfficeHRIDFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHRIDFrameEntry.delete(0,END)
            self.FacilityOfficeHRIDFrameEntry.insert(0,self.FacilityTableFacilityOfficeHRIDArray[curItem])
            self.FacilityOfficeHRIDFrameEntry['state'] = 'readonly'

            i = 0
            while (i < len (self.OfficeStatusValues)):
                if (self.FacilityTableOfficeStatusArray[curItem] == self.OfficeStatusValues[i]):
                    self.ComboBoxOfficeStatus.current(i)
                    i = i + len (self.OfficeStatusValues)
                i = i + 1
                
            self.ComboBoxOfficeStatus['state'] = 'disabled'
                            
            self.FacilityStrategyFrameEntry['state'] = 'normal'
            self.FacilityStrategyFrameEntry.delete(0,END)
            self.FacilityStrategyFrameEntry.insert(0,self.FacilityTableStrategyArray[curItem])
            self.FacilityStrategyFrameEntry['state'] = 'readonly'

            self.FacilityActionFrameEntry['state'] = 'normal'
            self.FacilityActionFrameEntry.delete(0,END)
            self.FacilityActionFrameEntry.insert(0,self.FacilityTableActionArray[curItem])
            self.FacilityActionFrameEntry['state'] = 'readonly'

            self.FacilityNotesFrameEntry['state'] = 'normal'
            self.FacilityNotesFrameEntry.delete(0,END)
            self.FacilityNotesFrameEntry.insert(0,self.FacilityTableNotesArray[curItem])
            self.FacilityNotesFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeNewAddress1FrameEntry.insert(0,self.FacilityTableFacilityNewOfficeAddress1Array[curItem])
            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'readonly'
            
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress2FrameEntry.delete(0,END)
            self.FacilityOfficeNewAddress2FrameEntry.insert(0,self.FacilityTableFacilityNewOfficeAddress2Array[curItem])
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCityFrameEntry.delete(0,END)
            self.FacilityOfficeNewCityFrameEntry.insert(0,self.FacilityTableFacilityNewOfficeCityArray[curItem])
            self.FacilityOfficeNewCityFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewStateFrameEntry.delete(0,END)
            self.FacilityOfficeNewStateFrameEntry.insert(0,self.FacilityTableFacilityNewOfficeStateArray[curItem])
            self.FacilityOfficeNewStateFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCountryFrameEntry.delete(0,END)
            self.FacilityOfficeNewCountryFrameEntry.insert(0,self.FacilityTableFacilityNewOfficeCountryArray[curItem]) 
            self.FacilityOfficeNewCountryFrameEntry['state'] = 'readonly'

            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewZIPCodeFrameEntry.delete(0,END)
            self.FacilityOfficeNewZIPCodeFrameEntry.insert(0,self.FacilityTableFacilityNewOfficeZIPCodeArray[curItem])
            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'readonly'

            # CheckButton
            self.FacilityBusinessUnitPowerCheckbutton.deselect()
            self.FacilityBusinessUnitWaterCheckbutton.deselect()
            self.FacilityBusinessUnitTelecomCheckbutton.deselect()
            self.FacilityBusinessUnitSPCheckbutton.deselect()
            self.FacilityBusinessUnitCorpCheckbutton.deselect()
            self.FacilityBusinessUnitDCCheckbutton.deselect()
            self.FacilityBusinessUnitWHCheckbutton.deselect()
            
            if (self.FacilityTableBusinessUnitArray[curItem].find("power",0) != -1): 
                self.FacilityBusinessUnitPowerCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('water',0) != -1): 
                self.FacilityBusinessUnitWaterCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('telecom',0) != -1): 
                self.FacilityBusinessUnitTelecomCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('sp',0) != -1):                 
                self.FacilityBusinessUnitSPCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('corp',0) != -1):                 
                self.FacilityBusinessUnitCorpCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('dc',0) != -1):                 
                self.FacilityBusinessUnitDCCheckbutton.select()
            if (self.FacilityTableBusinessUnitArray[curItem].find('wh',0) != -1):                 
                self.FacilityBusinessUnitWHCheckbutton.select()
            
            self.FacilityBusinessUnitPowerCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitWaterCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitTelecomCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitSPCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitCorpCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitDCCheckbutton['state'] = 'disable'
            self.FacilityBusinessUnitWHCheckbutton['state'] = 'disable'


    def Enable_Screen(self,option):
        # This function is used when the ADD button is selected

        #self.FacilityBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

        self.ComboBoxCoutryID['state'] = DISABLED
        self.ComboBoxRegionID['state'] = DISABLED
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = DISABLED
            self.ButtonCountryRefresh['state'] = DISABLED
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = DISABLED
            self.ButtonRegionRefresh['state'] = DISABLED
            
        self.ButtonFacilityAdd['state'] = DISABLED
        self.ButtonFacilityEdit['state'] = DISABLED
        self.ButtonFacilityRemove['state'] = DISABLED
        self.ButtonFacilityOK['state'] = ACTIVE
        self.ButtonFacilityCancel['state'] = ACTIVE
        self.ButtonFacilityValidate['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        if (option == 'add'):
            self.FacilityIDFrameEntry['state'] = 'normal'
            self.FacilityIDFrameEntry.delete(0,END)

            self.FacilityNameFrameEntry['state'] = 'normal'
            self.FacilityNameFrameEntry.delete(0,END)
                        
            self.ComboBoxOfficeType['state'] = 'readonly'
            self.ComboBoxOfficeSize['state'] = 'readonly'

            self.FacilityLeaseEndNameFrameEntry['state'] = 'normal'
            self.FacilityLeaseEndNameFrameEntry.delete(0,END)
            self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
            self.ButtonFacilityLeaseEnd['state'] = 'active'

            self.FacilityOfficeAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress2FrameEntry.delete(0,END)

            self.FacilityOfficeCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCityFrameEntry.delete(0,END)               

            self.FacilityOfficeStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeStateFrameEntry.delete(0,END)               

            self.FacilityOfficeCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCountryFrameEntry.delete(0,END)               

            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeZIPCodeFrameEntry.delete(0,END)               

            self.FacilityOfficeHeadcountFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHeadcountFrameEntry.delete(0,END)               

            self.FacilityOfficeHRIDFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHRIDFrameEntry.delete(0,END)               

            self.ComboBoxOfficeStatus['state'] = 'readonly'

            self.FacilityStrategyFrameEntry['state'] = 'normal'
            self.FacilityStrategyFrameEntry.delete(0,END)

            self.FacilityActionFrameEntry['state'] = 'normal'
            self.FacilityActionFrameEntry.delete(0,END)

            self.FacilityNotesFrameEntry['state'] = 'normal'
            self.FacilityNotesFrameEntry.delete(0,END)

            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress1FrameEntry.delete(0,END)
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress2FrameEntry.delete(0,END)

            self.FacilityOfficeNewCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCityFrameEntry.delete(0,END)               

            self.FacilityOfficeNewStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewStateFrameEntry.delete(0,END)               

            self.FacilityOfficeNewCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCountryFrameEntry.delete(0,END)               

            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewZIPCodeFrameEntry.delete(0,END)               

            # CheckButton
                  
            self.FacilityBusinessUnitPowerCheckbutton.deselect()
            self.FacilityBusinessUnitWaterCheckbutton.deselect()
            self.FacilityBusinessUnitTelecomCheckbutton.deselect()
            self.FacilityBusinessUnitSPCheckbutton.deselect()
            self.FacilityBusinessUnitCorpCheckbutton.deselect()
            self.FacilityBusinessUnitDCCheckbutton.deselect()
            self.FacilityBusinessUnitWHCheckbutton.deselect()

        if (option == 'edit'):
            self.FacilityIDFrameEntry['state'] = 'readonly'
            self.FacilityNameFrameEntry['state'] = 'normal'
                
            self.ComboBoxOfficeType['state'] = 'readonly'
            self.ComboBoxOfficeSize['state'] = 'readonly'

            self.FacilityLeaseEndNameFrameEntry['state'] = 'normal'
            self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
            self.ButtonFacilityLeaseEnd['state'] = 'active'

            self.FacilityOfficeAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeCityFrameEntry['state'] = 'normal'                      
            self.FacilityOfficeStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeZIPCodeFrameEntry['state'] = 'normal'               

            self.FacilityOfficeHeadcountFrameEntry['state'] = 'normal'               
            self.FacilityOfficeHRIDFrameEntry['state'] = 'normal'                              

            self.ComboBoxOfficeStatus['state'] = 'readonly'

            self.FacilityStrategyFrameEntry['state'] = 'normal'
            self.FacilityActionFrameEntry['state'] = 'normal'
            self.FacilityNotesFrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress1FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewAddress2FrameEntry['state'] = 'normal'
            self.FacilityOfficeNewCityFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewStateFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewCountryFrameEntry['state'] = 'normal'               
            self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'normal'               

        # CheckButton          
        self.FacilityBusinessUnitPowerCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitWaterCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitTelecomCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitSPCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitCorpCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitDCCheckbutton['state'] = 'normal'
        self.FacilityBusinessUnitWHCheckbutton['state'] = 'normal'

        
    def Disable_Screen(self):
        # This function is used when the entry was added.modified to the Database

        #self.FacilityBusinessUnitPowerCheckbutton.select()
        #print (self.varpower.get())

        self.ComboBoxCoutryID['state'] = 'readonly'
        self.ComboBoxRegionID['state'] = 'readonly'
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = ACTIVE
            self.ButtonCountryRefresh['state'] = ACTIVE
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = ACTIVE
            self.ButtonRegionRefresh['state'] = ACTIVE
            
        self.ButtonFacilityAdd['state'] = ACTIVE
        self.ButtonFacilityEdit['state'] = DISABLED
        self.ButtonFacilityRemove['state'] = DISABLED
        self.ButtonFacilityOK['state'] = DISABLED
        self.ButtonFacilityCancel['state'] = ACTIVE
        self.ButtonFacilityValidate['state'] = DISABLED

        # Create Progress Bar
        self.progress['maximum'] = 100
        self.progress['value'] = 0

        # Setup Labels and Entry
        self.FacilityIDFrameEntry['state'] = 'readonly'
        self.FacilityNameFrameEntry['state'] = 'readonly'
                    
        self.ComboBoxOfficeType['state'] = 'disable'
        self.ComboBoxOfficeSize['state'] = 'disable'

        self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
        self.ButtonFacilityLeaseEnd['state'] = 'disabled'
        self.FacilityOfficeAddress1FrameEntry['state'] = 'readonly'
        self.FacilityOfficeAddress2FrameEntry['state'] = 'readonly'
        self.FacilityOfficeCityFrameEntry['state'] = 'readonly'                       
        self.FacilityOfficeStateFrameEntry['state'] = 'readonly'        
        self.FacilityOfficeCountryFrameEntry['state'] = 'readonly'                       
        self.FacilityOfficeZIPCodeFrameEntry['state'] = 'readonly'                       
        self.FacilityOfficeHeadcountFrameEntry['state'] = 'readonly'                       
        self.FacilityOfficeHRIDFrameEntry['state'] = 'readonly'               

        self.ComboBoxOfficeStatus['state'] = 'disabled'

        self.FacilityStrategyFrameEntry['state'] = 'readonly'
        self.FacilityActionFrameEntry['state'] = 'readonly'
        self.FacilityNotesFrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewAddress1FrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewAddress2FrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewCityFrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewStateFrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewCountryFrameEntry['state'] = 'readonly'
        self.FacilityOfficeNewZIPCodeFrameEntry['state'] = 'readonly'

        # CheckButton
        self.FacilityBusinessUnitPowerCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitWaterCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitTelecomCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitSPCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitCorpCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitDCCheckbutton['state'] = 'disable'
        self.FacilityBusinessUnitWHCheckbutton['state'] = 'disable'


    def Collect_Screen(self):
        # This function is used when the ADD button is selected

        self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
        self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
        self.FacilityID = self.FacilityIDFrameEntry.get()
        self.FacilityName = self.FacilityNameFrameEntry.get()
                    
        self.OfficeType = self.OfficeTypeValues[self.ComboBoxOfficeType.current()]
        self.OfficeSize = self.OfficeSizeValues[self.ComboBoxOfficeSize.current()]
        self.FacilityLeaseEndName = self.FacilityLeaseEndNameFrameEntry.get()
        if (len(self.FacilityLeaseEndName) > 0):
            self.date = str(self.data['month_selected']) + '/' + str(self.data['day_selected']) + '/' + str(self.data['year_selected'])
            self.FacilityLeaseEndMonth = self.data['month_selected']
            self.FacilityLeaseEndDay = self.data['day_selected']
            self.FacilityLeaseEndYear = self.data['year_selected']
        else:
            self.date = ""
            self.FacilityLeaseEndMonth = "0"
            self.FacilityLeaseEndDay = "0"
            self.FacilityLeaseEndYear = "0"
        self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1FrameEntry.get()
        self.FacilityOfficeAddress2 = self.FacilityOfficeAddress2FrameEntry.get()
        self.FacilityOfficeCity = self.FacilityOfficeCityFrameEntry.get()             
        self.FacilityOfficeState = self.FacilityOfficeStateFrameEntry.get()            
        self.FacilityOfficeCountry = self.FacilityOfficeCountryFrameEntry.get()             
        self.FacilityOfficeZIPCode = self.FacilityOfficeZIPCodeFrameEntry.get()
        self.FacilityOfficeHeadcount = self.FacilityOfficeHeadcountFrameEntry.get()
        if (len(self.FacilityOfficeHeadcount) <= 0):
            self.FacilityOfficeHeadcount = "0"
        else:
            try:
                x = int(self.FacilityOfficeHeadcount)
            except ValueError:
                self.FacilityOfficeHeadcount = "0"
        self.FacilityOfficeHRID = self.FacilityOfficeHRIDFrameEntry.get()             
        self.OfficeStatus = self.OfficeStatusValues[self.ComboBoxOfficeStatus.current()]
                
        self.FacilityBusinessUnit = ""
        if (self.varpower.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'power,'
        if (self.varwater.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'water,'
        if (self.vartelecom.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'telecom,'
        if (self.varSP.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'sp,'
        if (self.varCorp.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'corp,'
        if (self.varDC.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'dc,'
        if (self.varWH.get() == 1):
            self.FacilityBusinessUnit = self.FacilityBusinessUnit+'wh'

        self.FacilityStrategy = self.FacilityStrategyFrameEntry.get()
        self.FacilityAction = self.FacilityActionFrameEntry.get()
        self.FacilityNotes = self.FacilityNotesFrameEntry.get()
        self.FacilityOfficeNewAddress1 = self.FacilityOfficeNewAddress1FrameEntry.get()
        self.FacilityOfficeNewAddress2 = self.FacilityOfficeNewAddress2FrameEntry.get()
        self.FacilityOfficeNewCity = self.FacilityOfficeNewCityFrameEntry.get()               
        self.FacilityOfficeNewState = self.FacilityOfficeNewStateFrameEntry.get()               
        self.FacilityOfficeNewCountry = self.FacilityOfficeNewCountryFrameEntry.get()               
        self.FacilityOfficeNewZIPCode = self.FacilityOfficeNewZIPCodeFrameEntry.get()               


    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        self.Clean_Screen('country','country-combo')
        if self.db.Connect():
            # SQL Querry to the Facility Table
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
                else:
                    self.ComboBoxRegionID['state'] = 'readonly'
                    self.ComboBoxRegionID.set("")
                '''
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [self.db.results[i][1],
                            self.db.results[i][2],
                            self.db.results[i][3],
                            self.db.results[i][4]]
                    self.FacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonFacilityAdd['state']    = ACTIVE
                self.ButtonFacilityEdit['state']   = DISABLED
                self.ButtonFacilityRemove['state'] = DISABLED
                self.ButtonFacilityOK['state']     = DISABLED
                self.ButtonFacilityCancel['state'] = DISABLED
                self.ButtonFacilityValidate['state'] = DISABLED
                '''
            else:
                mbox.showerror(master=self.FacilityFrame,title='Facility',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDArray = []
            self.CountryNameArray = []               

            # SQL Querry to the Facility Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Country
                        WHERE Country_ID = '%s'
                      """ % (self.CountryID_Pre)
            else:
                sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            #sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                self.sql_querry = True
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
                mbox.showerror(master=self.FacilityFrame,title='Facility',
                message = 'No Country Records found')
                self.sql_querry = False
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        

#**************************************************************************************
#---------------------------- COUNTRY SELECTION SECTION <END> ------------------------*
#**************************************************************************************

#***************************************************************************************
#---------------------------- REGION SELECTION SECTION <BEGIN> ------------------------*
#***************************************************************************************

    def Find_Devices_and_Circuits(self,Location):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','VALIDATE DEVICES AND CIRCUITS Button']    
            Logging.Log(Parameter)
        self.Device_Exist = "NO"
        self.Circuit_Exist = "NO"
        sql = """
                SELECT * FROM Devices
                WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                """ % (Location[0],Location[1],Location[2])
        if (self.db2.Execute(sql)):
            self.Device_Exist = "YES"
                
        sql = """
                SELECT * FROM Circuits
                WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s' 
                """ % (Location[0],Location[1],Location[2])
        if (self.db2.Execute(sql)):
            self.Circuit_Exist = "YES"
        
    def Display_Region_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','Region Window']    
            Logging.Log(Parameter)
        Region.Display_Region_Window()

    def on_region_combo_changed(self,event):
        #print (self.ComboBoxCoutryID.current())
        #print (self.CountryIDArray[self.ComboBoxCoutryID.current()])
        #print (self.CountryNameArray[self.ComboBoxCoutryID.current()])               
        if (self.db.Connect() and self.db2.Connect()):
            # SQL Querry to the Facility Table

            if (self.Selection == 'cancel_edit'):
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                      """ % (self.CountryID_Pre,self.RegionID_Pre,self.FacilityID_Pre)
            else:
                sql = """
                        SELECT * FROM Facility
                        WHERE Country_ID = '%s' AND Region_ID = '%s'
                        ORDER BY Status desc, Facility_Name ASC
                      """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],
                             self.RegionIDArray[self.ComboBoxRegionID.current()])

            '''
            sql = """
                    SELECT * FROM Facility
                    WHERE Country_ID = '%s' AND Region_ID = '%s'
                    ORDER BY Facility_Name ASC
                  """ % (self.CountryIDArray[self.ComboBoxCoutryID.current()],self.RegionIDArray[self.ComboBoxRegionID.current()])
            '''
            if (self.db.Execute(sql)):
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.FacilityTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.FacilityTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
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
                self.progress['maximum'] = len(self.db.results)
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
                    self.FacilityHasDevices.append("?")
                    self.FacilityHasCircuits.append("?")
                    self.progress['value'] = i
                    i = i + 1
                self.progress['value'] =  len(self.db.results)
                i = 0
                while (i < len(self.FacilityTablePriaryKeyArray)):
                    num = i + 1
                    if (self.FacilityHasDevices[i] == "YES" and self.FacilityHasCircuits[i] == "YES"):
                        tags = "YES"
                    else:
                        tags = "NO"
                    tags = self.FacilityTableOfficeStatusArray[i].replace(" ","_")
                    item = [
                            #self.FacilityTablePriaryKeyArray[i],
                            #self.FacilityTableCountryIDArray[i], 
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
                            self.FacilityHasDevices[i],
                            self.FacilityHasCircuits[i],
                            self.FacilityTableExecutedByUserArray[i]]
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
                    
                    '''
                                0  Country_ID_Region_ID_Facility_ID    CHAR(65) NOT NULL PRIMARY KEY,
                                1  Country_ID                          CHAR(20) NOT NULL,
                                2  Region_ID                           CHAR(20) NOT NULL,
                                3  Facility_ID                         CHAR(20) NOT NULL,
                                4  Facility_Name                       CHAR(50) NOT NULL,
                                5  Business_Unit_ID                    CHAR(50) NOT NULL,
                                6  Office_Type                         CHAR(20) NOT NULL,
                                7  Office_Size                         CHAR(20) NOT NULL,
                                8  Lease_End_Date                      CHAR(20),
                                9  Lease_End_Month                     INT,
                                10 Lease_End_Day                       INT,
                                11 Lease_End_Year                      INT,
                                12 Address1                            CHAR(100),
                                13 Address2                            CHAR(100),
                                14 City                                CHAR(50),
                                15 State                               CHAR(50),
                                16 Zip_Code                            CHAR(20),
                                17 Country                             CHAR(50),
                                18 Headcount                           INT,
                                19 HR_ID_Location                      CHAR(20),
                                20 HR_Office_ID                        CHAR(20),
                                21 Status                              CHAR(20),
                                22 Executed_by_UserID                  CHAR(20))"""

                                self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                                self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
                                self.FacilityID = self.FacilityIDFrameEntry.get()
                                self.FacilityName = self.FacilityNameFrameEntry.get()
                                self.OfficeType = self.OfficeTypeValues[self.ComboBoxOfficeType.current()]
                                self.OfficeSize = self.OfficeSizeValues[self.ComboBoxOfficeSize.current()]
                                self.FacilityLeaseEndName = self.FacilityLeaseEndNameFrameEntry.get()
                                self.date = str(self.data['month_selected']) + '/' + str(self.data['day_selected']) + '/' + str(self.data['year_selected'])
                                self.FacilityLeaseEndMonth = self.data['month_selected']
                                self.FacilityLeaseEndDay = self.data['day_selected']
                                self.FacilityLeaseEndYear = self.data['year_selected']
                                self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1FrameEntry.get()
                                self.FacilityOfficeAddress2 = self.FacilityOfficeAddress2FrameEntry.get()
                                self.FacilityOfficeCity = self.FacilityOfficeCityFrameEntry.get()             
                                self.FacilityOfficeState = self.FacilityOfficeStateFrameEntry.get()            
                                self.FacilityOfficeCountry = self.FacilityOfficeCountryFrameEntry.get()             
                                self.FacilityOfficeZIPCode = self.FacilityOfficeZIPCodeFrameEntry.get()
                                self.FacilityOfficeHeadcount = self.FacilityOfficeHeadcountFrameEntry.get()            
                                self.FacilityOfficeHRID = self.FacilityOfficeHRIDFrameEntry.get()             
                                self.OfficeStatus = self.OfficeStatusValues[self.ComboBoxOfficeStatus.current()]                                
                                self.FacilityBusinessUnit = ""
                    '''
                    
                    self.FacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonFacilityAdd['state']    = ACTIVE
                self.ButtonFacilityEdit['state']   = DISABLED
                self.ButtonFacilityRemove['state'] = DISABLED
                self.ButtonFacilityOK['state']     = DISABLED
                self.ButtonFacilityCancel['state'] = DISABLED
                self.ButtonFacilityValidate['state'] = DISABLED
            else:
                mbox.showerror(master=self.FacilityFrame,title='Facility',
                message = 'No Records found')
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.FacilityTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.FacilityTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                self.ButtonFacilityAdd['state']    = ACTIVE
                self.ButtonFacilityEdit['state']   = DISABLED
                self.ButtonFacilityRemove['state'] = DISABLED
                self.ButtonFacilityOK['state']     = DISABLED
                self.ButtonFacilityCancel['state'] = DISABLED
                self.ButtonFacilityValidate['state'] = DISABLED

                self.FacilityIDFrameEntry['state'] = 'normal'
                self.FacilityIDFrameEntry.delete(0,END)
                self.FacilityIDFrameEntry['state'] = 'readonly'
                
                self.FacilityNameFrameEntry['state'] = 'normal'
                self.FacilityNameFrameEntry.delete(0,END)
                self.FacilityNameFrameEntry['state'] = 'readonly'
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    def on_Region_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','Region Refresh']    
            Logging.Log(Parameter)
        self.Clean_Screen('region','all')
        if self.db.Connect():
            # SQL Querry to the Facility Table
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
                else:
                    self.ComboBoxRegionID['state'] = 'readonly'
                    self.ComboBoxRegionID.set("")
                    if (Is_Region_Available):
                        self.ButtonRegionAdd['state'] = 'active'
                        self.ButtonRegionRefresh['state'] = 'active'
            else:
                mbox.showerror(master=self.FacilityFrame,title='Facility',
                message = 'No Records found')
                self.sql_querry = False
                self.ButtonFacilityAdd['state']    = DISABLED
                self.ButtonFacilityEdit['state']   = DISABLED
                self.ButtonFacilityRemove['state'] = DISABLED
                self.ButtonFacilityOK['state']     = DISABLED
                self.ButtonFacilityCancel['state'] = DISABLED
                self.ButtonFacilityValidate['state'] = DISABLED

                self.FacilityIDFrameEntry['state'] = 'normal'
                self.FacilityIDFrameEntry.delete(0,END)
                self.FacilityIDFrameEntry['state'] = 'readonly'
                
                self.FacilityNameFrameEntry['state'] = 'normal'
                self.FacilityNameFrameEntry.delete(0,END)
                self.FacilityNameFrameEntry['state'] = 'readonly'
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
    
#*************************************************************************************
#---------------------------- REGION SELECTION SECTION <END> ------------------------*
#*************************************************************************************
        
    def on_FacilityWindow_quit(self):
        if (self.FacilityWindowExist):
            self.FacilityWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Facility','CLOSE Window']    
                Logging.Log(Parameter)
            self.FacilityWindow.destroy()

    def Call_Button_Facility_Validate(self):

        #------------------------------- Deleting Tree View <BEGIN> --------
        x = self.FacilityTreeview.get_children()
        if x != '()': # checks if there is something in the first row
            for child in x:
                #print (child)
                self.FacilityTreeview.delete(child)
        #------------------------------- Deleting Tree View <END> --------        
        i = 0
        while (i < len(self.FacilityTablePriaryKeyArray)):
            Location = []
            Location = [self.FacilityTableCountryIDArray[i],self.FacilityTableRegionIDArray[i],self.FacilityTableFacilityIDArray[i]]
            self.Find_Devices_and_Circuits(Location)
            self.FacilityHasDevices[i] = self.Device_Exist
            self.FacilityHasCircuits[i] = self.Circuit_Exist
            num = i + 1
            if (self.FacilityHasDevices[i] == "YES" and self.FacilityHasCircuits[i] == "YES"):
                tags = "YES"
            else:
                tags = "NO"
            item = [
                    #self.FacilityTablePriaryKeyArray[i],
                    #self.FacilityTableCountryIDArray[i], 
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
                    self.FacilityHasDevices[i],
                    self.FacilityHasCircuits[i],
                    self.FacilityTableExecutedByUserArray[i]]
                    
            self.FacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
            i = i + 1
        self.ButtonFacilityAdd['state']    = ACTIVE
        self.ButtonFacilityEdit['state']   = DISABLED
        self.ButtonFacilityRemove['state'] = DISABLED
        self.ButtonFacilityOK['state']     = DISABLED
        self.ButtonFacilityCancel['state'] = DISABLED
        self.ButtonFacilityValidate['state'] = DISABLED

    def on_Facility_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','Facility Refresh']    
            Logging.Log(Parameter)
        self.on_country_combo_changed("event")
       
    def Call_Button_Facility_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','ADD Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.Enable_Screen('add')
        self.Selection = 'add'

    def Call_Button_Facility_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        self.Enable_Screen('edit')

    def Call_Button_Facility_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
            self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
            self.FacilityID = self.FacilityIDFrameEntry.get()
            PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID)
            if (mbox.askyesnocancel(master=self.FacilityFrame,title='Facility',message = 'Are you Sure you want to Remove it?')):
                #PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID)
                #print (PrimaryKey)
                if (self.FacilityID == 'UNKNOWN'):
                    mbox.showerror(master=self.FacilityFrame,title='Facility',
                        message = '*** The Facility ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Facility --------
                    sql = """
                            SELECT * FROM SITES
                            WHERE Country_ID = '%s' AND Region_ID = '%s' AND Facility_ID = '%s'
                          """ % (self.CountryID,self.RegionID,self.FacilityID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.FacilityFrame,title='Facility',
                                    message = '*** The Facility ID you entered CANNOT be Removed Becuase is used in Sites ***')
                            self.on_region_combo_changed("event")
                            self.Disable_Screen()
                    else:
                        sql = """
                                SELECT * FROM Facility
                                WHERE Country_ID_Region_ID_Facility_ID = '%s'
                              """ % (PrimaryKey)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Facility WHERE Country_ID_Region_ID_Facility_ID = '%s'" % (PrimaryKey)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.FacilityFrame,title='Facility',
                                        message = '*** The Facility ID you entered was Removed ***')
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.FacilityFrame,title='Facility',
                                        message = '*** The Facility ID you entered was NOT Removed ***')               
                            self.on_region_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            mbox.showerror(master=self.FacilityFrame,title='Facility',
                            message = '*** The Facility ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Facility_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ButtonFacilityAdd['state']    = DISABLED
            self.ButtonFacilityEdit['state']   = DISABLED
            self.ButtonFacilityRemove['state'] = DISABLED
            self.ButtonFacilityOK['state']     = ACTIVE
            self.ButtonFacilityCancel['state'] = ACTIVE
            self.ButtonFacilityValidate['state'] = DISABLED
            self.Collect_Screen()
            #CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
            #RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
            #FacilityID = self.FacilityIDFrameEntry.get()
            PrimaryKey = (self.CountryID+"-"+self.RegionID+"-"+self.FacilityID)
            #print (PrimaryKey+"....")
            #self.Selection = 'notyet'
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(self.FacilityID) > 0) and (len(self.FacilityName) > 0)):
                    sql = """
                            SELECT * FROM Facility
                            WHERE Country_ID_Region_ID_Facility_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.FacilityFrame,title='Facility',
                        message = '*** The Facility ID you entered already exist ***')
                    else:
                        self.FacilityHROfficeID = "-"
                        if (len(self.FacilityOfficeAddress1) > 100):
                            self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1[:100]
                        if (len(self.FacilityOfficeAddress2) > 100):
                            self.FacilityOfficeAddress2 = self.FacilityOfficeAddress2[:100]
                        sql = "INSERT INTO Facility(Country_ID_Region_ID_Facility_ID, Country_ID, Region_ID, Facility_ID, Facility_Name, \
                                   Business_Unit_ID, Office_Type, Office_Size, Lease_End_Date, Lease_End_Month, \
                                   Lease_End_Day, Lease_End_Year, Address1, Address2, City, State, Zip_Code, Country, Headcount, \
                                   HR_ID_Location, HR_Office_ID, Status, Executed_by_UserID, Strategy, Action, Notes, New_Address1,  New_Address2, \
                                   New_City, New_State, New_Zip_Code, New_Country) \
                                   VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%d','%d','%d','%s','%s','%s','%s','%s','%s','%d','%s', \
                                   '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" % (PrimaryKey,self.CountryID,self.RegionID,self.FacilityID,self.FacilityName,self.FacilityBusinessUnit,
                                    self.OfficeType,self.OfficeSize,self.date,int(self.FacilityLeaseEndMonth), int(self.FacilityLeaseEndDay),
                                    int(self.FacilityLeaseEndYear),self.FacilityOfficeAddress1,self.FacilityOfficeAddress2,self.FacilityOfficeCity,
                                    self.FacilityOfficeState,self.FacilityOfficeZIPCode,self.FacilityOfficeCountry,int(self.FacilityOfficeHeadcount),
                                    self.FacilityOfficeHRID,self.FacilityHROfficeID,self.OfficeStatus,self.Username,self.FacilityStrategy,self.FacilityAction,
                                    self.FacilityNotes,self.FacilityOfficeNewAddress1,self.FacilityOfficeNewAddress2,self.FacilityOfficeNewCity,
                                    self.FacilityOfficeNewState,self.FacilityOfficeNewZIPCode,self.FacilityOfficeNewCountry)
                        '''
                                0  Country_ID_Region_ID_Facility_ID    CHAR(65) NOT NULL PRIMARY KEY,
                                1  Country_ID                          CHAR(20) NOT NULL,
                                2  Region_ID                           CHAR(20) NOT NULL,
                                3  Facility_ID                         CHAR(20) NOT NULL,
                                4  Facility_Name                       CHAR(50) NOT NULL,
                                5  Business_Unit_ID                    CHAR(50) NOT NULL,
                                6  Office_Type                         CHAR(20) NOT NULL,
                                7  Office_Size                         CHAR(20) NOT NULL,
                                8  Lease_End_Date                      CHAR(20),
                                9  Lease_End_Month                     INT,
                                10 Lease_End_Day                       INT,
                                11 Lease_End_Year                      INT,
                                12 Address1                            CHAR(100),
                                13 Address2                            CHAR(100),
                                14 City                                CHAR(50),
                                15 State                               CHAR(50),
                                16 Zip_Code                            CHAR(20),
                                17 Country                             CHAR(50),
                                18 Headcount                           INT,
                                19 HR_ID_Location                      CHAR(20),
                                20 HR_Office_ID                        CHAR(20),
                                21 Status                              CHAR(20),
                                22 Executed_by_UserID                  CHAR(20))"""
                                23 Strategy                            CHAR(250),
                                24 Action                              CHAR(250),
                                25 Notes                               CHAR(250),
                                26 New_Address1                        CHAR(100),
                                27 New_Address2                        CHAR(100),
                                28 New_City                            CHAR(50),
                                29 New_State                           CHAR(50),
                                30 New_Zip_Code                        CHAR(20),
                                31 New_Country                         CHAR(50))"""

                                self.CountryID = self.CountryIDArray[self.ComboBoxCoutryID.current()]
                                self.RegionID = self.RegionIDArray[self.ComboBoxRegionID.current()]
                                self.FacilityID = self.FacilityIDFrameEntry.get()
                                self.FacilityName = self.FacilityNameFrameEntry.get()
                                self.OfficeType = self.OfficeTypeValues[self.ComboBoxOfficeType.current()]
                                self.OfficeSize = self.OfficeSizeValues[self.ComboBoxOfficeSize.current()]
                                self.FacilityLeaseEndName = self.FacilityLeaseEndNameFrameEntry.get()
                                self.date = str(self.data['month_selected']) + '/' + str(self.data['day_selected']) + '/' + str(self.data['year_selected'])
                                self.FacilityLeaseEndMonth = self.data['month_selected']
                                self.FacilityLeaseEndDay = self.data['day_selected']
                                self.FacilityLeaseEndYear = self.data['year_selected']
                                self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1FrameEntry.get()
                                self.FacilityOfficeAddress2 = self.FacilityOfficeAddress2FrameEntry.get()
                                self.FacilityOfficeCity = self.FacilityOfficeCityFrameEntry.get()             
                                self.FacilityOfficeState = self.FacilityOfficeStateFrameEntry.get()            
                                self.FacilityOfficeCountry = self.FacilityOfficeCountryFrameEntry.get()             
                                self.FacilityOfficeZIPCode = self.FacilityOfficeZIPCodeFrameEntry.get()
                                self.FacilityOfficeHeadcount = self.FacilityOfficeHeadcountFrameEntry.get()            
                                self.FacilityOfficeHRID = self.FacilityOfficeHRIDFrameEntry.get()             
                                self.OfficeStatus = self.OfficeStatusValues[self.ComboBoxOfficeStatus.current()]                                
                                self.FacilityBusinessUnit = ""
                        '''
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.FacilityFrame,title='Facility',
                                message = '*** The Facility ID you entered was Added ***')
                            self.on_region_combo_changed("event")
                            self.Disable_Screen()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.FacilityFrame,title='Facility',
                                message = '*** The Facility ID you entered was NOT Added ***')
                            self.on_region_combo_changed("event")
                            self.Disable_Screen()
                else:
                     mbox.showerror(master=self.FacilityFrame,title='Facility',
                        message = '*** The Facility ID and Facility Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                if (len(self.FacilityName) == 0):
                    mbox.showerror(master=self.FacilityFrame,title='Facility',
                    message = '*** The Facility Name Cannot be BLANK ***')
                else:
                    '''
                    sql = """
                            SELECT * FROM Facility
                            WHERE Country_ID_Region_ID_Facility_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.FacilityFrame,title='Facility',
                        message = '*** The Facility ID you entered already exist ***')
                    else:
                        self.FacilityHROfficeID = "-"
                        if (len(self.FacilityOfficeAddress1) > 100):
                            self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1[:100]
                        if (len(self.FacilityOfficeAddress2) > 100):
                            self.FacilityOfficeAddress1 = self.FacilityOfficeAddress2[:100]
                        sql = "INSERT INTO Facility(Country_ID_Region_ID_Facility_ID, Country_ID, Region_ID, Facility_ID, Facility_Name, \
                                   Business_Unit_ID, Office_Type, Office_Size, Lease_End_Date, Lease_End_Month, \
                                   Lease_End_Day, Lease_End_Year, Address1, Address2, City, State, Zip_Code, Country, Headcount, \
                                   HR_ID_Location, HR_Office_ID, Status, Executed_by_UserID) \
                                   VALUES ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%d','%d','%d','%s','%s','%s','%s','%s','%s','%d','%s', \
                                   '%s','%s','%s')" % (PrimaryKey,self.CountryID,self.RegionID,self.FacilityID,self.FacilityName,self.FacilityBusinessUnit,
                                    self.OfficeType,self.OfficeSize,self.date,int(self.FacilityLeaseEndMonth), int(self.FacilityLeaseEndDay),
                                    int(self.FacilityLeaseEndYear),self.FacilityOfficeAddress1,self.FacilityOfficeAddress2,self.FacilityOfficeCity,
                                    self.FacilityOfficeState,self.FacilityOfficeZIPCode,self.FacilityOfficeCountry,int(self.FacilityOfficeHeadcount),
                                    self.FacilityOfficeHRID,self.FacilityHROfficeID,self.OfficeStatus,self.Username)

                    '''                    
                    sql = """
                            SELECT * FROM Facility
                            WHERE Country_ID_Region_ID_Facility_ID = '%s'
                          """ % (PrimaryKey)                   
                    if (self.db.Execute(sql)):   
                        #sql = "UPDATE Facility SET Facility_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID_Facility_ID = '%s'" % (FacilityName,self.Username,PrimaryKey)
                        self.FacilityHROfficeID = "-"
                        if (len(self.FacilityOfficeAddress1) > 100):
                            self.FacilityOfficeAddress1 = self.FacilityOfficeAddress1[:100]
                        if (len(self.FacilityOfficeAddress2) > 100):
                            self.FacilityOfficeAddress2 = self.FacilityOfficeAddress2[:100]
                        sql = "UPDATE Facility SET Facility_Name  = '%s', Business_Unit_ID = '%s', Office_Type = '%s', Office_Size = '%s', \
                                   Lease_End_Date = '%s', Lease_End_Month = '%d', Lease_End_Day = '%d', Lease_End_Year = '%d', Address1 = '%s', \
                                   Address2 = '%s', City = '%s', State = '%s', Zip_Code = '%s', Country = '%s', Headcount = '%d', HR_ID_Location = '%s', \
                                   HR_Office_ID = '%s', Status = '%s', Executed_by_UserID = '%s', Strategy = '%s', Action = '%s', Notes = '%s', New_Address1 = '%s', \
                                   New_Address2 = '%s', New_City = '%s', New_State = '%s', New_Zip_Code = '%s', New_Country = '%s' \
                                   WHERE Country_ID_Region_ID_Facility_ID = '%s'" %(self.FacilityName,
                                    self.FacilityBusinessUnit,self.OfficeType,self.OfficeSize,self.date,int(self.FacilityLeaseEndMonth),
                                    int(self.FacilityLeaseEndDay),int(self.FacilityLeaseEndYear),self.FacilityOfficeAddress1,self.FacilityOfficeAddress2,
                                    self.FacilityOfficeCity,self.FacilityOfficeState,self.FacilityOfficeZIPCode,self.FacilityOfficeCountry,
                                    int(self.FacilityOfficeHeadcount),self.FacilityOfficeHRID,self.FacilityHROfficeID,self.OfficeStatus,self.Username,
                                    self.FacilityStrategy,self.FacilityAction,self.FacilityNotes,self.FacilityOfficeNewAddress1,self.FacilityOfficeNewAddress2,
                                    self.FacilityOfficeNewCity,self.FacilityOfficeNewState,self.FacilityOfficeNewZIPCode,self.FacilityOfficeNewCountry,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.FacilityFrame,title='Facility',
                                message = '*** The Facility ID you entered was Updated ***')
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.FacilityFrame,title='Facility',
                                message = '*** The Facility ID you entered was NOT Upadted ***')
                        self.on_region_combo_changed("event")
                        self.Disable_Screen()
                    else:
                        mbox.showerror(master=self.FacilityFrame,title='Facility',
                        message = '*** The Facility ID you try to Edit Does not exist Anymore ***')
                        self.on_region_combo_changed("event")
                        self.Disable_Screen()                        
        else:
            mbox.showerror(master=self.FacilityFrame,title='Facility',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Facility_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Facility','CANCEL Button']    
            Logging.Log(Parameter)
        self.Clean_Screen('facility','all')
        self.on_region_combo_changed("test")
        self.ComboBoxCoutryID['state'] = 'readonly'
        self.ComboBoxRegionID['state'] = 'readonly'
        if (Is_Country_Available):
            self.ButtonCountryAdd['state'] = ACTIVE
            self.ButtonCountryRefresh['state'] = ACTIVE
        if (Is_Region_Available):
            self.ButtonRegionAdd['state'] = ACTIVE
            self.ButtonRegionRefresh['state'] = ACTIVE
        self.Selection = 'cancel'

    def on_Facility_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.FacilityTreeview.focus()  
        dic = self.FacilityTreeview.item(curItem)
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
            if (Is_Country_Available):
                self.ButtonCountryAdd['state'] = ACTIVE
                self.ButtonCountryRefresh['state'] = ACTIVE
            if (Is_Region_Available):
                self.ButtonRegionAdd['state'] = ACTIVE
                self.ButtonRegionRefresh['state'] = ACTIVE            
            curItem = int(dic.get('text')) - 1
            self.ButtonFacilityAdd['state']    = ACTIVE
            self.ButtonFacilityEdit['state']   = ACTIVE
            self.ButtonFacilityRemove['state'] = ACTIVE
            self.ButtonFacilityOK['state']     = DISABLED
            self.ButtonFacilityCancel['state'] = DISABLED
            self.ButtonFacilityValidate['state'] = ACTIVE
            self.Display_Screen(curItem)

    def On_Facility_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_Facility_Table_Refresh()

    def on_FacilityCalendarWindow_quit(self):
        self.data = self.cal.kill_and_save()
        self.FacilityCalendarExist = False
        self.FacilityCalendarWindow.destroy()
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
        self.FacilityLeaseEndNameFrameEntry['state'] = 'normal'
        self.FacilityLeaseEndNameFrameEntry.delete(0,END)
        self.FacilityLeaseEndNameFrameEntry.insert(0,self.date)
        self.FacilityLeaseEndNameFrameEntry['state'] = 'readonly'
        
    def Call_Button_Facility_Calendar_Lease_End(self):
        if not (self.FacilityCalendarExist):
            self.FacilityCalendarExist = True
            self.FacilityCalendarWindow = Tk()
            self.FacilityCalendarWindow.title("Lease End Calendar")
            self.FacilityCalendarWindow.protocol("WM_DELETE_WINDOW", self.on_FacilityCalendarWindow_quit)
            self.FacilityCalendarWindow.call('tk', 'scaling', self.Windows_Scaling)
            self.data = {}
            if (self.Selection == 'edit'):
                self.FacilityLeaseEndName = self.FacilityLeaseEndNameFrameEntry.get()
                curItem = self.FacilityTreeview.focus()  
                dic = self.FacilityTreeview.item(curItem)
                curItem = int(dic.get('text')) - 1
                if (len(self.FacilityLeaseEndName) > 0):
                    if (self.FacilityTableFacilityLeaseEndDayArray[curItem] != 0):
                        self.data['day_selected'] = self.FacilityTableFacilityLeaseEndDayArray[curItem]
                        self.data['month_selected'] = self.FacilityTableFacilityLeaseEndMonthArray[curItem]
                        self.data['year_selected'] = self.FacilityTableFacilityLeaseEndYearArray[curItem]
            self.cal = Class_Calendar(self.FacilityCalendarWindow, self.data)
            self.cal.setup()


##############
    def showstate(self,*args):
        print (*args)
        if (self.varpower.get()):
            print ("the Click is on")
        else:
            print ("the Click is off")

#var.trace_variable("w", showstate)

##############


    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Facility List"
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
                    mbox.showinfo(master=self.FacilityFrame,title='Facility List',
                            message = '!!! The File was saved !!!')
                    self.FacilityWindow.title("Facility                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Facility','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FacilityFrame,title='Facility List',
                            message = '*** The File was not saved, Perhaps It is already open ***')




    def Display_Facility_Window(self):
        #global FacilityLabel
 
        if not self.FacilityWindowExist:
            # Set up the Window
            self.FacilityWindowExist = True
            self.FacilityWindow = Tk()
            self.FacilityWindow.geometry('1350x700+350+70')
            self.FacilityWindow.title("Facility")
            self.FacilityWindow.protocol("WM_DELETE_WINDOW", self.on_FacilityWindow_quit)
            self.FacilityWindow.call('tk', 'scaling', self.Windows_Scaling) # This is GREAT for Windows 1
            FacilityLabel = Label(self.FacilityWindow,text="Helvetica", font=("Helvetica", 19))
            FacilityLabel["text"] = "Office / Facility"
            FacilityLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FacilityWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FacilityWindow_quit)
            self.FacilityWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.FacilityFrame = Frame(self.FacilityWindow)
            self.FacilityFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FacilityFrame.rowconfigure(0, weight=1)
            self.FacilityFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.FacilityFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 7, width=75, height=25)

                        self.ButtonCountryRefresh = Button(self.FacilityFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 7, width=75, height=25)

                    if (Is_Region_Available):
                        self.ButtonRegionAdd = Button(self.FacilityFrame, text = '+ Region', command = self.Display_Region_Window, state=DISABLED)
                        self.ButtonRegionAdd.place(x = 450, y = 35, width=75, height=25)

                        self.ButtonRegionRefresh = Button(self.FacilityFrame, text = 'Refresh', command = self.on_Region_Table_Refresh, state=DISABLED)
                        self.ButtonRegionRefresh.place(x = 550, y = 35, width=75, height=25)

                    #self.ButtonFacilityLocalPointOfContacts = Button(self.FacilityFrame, text = 'Local Point Of Contacts', command = self.Call_Button_Facility_Validate, state=DISABLED)
                    #self.ButtonFacilityLocalPointOfContacts.place(x = 725, y = 4, width=150, height=25)

                    self.ButtonFacilityValidate = Button(self.FacilityFrame, text = 'Validate HW & Circuits', command = self.Call_Button_Facility_Validate, state=DISABLED)
                    self.ButtonFacilityValidate.place(x = 725, y = 30, width=150, height=25)

                    self.ButtonFacilityAdd = Button(self.FacilityFrame, text = 'Add', command = self.Call_Button_Facility_Add, state=DISABLED)
                    self.ButtonFacilityAdd.place(x = 450, y = 60, width=75, height=25)

                    self.ButtonFacilityEdit = Button(self.FacilityFrame, text = 'Edit', command = self.Call_Button_Facility_Edit, state=DISABLED)
                    self.ButtonFacilityEdit.place(x = 550, y = 60, width=75, height=25)

                    self.ButtonFacilityRemove = Button(self.FacilityFrame, text = 'Remove', command = self.Call_Button_Facility_Remove, state=DISABLED)
                    self.ButtonFacilityRemove.place(x = 650, y = 60, width=75, height=25)

                    self.ButtonFacilityOK = Button(self.FacilityFrame, text = 'OK / UPDATE', command = self.Call_Button_Facility_OK, state=DISABLED)
                    self.ButtonFacilityOK.place(x = 750, y = 60, width=100, height=25)

                    self.ButtonFacilityCancel = Button(self.FacilityFrame, text = 'Cancel', command = self.Call_Button_Facility_Cancel, state=DISABLED)
                    self.ButtonFacilityCancel.place(x = 875, y = 60, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.FacilityFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=88)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)

                    RegionIDFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region Name:"
                    RegionIDFrameLabel.place(x=10, y=35)

                    # Setup Labels and Entry
                    FacilityIDFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityIDFrameLabel["text"] = "Facility ID:"
                    FacilityIDFrameLabel.place(x=10, y=60)
                    self.FacilityIDFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityIDFrameEntry['width']=50
                    self.FacilityIDFrameEntry.place(x=110, y=60)
                    self.FacilityIDFrameEntry['state'] = DISABLED

                    
                    FacilityNameFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityNameFrameLabel["text"] = "Facility Name:"
                    FacilityNameFrameLabel.place(x=10, y=85)
                    self.FacilityNameFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityNameFrameEntry['width']=50
                    self.FacilityNameFrameEntry.place(x=110, y=85)
                    self.FacilityNameFrameEntry['state'] = DISABLED               

                    FacilityBusinessUnitFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityBusinessUnitFrameLabel["text"] = "Business Unit:"
                    FacilityBusinessUnitFrameLabel.place(x=10, y=110)

                    FacilityOfficeTypeFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeTypeFrameLabel["text"] = "Office Type:"
                    FacilityOfficeTypeFrameLabel.place(x=10, y=135)
                    self.OfficeTypeValues = ['Trailer','Fixed','HR','Corporate','Job Site']
                    #self.OfficeTypeValues = ['Trailer','Fixed','HR','Corporate']
                    self.ComboBoxOfficeType = ttk.Combobox(self.FacilityFrame, state='disabled', width = 25)
                    self.ComboBoxOfficeType['values'] = self.OfficeTypeValues
                    self.ComboBoxOfficeType.current(0)
                    self.ComboBoxOfficeType.place(x = 110, y = 137)

                    FacilityOfficeSizeFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeSizeFrameLabel["text"] = "Office Size:"
                    FacilityOfficeSizeFrameLabel.place(x=300, y=135)
                    self.OfficeSizeValues = ['Small','Medium','Large','Corporate','Datacenter']
                    self.ComboBoxOfficeSize = ttk.Combobox(self.FacilityFrame, state='disabled', width = 25)
                    self.ComboBoxOfficeSize['values'] = self.OfficeSizeValues
                    self.ComboBoxOfficeSize.current(0)
                    self.ComboBoxOfficeSize.place(x = 380, y = 137)

                    FacilityStrategyFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityStrategyFrameLabel["text"] = "Strategy:"
                    FacilityStrategyFrameLabel.place(x=580, y=137)
                    self.FacilityStrategyFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityStrategyFrameEntry['width']=50
                    self.FacilityStrategyFrameEntry.place(x=650, y=137)
                    self.FacilityStrategyFrameEntry['state'] = DISABLED               

                    FacilityActionFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityActionFrameLabel["text"] = "Action:"
                    FacilityActionFrameLabel.place(x=965, y=137)
                    self.FacilityActionFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityActionFrameEntry['width']=48
                    self.FacilityActionFrameEntry.place(x=1020, y=137)
                    self.FacilityActionFrameEntry['state'] = DISABLED               

                    FacilityNotesFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityNotesFrameLabel["text"] = "Notes:"
                    FacilityNotesFrameLabel.place(x=580, y=162)
                    self.FacilityNotesFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityNotesFrameEntry['width']=110
                    self.FacilityNotesFrameEntry.place(x=650, y=162)
                    self.FacilityNotesFrameEntry['state'] = DISABLED               

                    FacilityOfficeLeaseFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeLeaseFrameLabel["text"] = "Lease End:"
                    FacilityOfficeLeaseFrameLabel.place(x=10, y=162)
                    self.FacilityLeaseEndNameFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityLeaseEndNameFrameEntry['width']=35
                    self.FacilityLeaseEndNameFrameEntry.place(x=110, y=165)
                    self.FacilityLeaseEndNameFrameEntry['state'] = DISABLED               
                    self.ButtonFacilityLeaseEnd = Button(self.FacilityFrame, text = 'Calendar', command = self.Call_Button_Facility_Calendar_Lease_End, state=DISABLED)
                    self.ButtonFacilityLeaseEnd.place(x = 340, y = 161, width=75, height=25)

                    FacilityOfficeAddressFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeAddressFrameLabel["text"] = "Address:"
                    FacilityOfficeAddressFrameLabel.place(x=10, y=190)
                    self.FacilityOfficeAddress1FrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeAddress1FrameEntry['width']=60
                    self.FacilityOfficeAddress1FrameEntry.place(x=110, y=190)
                    self.FacilityOfficeAddress1FrameEntry['state'] = DISABLED               
                    self.FacilityOfficeAddress2FrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeAddress2FrameEntry['width']=60
                    self.FacilityOfficeAddress2FrameEntry.place(x=110, y=210)
                    self.FacilityOfficeAddress2FrameEntry['state'] = DISABLED               

                    FacilityOfficeNewAddressFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeNewAddressFrameLabel["text"] = "New Address:"
                    FacilityOfficeNewAddressFrameLabel.place(x=550, y=190)
                    self.FacilityOfficeNewAddress1FrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewAddress1FrameEntry['width']=60
                    self.FacilityOfficeNewAddress1FrameEntry.place(x=650, y=190)
                    self.FacilityOfficeNewAddress1FrameEntry['state'] = DISABLED               
                    self.FacilityOfficeNewAddress2FrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewAddress2FrameEntry['width']=60
                    self.FacilityOfficeNewAddress2FrameEntry.place(x=650, y=210)
                    self.FacilityOfficeNewAddress2FrameEntry['state'] = DISABLED               

                    FacilityOfficeCityFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeCityFrameLabel["text"] = "City:"
                    FacilityOfficeCityFrameLabel.place(x=10, y=235)
                    self.FacilityOfficeCityFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeCityFrameEntry['width']=50
                    self.FacilityOfficeCityFrameEntry.place(x=110, y=235)
                    self.FacilityOfficeCityFrameEntry['state'] = DISABLED               

                    FacilityOfficeNewCityFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeNewCityFrameLabel["text"] = "New City:"
                    FacilityOfficeNewCityFrameLabel.place(x=550, y=235)
                    self.FacilityOfficeNewCityFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewCityFrameEntry['width']=50
                    self.FacilityOfficeNewCityFrameEntry.place(x=650, y=235)
                    self.FacilityOfficeNewCityFrameEntry['state'] = DISABLED               

                    FacilityOfficeStateFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeStateFrameLabel["text"] = "State:"
                    FacilityOfficeStateFrameLabel.place(x=10, y=260)
                    self.FacilityOfficeStateFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeStateFrameEntry['width']=40
                    self.FacilityOfficeStateFrameEntry.place(x=110, y=260)
                    self.FacilityOfficeStateFrameEntry['state'] = DISABLED               

                    FacilityOfficeNewStateFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeNewStateFrameLabel["text"] = "New State:"
                    FacilityOfficeNewStateFrameLabel.place(x=550, y=260)
                    self.FacilityOfficeNewStateFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewStateFrameEntry['width']=40
                    self.FacilityOfficeNewStateFrameEntry.place(x=650, y=260)
                    self.FacilityOfficeNewStateFrameEntry['state'] = DISABLED               

                    FacilityOfficeCountryFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeCountryFrameLabel["text"] = "Country:"
                    FacilityOfficeCountryFrameLabel.place(x=10, y=287)
                    self.FacilityOfficeCountryFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeCountryFrameEntry['width']=25
                    self.FacilityOfficeCountryFrameEntry.place(x=110, y=287)
                    self.FacilityOfficeCountryFrameEntry['state'] = DISABLED               

                    FacilityOfficeNewCountryFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeNewCountryFrameLabel["text"] = "New Country:"
                    FacilityOfficeNewCountryFrameLabel.place(x=550, y=287)
                    self.FacilityOfficeNewCountryFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewCountryFrameEntry['width']=25
                    self.FacilityOfficeNewCountryFrameEntry.place(x=650, y=287)
                    self.FacilityOfficeNewCountryFrameEntry['state'] = DISABLED               

                    FacilityOfficeZIPCodeFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeZIPCodeFrameLabel["text"] = "ZIP Code:"
                    FacilityOfficeZIPCodeFrameLabel.place(x=280, y=287)
                    self.FacilityOfficeZIPCodeFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeZIPCodeFrameEntry['width']=25
                    self.FacilityOfficeZIPCodeFrameEntry.place(x=350, y=287)
                    self.FacilityOfficeZIPCodeFrameEntry['state'] = DISABLED               

                    FacilityOfficeNewZIPCodeFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeNewZIPCodeFrameLabel["text"] = "New ZIP Code:"
                    FacilityOfficeNewZIPCodeFrameLabel.place(x=820, y=287)
                    self.FacilityOfficeNewZIPCodeFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeNewZIPCodeFrameEntry['width']=25
                    self.FacilityOfficeNewZIPCodeFrameEntry.place(x=920, y=287)
                    self.FacilityOfficeNewZIPCodeFrameEntry['state'] = DISABLED               

                    FacilityOfficeHeadcountFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeHeadcountFrameLabel["text"] = "Headcount:"
                    FacilityOfficeHeadcountFrameLabel.place(x=10, y=315)
                    self.FacilityOfficeHeadcountFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeHeadcountFrameEntry['width']=10
                    self.FacilityOfficeHeadcountFrameEntry.place(x=110, y=315)
                    self.FacilityOfficeHeadcountFrameEntry['state'] = DISABLED               

                    FacilityOfficeHRIDFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeHRIDFrameLabel["text"] = "HR Location ID:"
                    FacilityOfficeHRIDFrameLabel.place(x=190, y=315)
                    self.FacilityOfficeHRIDFrameEntry = Entry(self.FacilityFrame)
                    self.FacilityOfficeHRIDFrameEntry['width']=20
                    self.FacilityOfficeHRIDFrameEntry.place(x=300, y=315)
                    self.FacilityOfficeHRIDFrameEntry['state'] = DISABLED               

                    FacilityOfficeStatusFrameLabel = Label(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10))
                    FacilityOfficeStatusFrameLabel["text"] = "Status:"
                    FacilityOfficeStatusFrameLabel.place(x=10, y=340)
                    self.OfficeStatusValues = ['New Site','Quote Requested','Quote Received','Quote Approved',
                                               'Active','Renovation','Closed']
                    self.ComboBoxOfficeStatus = ttk.Combobox(self.FacilityFrame, state='disable', width = 25)
                    self.ComboBoxOfficeStatus['values'] = self.OfficeStatusValues
                    self.ComboBoxOfficeStatus.current(0)
                    self.ComboBoxOfficeStatus.place(x = 110, y = 340)


                    # CheckButton
                    #self.varpower = StringVar()
                    #self.FacilityBusinessUnitPowerCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varpower,
                    #                                                        onvalue='1',offvalue='0')

                    self.varpower = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitPowerCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varpower)
                    self.FacilityBusinessUnitPowerCheckbutton['text'] = 'Power'
                    self.FacilityBusinessUnitPowerCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitPowerCheckbutton.place(x=105, y=110)
                    #self.varpower.trace_variable("w", self.showstate) # this is pretty cool to trace the users :)
                    
                    self.varwater = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitWaterCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varwater)
                    self.FacilityBusinessUnitWaterCheckbutton['text'] = 'Water'
                    self.FacilityBusinessUnitWaterCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitWaterCheckbutton.place(x=170, y=110)

                    self.vartelecom = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitTelecomCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.vartelecom)
                    self.FacilityBusinessUnitTelecomCheckbutton['text'] = 'Telecom'
                    self.FacilityBusinessUnitTelecomCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitTelecomCheckbutton.place(x=235, y=110)

                    self.varSP = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitSPCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varSP)
                    self.FacilityBusinessUnitSPCheckbutton['text'] = 'Special Projects'
                    self.FacilityBusinessUnitSPCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitSPCheckbutton.place(x=315, y=110)

                    self.varCorp = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitCorpCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varCorp)
                    self.FacilityBusinessUnitCorpCheckbutton['text'] = 'Corporate'
                    self.FacilityBusinessUnitCorpCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitCorpCheckbutton.place(x=440, y=110)

                    self.varDC = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitDCCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varDC)
                    self.FacilityBusinessUnitDCCheckbutton['text'] = 'Datacenter'
                    self.FacilityBusinessUnitDCCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitDCCheckbutton.place(x=530, y=110)

                    self.varWH = IntVar(master=self.FacilityFrame)
                    self.FacilityBusinessUnitWHCheckbutton = Checkbutton(self.FacilityFrame,text="Helvetica", font=("Helvetica", 10),variable=self.varWH)
                    self.FacilityBusinessUnitWHCheckbutton['text'] = 'Warehouse'
                    self.FacilityBusinessUnitWHCheckbutton['state'] = 'disabled'
                    self.FacilityBusinessUnitWHCheckbutton.place(x=620, y=110)
                    
                    #------------------ TREE VIEW For Facility Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.FacilityTreeviewDataColumns = ('Facility ID','Facility Name','Business Unit','Office Type','Office Size','Lease End','Address',
                                                        'City','State','Head Count','Status','HW','Circuit','Added By')
                    
                    self.FacilityTreeview = ttk.Treeview(self.FacilityFrame,columns=self.FacilityTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.FacilityTreeviewysb = Scrollbar(self.FacilityFrame,orient=VERTICAL, command=self.FacilityTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.FacilityTreeviewxsb = Scrollbar(self.FacilityFrame,orient=HORIZONTAL, command=self.FacilityTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.FacilityTreeview['yscroll'] = self.FacilityTreeviewysb.set
                    self.FacilityTreeview['xscroll'] = self.FacilityTreeviewxsb.set

                    # setup headings and column 
                    self.FacilityTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                    self.FacilityTreeview.heading('#1', text='Facility ID',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Facility ID', False)) # E for East and W for West
                    self.FacilityTreeview.heading('#2', text='Facility Name', anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Facility Name', False)) # E for East and W for West
                    self.FacilityTreeview.heading('#3', text='Business Unit', anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Business Unit', False)) 
                    self.FacilityTreeview.heading('#4', text='Office Type',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Office Type', False)) 
                    self.FacilityTreeview.heading('#5', text='Office Size',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Office Size', False)) 
                    self.FacilityTreeview.heading('#6', text='Lease End',     anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Lease End', False))
                    self.FacilityTreeview.heading('#7', text='Address',       anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Address', False))
                    self.FacilityTreeview.heading('#8', text='City',          anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'City', False))
                    self.FacilityTreeview.heading('#9', text='State',         anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'State', False))
                    self.FacilityTreeview.heading('#10', text='Head Count',   anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Head Count', False)) 
                    self.FacilityTreeview.heading('#11', text='Status',       anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Status', False))
                    self.FacilityTreeview.heading('#12', text='HW',           anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'HW', False))
                    self.FacilityTreeview.heading('#13', text='Circuit',      anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Circuit', False)) 
                    self.FacilityTreeview.heading('#14', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.FacilityTreeview, 'Added By', False)) 
                    
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
                    self.FacilityTreeview.column('#13', stretch=1, width=10)
                    self.FacilityTreeview.column('#14', stretch=1, width=10)

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
                    #self.FacilityTreeview.bind("<Double-1>", self.On_Facility_Tree_Refresh)         

                    #------------------ TREE VIEW For Facility Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.FacilityFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)

                    self.ComboBoxRegionID = ttk.Combobox(self.FacilityFrame, state='disabled', width = 50)
                    self.ComboBoxRegionID.bind("<<ComboboxSelected>>", self.on_region_combo_changed)
                    self.ComboBoxRegionID.place(x = 110, y = 35)

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
                                self.on_region_combo_changed("event")
                                # Jorge
                                if (self.SitesID_Pre == "YES"):
                                    self.Call_Button_Facility_Validate()
                        self.Selection = 'edit_ok'
                    else:                        
                        self.on_Country_Table_Refresh()
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.FacilityWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.FacilityFrame,title='Facility',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")


def Main():
    print ("Testing the Facility Class....:")
    #root = widget.winfo_toplevel()
    #varpower = IntVar()
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','NO']    
    Facility = Class_Facility("BV",Windows_Scaling,location)
    Facility.Display_Facility_Window()


if __name__ == '__main__':
    Main()

