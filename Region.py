#==========================================================================
# Program: Region.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-23-2017
# Date Last Modified: Dec-29-2017
# Summary: This is Class to for the Region
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
    from Country import *
    Is_Country_Available = True
    Country = Class_Country(ODBC_DSN_name,Windows_Scaling)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Country Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_Country_Available = False  

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Region','OPEN Window']    
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

class Class_Region:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.RegionWindowExist = False
        self.Username = os.getlogin()
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

#---------------------------- COUNTRY SELECTION SECTION <BEGIN> ------------------------
        
    def Display_Country_Window(self):    
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','Country Window']    
            Logging.Log(Parameter)
        Country.Display_Country_Window()

    def on_country_combo_changed(self,event):
        #print (self.ComboBoxCoutryID.current())
        #print (self.CountryIDComboBox[self.ComboBoxCoutryID.current()])
        #print (self.CountryNameComboBox[self.ComboBoxCoutryID.current()])

        self.RegionIDFrameEntry['state'] = 'normal'
        self.RegionIDFrameEntry.delete(0,END)
        self.RegionIDFrameEntry['state'] = 'readonly'
            
        self.RegionNameFrameEntry['state'] = 'normal'
        self.RegionNameFrameEntry.delete(0,END)
        self.RegionNameFrameEntry['state'] = 'readonly'
        
        if self.db.Connect():
            # SQL Querry to the Region Table
            sql = """
                    SELECT * FROM Region
                    WHERE Country_ID = '%s'
                  """ % (self.CountryIDComboBox[self.ComboBoxCoutryID.current()])
            
            if (self.db.Execute(sql)):
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.RegionTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.RegionTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                self.data_ready = True
                self.results = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [self.db.results[i][1],
                            self.db.results[i][2],
                            self.db.results[i][3],
                            self.db.results[i][4]]
                    self.results.append(item)
                    self.RegionTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                self.ButtonRegionAdd['state']    = ACTIVE
                self.ButtonRegionEdit['state']   = DISABLED
                self.ButtonRegionRemove['state'] = DISABLED
                self.ButtonRegionOK['state']     = DISABLED
                self.ButtonRegionCancel['state'] = DISABLED
            else:
                mbox.showerror(master=self.RegionFrame,title='Region',
                message = 'No Records found')
                self.data_ready = False
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.RegionTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.RegionTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                self.ButtonRegionAdd['state']    = ACTIVE
                self.ButtonRegionEdit['state']   = DISABLED
                self.ButtonRegionRemove['state'] = DISABLED
                self.ButtonRegionOK['state']     = DISABLED
                self.ButtonRegionCancel['state'] = DISABLED

                self.RegionIDFrameEntry['state'] = 'normal'
                self.RegionIDFrameEntry.delete(0,END)
                self.RegionIDFrameEntry['state'] = 'readonly'
                
                self.RegionNameFrameEntry['state'] = 'normal'
                self.RegionNameFrameEntry.delete(0,END)
                self.RegionNameFrameEntry['state'] = 'readonly'
        else:
            mbox.showerror(master=self.RegionFrame,title='Region',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

            
    def on_Country_Table_Refresh(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','Country Refresh']    
            Logging.Log(Parameter)
        if self.db.Connect():            
            self.CountryIDComboBox = []
            self.CountryNameComboBox = []               

            # SQL Querry to the Region Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """ SELECT * FROM COUNTRY ORDER BY Country_Name ASC """
            if (self.db.Execute(sql)):
                i = 0
                self.progress['maximum'] = len(self.db.results)
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CountryIDComboBox.append(self.db.results[i][0].strip())
                    self.CountryNameComboBox.append(self.db.results[i][1].strip())
                    i = i + 1
                    self.progress['value'] = i
                self.ComboBoxCoutryID['values'] = self.CountryNameComboBox
                if (len(self.CountryNameComboBox)== 0):
                    self.ComboBoxCoutryID['state'] = DISABLED
                else:
                    self.ComboBoxCoutryID.current(0)
                    self.on_Region_Table_Refresh()
            else:
                mbox.showerror(master=self.RegionFrame,title='Region',
                message = 'No Country Records found')
            ##self.db.Disconnect()
        else:
            mbox.showerror(master=self.RegionFrame,title='Region',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
        

#---------------------------- COUNTRY SELECTION SECTION <END> ------------------------

        
    def on_RegionWindow_quit(self):
        if (self.RegionWindowExist):
            self.RegionWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Region','CLOSE Window']    
                Logging.Log(Parameter)
            self.RegionWindow.destroy()


    def on_Region_Table_Refresh(self):
        self.on_country_combo_changed("event")
       
    def Call_Button_Region_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','ADD Button']    
            Logging.Log(Parameter)
        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.ComboBoxCoutryID['state'] = DISABLED
        
        self.RegionIDFrameEntry['state'] = 'normal'
        self.RegionNameFrameEntry['state'] = 'normal'
        self.RegionIDFrameEntry.delete(0,END)
        self.RegionNameFrameEntry.delete(0,END)

        self.ButtonRegionAdd['state']    = DISABLED
        self.ButtonRegionEdit['state']   = DISABLED
        self.ButtonRegionRemove['state'] = DISABLED
        self.ButtonRegionOK['state']     = ACTIVE
        self.ButtonRegionCancel['state'] = ACTIVE


    def Call_Button_Region_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','EDIT Button']    
            Logging.Log(Parameter)
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'

        self.ComboBoxCoutryID['state'] = DISABLED
        
        self.RegionIDFrameEntry['state'] = 'readonly'
        self.RegionNameFrameEntry['state'] = 'normal'

        self.ButtonRegionAdd['state']    = DISABLED
        self.ButtonRegionEdit['state']   = DISABLED
        self.ButtonRegionRemove['state'] = DISABLED
        self.ButtonRegionOK['state']     = ACTIVE
        self.ButtonRegionCancel['state'] = ACTIVE


    def Call_Button_Region_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','REMOVE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.RegionFrame,title='Region',message = 'Are you Sure you want to Remove it?')):
                RegionID = self.RegionIDFrameEntry.get()
                RegionName = self.RegionNameFrameEntry.get()
                CountryID = self.CountryIDComboBox[self.ComboBoxCoutryID.current()]
                PrimaryKey = (CountryID+"-"+RegionID)
                if (RegionID == 'UNKNOWN'):
                    mbox.showerror(master=self.RegionFrame,title='Region',
                        message = '*** The Region ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Facility --------
                    sql = """
                            SELECT * FROM FACILITY
                            WHERE Region_ID = '%s'
                          """ % (RegionID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.RegionFrame,title='Region',
                                    message = '*** The Region ID you entered CANNOT be Removed Becuase is used in a Facility ***')
                            self.on_Region_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Region
                                WHERE Country_ID_Region_ID = '%s'
                              """ % (PrimaryKey)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Region WHERE Country_ID_Region_ID = '%s'" % (PrimaryKey)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.RegionFrame,title='Region',
                                        message = '*** The Region ID you entered was Removed ***')
                                self.on_Region_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.RegionFrame,title='Region',
                                        message = '*** The Region ID you entered was NOT Removed ***')
                                self.on_Region_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.RegionFrame,title='Region',
                            message = '*** The Region ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.RegionFrame,title='Region',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Region_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','OK/UPDATE Button']    
            Logging.Log(Parameter)
        if self.db.Connect():
            self.ComboBoxCoutryID['state'] = 'readonly'

            self.ButtonRegionAdd['state']    = DISABLED
            self.ButtonRegionEdit['state']   = DISABLED
            self.ButtonRegionRemove['state'] = DISABLED
            self.ButtonRegionOK['state']     = DISABLED #
            self.ButtonRegionCancel['state'] = DISABLED #

            RegionID = self.RegionIDFrameEntry.get()
            RegionName = self.RegionNameFrameEntry.get()
            CountryID = self.CountryIDComboBox[self.ComboBoxCoutryID.current()]
            PrimaryKey = (CountryID+"-"+RegionID)
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(RegionID) > 0) and (len(RegionName) > 0)):
                    sql = """
                            SELECT * FROM Region
                            WHERE Region_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):
                        mbox.showerror(master=self.RegionFrame,title='Region',
                        message = '*** The Region ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Region(Country_ID_Region_ID, Country_ID, Region_ID, \
                                   Region_Name, Executed_by_UserID) \
                                   VALUES ('%s','%s','%s', '%s', '%s')" % \
                                   (PrimaryKey,CountryID,RegionID, RegionName, self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.RegionFrame,title='Region',
                                message = '*** The Region ID you entered was Added ***')
                            self.on_Region_Table_Refresh()
                            self.RegionIDFrameEntry['state'] = 'readonly'                           
                            self.RegionNameFrameEntry['state'] = 'readonly'

                        else:
                            #print ("Error adding the record, posible dupliated it")
                            #self.db.Disconnect()
                            mbox.showerror(master=self.RegionFrame,title='Region',
                                message = '*** The Region ID you entered was NOT Added ***')
                            self.on_Region_Table_Refresh()
                else:
                     mbox.showerror(master=self.RegionFrame,title='Region',
                        message = '*** The Region ID and Region Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                if (len(RegionName) == 0):
                    mbox.showerror(master=self.RegionFrame,title='Region',
                    message = '*** The Region Name Cannot be BLANK ***')
                else:
                    sql = """
                            SELECT * FROM Region
                            WHERE Country_ID_Region_ID = '%s'
                          """ % (PrimaryKey)
                    if (self.db.Execute(sql)):

                        sql = "UPDATE Region SET Region_Name = '%s', Executed_by_UserID = '%s' WHERE Country_ID_Region_ID = '%s'" % (RegionName,self.Username,PrimaryKey)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.RegionFrame,title='Region',
                                message = '*** The Region ID you entered was Updated ***')
                            self.on_Region_Table_Refresh()
                            self.RegionNameFrameEntry['state'] = 'readonly'
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.RegionFrame,title='Region',
                                message = '*** The Region ID you entered was NOT Upadted ***')
                            self.on_Region_Table_Refresh()
                            self.RegionNameFrameEntry['state'] = 'readonly'
                    else:
                        mbox.showerror(master=self.RegionFrame,title='Region',
                        message = '*** The Region ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.RegionFrame,title='Region',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_Region_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Region','CANCEL Button']    
            Logging.Log(Parameter)
        self.ButtonRegionAdd['state']    = ACTIVE
        self.ButtonRegionEdit['state']   = DISABLED
        self.ButtonRegionRemove['state'] = DISABLED
        self.ButtonRegionOK['state']     = DISABLED
        self.ButtonRegionCancel['state'] = DISABLED

        self.ComboBoxCoutryID['state'] = 'readonly'

        self.RegionIDFrameEntry['state'] = 'normal'
        self.RegionIDFrameEntry.delete(0,END)
        self.RegionIDFrameEntry['state'] = 'readonly'
            
        self.RegionNameFrameEntry['state'] = 'normal'
        self.RegionNameFrameEntry.delete(0,END)
        self.RegionNameFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_Region_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.RegionTreeview.focus()
        dic = self.RegionTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            #print (values)
            #print (values[0])
            #print (values[1])

            self.ButtonRegionAdd['state']    = ACTIVE
            self.ButtonRegionEdit['state']   = ACTIVE
            self.ButtonRegionRemove['state'] = ACTIVE
            self.ButtonRegionOK['state']     = DISABLED
            self.ButtonRegionCancel['state'] = DISABLED

            self.ComboBoxCoutryID['state'] = 'readonly'

            self.RegionIDFrameEntry['state'] = 'normal'
            self.RegionIDFrameEntry.delete(0,END)
            self.RegionIDFrameEntry.insert(0,str(values[1]).strip())
            self.RegionIDFrameEntry['state'] = 'readonly'
            
            self.RegionNameFrameEntry['state'] = 'normal'
            self.RegionNameFrameEntry.delete(0,END)
            self.RegionNameFrameEntry.insert(0,str(values[2]).strip())
            self.RegionNameFrameEntry['state'] = 'readonly'

    def On_Region_Tree_Refresh(self,event):
        #--- Double Click --
        self.on_Region_Table_Refresh()


    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Region List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.RegionTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.RegionFrame,title='Region List',
                            message = '!!! The File was saved !!!')
                    self.RegionWindow.title("Region                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Region','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.RegionFrame,title='Region List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Display_Region_Window(self):
        #global RegionLabel
 
        if not self.RegionWindowExist:
            # Set up the Window
            self.RegionWindowExist = True
            self.RegionWindow = Tk()
            self.RegionWindow.geometry('1000x410+350+70')
            self.RegionWindow.title("Region")
            self.RegionWindow.protocol("WM_DELETE_WINDOW", self.on_RegionWindow_quit)
            self.RegionWindow.call('tk', 'scaling', self.Windows_Scaling)
            RegionLabel = Label(self.RegionWindow,text="Helvetica", font=("Helvetica", 19))
            RegionLabel["text"] = "Region"
            RegionLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.RegionWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_RegionWindow_quit)
            self.RegionWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.RegionFrame = Frame(self.RegionWindow)
            self.RegionFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.RegionFrame.rowconfigure(0, weight=1)
            self.RegionFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons
                    if (Is_Country_Available):
                        self.ButtonCountryAdd = Button(self.RegionFrame, text = '+ Country', command = self.Display_Country_Window, state=ACTIVE)
                        self.ButtonCountryAdd.place(x = 450, y = 7, width=75, height=25)

                        self.ButtonCountryRefresh = Button(self.RegionFrame, text = 'Refresh', command = self.on_Country_Table_Refresh, state=ACTIVE)
                        self.ButtonCountryRefresh.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonRegionAdd = Button(self.RegionFrame, text = 'Add', command = self.Call_Button_Region_Add, state=DISABLED)
                    self.ButtonRegionAdd.place(x = 450, y = 35, width=75, height=25)

                    self.ButtonRegionEdit = Button(self.RegionFrame, text = 'Edit', command = self.Call_Button_Region_Edit, state=DISABLED)
                    self.ButtonRegionEdit.place(x = 550, y = 35, width=75, height=25)

                    self.ButtonRegionRemove = Button(self.RegionFrame, text = 'Remove', command = self.Call_Button_Region_Remove, state=DISABLED)
                    self.ButtonRegionRemove.place(x = 650, y = 35, width=75, height=25)

                    self.ButtonRegionOK = Button(self.RegionFrame, text = 'OK / UPDATE', command = self.Call_Button_Region_OK, state=DISABLED)
                    self.ButtonRegionOK.place(x = 750, y = 35, width=100, height=25)

                    self.ButtonRegionCancel = Button(self.RegionFrame, text = 'Cancel', command = self.Call_Button_Region_Cancel, state=DISABLED)
                    self.ButtonRegionCancel.place(x = 875, y = 35, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.RegionFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=63)

                    # Setup Labels
                    CountryIDFrameLabel = Label(self.RegionFrame,text="Helvetica", font=("Helvetica", 10))
                    CountryIDFrameLabel["text"] = "Country Name:"
                    CountryIDFrameLabel.place(x=10, y=10)
                    
                    RegionIDFrameLabel = Label(self.RegionFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionIDFrameLabel["text"] = "Region ID:"
                    RegionIDFrameLabel.place(x=10, y=35)
                    
                    RegionNameFrameLabel = Label(self.RegionFrame,text="Helvetica", font=("Helvetica", 10))
                    RegionNameFrameLabel["text"] = "Region Name:"
                    RegionNameFrameLabel.place(x=10, y=60)

                   # Setup Entry
                    self.RegionIDFrameEntry = Entry(self.RegionFrame)
                    self.RegionIDFrameEntry['width']=50
                    self.RegionIDFrameEntry.place(x=110, y=35)
                    self.RegionIDFrameEntry['state'] = DISABLED

                    self.RegionNameFrameEntry = Entry(self.RegionFrame)
                    self.RegionNameFrameEntry['width']=50
                    self.RegionNameFrameEntry.place(x=110, y=60)
                    self.RegionNameFrameEntry['state'] = DISABLED               
                    
                    #------------------ TREE VIEW For Region Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.RegionTreeviewDataColumns = ('Country ID','Region ID','Region Name','Added By')
                    
                    self.RegionTreeview = ttk.Treeview(self.RegionFrame,columns=self.RegionTreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                    self.RegionTreeviewysb = Scrollbar(self.RegionFrame,orient=VERTICAL, command=self.RegionTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.RegionTreeviewxsb = Scrollbar(self.RegionFrame,orient=HORIZONTAL, command=self.RegionTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.RegionTreeview['yscroll'] = self.RegionTreeviewysb.set
                    self.RegionTreeview['xscroll'] = self.RegionTreeviewxsb.set

                    # setup headings and column 
                    self.RegionTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.RegionTreeview.heading('#1', text='Country ID',   anchor=W,command=lambda: self.treeview_sort_column(self.RegionTreeview, 'Country ID', False)) # E for East and W for West
                    self.RegionTreeview.heading('#2', text='Region ID',    anchor=W,command=lambda: self.treeview_sort_column(self.RegionTreeview, 'Region ID', False)) # E for East and W for West
                    self.RegionTreeview.heading('#3', text='Region Name',  anchor=W,command=lambda: self.treeview_sort_column(self.RegionTreeview, 'Region Name', False)) 
                    self.RegionTreeview.heading('#4', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.RegionTreeview, 'Added By', False)) 
                    
                    self.RegionTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.RegionTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.RegionTreeview.column('#2', stretch=1, width=10)
                    self.RegionTreeview.column('#3', stretch=1, width=10)
                    self.RegionTreeview.column('#4', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.RegionTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.RegionTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.RegionTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.RegionTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.RegionTreeview.bind('<ButtonRelease-1>', self.on_Region_Tree_select_click) # When Select the Tree
                    #self.RegionTreeview.bind("<Double-1>", self.On_Region_Tree_Refresh)         

                    #------------------ TREE VIEW For Region Database <END> -----------------------------------

                    # Setup ComboBox
                    self.ComboBoxCoutryID = ttk.Combobox(self.RegionFrame, state='readonly', width = 50)
                    self.ComboBoxCoutryID.bind("<<ComboboxSelected>>", self.on_country_combo_changed)
                    self.ComboBoxCoutryID.place(x = 110, y = 10)
                    self.on_Country_Table_Refresh()           
                    #self.db.Disconnect() # No needit for now since the Connection is alrady been done and it is faster.
                    self.RegionWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.RegionFrame,title='Region',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the Region Class....:")
    Region = Class_Region("BV",Windows_Scaling)
    Region.Display_Region_Window()


if __name__ == '__main__':
    Main()

