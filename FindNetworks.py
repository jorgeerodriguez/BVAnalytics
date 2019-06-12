#==========================================================================
# Program: FindNetworks.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Sep-5-2018
# Date Last Modified: Sep-5-2018
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
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['Find Networks','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False


#*********************************************************************************************************************************************
#                                   Networks Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************
class Class_FindNetworks:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPSummaryWindowExist = False
        self.FindCircuitWindowExist = False
        self.FindDeviceWindowExist = False
        self.CircuitsWindowExist = False
        self.FindNetworksWindowExist = False
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


    def on_FindNetworksWindow_quit(self):
        if (self.FindNetworksWindowExist):
            self.FindNetworksWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Find Networks','CLOSED Window']    
                Logging.Log(Parameter)
            self.FindNetworksWindow.destroy()

    def Call_Networks_ButtonFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Networks','Facility Button']    
            Logging.Log(Parameter)

        if (Is_Networks_Available):
            Facility = Class_Facility(ODBC_DSN_name,Windows_Scaling,self.location)
            Facility.Display_Facility_Window()

    def Call_Networks_ButtonNetworks(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Networks','Networks Button']    
            Logging.Log(Parameter)

        if (Is_Networks_Available):
            Networks = Class_Networks(ODBC_DSN_name,Windows_Scaling,self.location)
            Networks.Display_Networks_Window()

    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.NetworksTreeview.focus()
        dic = self.NetworksTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.NetworksTreeview.next(curItem)
        self.NetworksTreeview.selection_set((NextItem, NextItem))
        self.NetworksTreeview.focus(NextItem)

        NextItem = self.NetworksTreeview.focus()
        dic = self.NetworksTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
            
    def on_Networks_Tree_select_click(self,event):
            curItem = self.NetworksTreeview.focus()  
            dic = self.NetworksTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []              
                self.location = [self.NetworksTableCountryIDArray[curItem],self.NetworksTableRegionIDArray[curItem],
                            self.NetworksTableFacilityIDArray[curItem],self.NetworksSites]
                if (Is_Networks_Available):
                    self.buttonNetworks['state'] = ACTIVE
                if (Is_Facility_Available):
                    self.buttonNetworksFacility['state'] = ACTIVE

    def on_Tree_Networks_double_click(self,event):
            curItem = self.NetworksTreeview.focus()  
            dic = self.NetworksTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1
                self.location = []
                self.location = [self.NetworksTableCountryIDArray[curItem],self.NetworksTableRegionIDArray[curItem],
                            self.NetworksTableFacilityIDArray[curItem],self.NetworksSites]
                if (Is_Networks_Available):
                    Networks = Class_Networks(ODBC_DSN_name,Windows_Scaling,self.location)
                    Networks.Display_Networks_Window()


    def Call_ButtonFindNetworks(self):

        if Is_logging_Available:
            Parameter = []
            Parameter = ['Find Networks','Find Button']    
            Logging.Log(Parameter)

        #self.buttonNetworks['state'] = DISABLED
        #self.buttonNetworksFacility['state'] = DISABLED
        value = self.FindNetworksFrameEntry.get() # Value to querry
        index = (self.ComboBoxNetworksTableFields.current())
        field = (self.NetworksTableFieldsSQL[index])
        self.NetworksSites = "NO"
        if (len(value) > 0):
                if self.db.Connect():
                    if (value == '*'):
                        sql = """
                                    SELECT * FROM Networks
                                    ORDER BY IP_From_Num ASC, Slash ASC
                              """
                    else:
                        sql = """SELECT * FROM Networks 
                                WHERE %s LIKE '%s' ORDER BY IP_From_Num ASC, Slash ASC
                              """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):                        
                        #------------------------------- Deleting Tree View <BEGIN> --------
                        x = self.NetworksTreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.NetworksTreeview.delete(child)
                        #------------------------------- Deleting Tree View <END> --------
                        #-------------- Initializing Arrays <BEGIN> ----------------------
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
                        self.NetworkTable_ssigned_Array = []
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
                        self.results = []

                        self.NetworksTableArrayColumns = ('Network-Slash','Network','Slash','Subnet Mask','IPA','IPB','IPC','IPD','SMA','SMB','SMC','SMD',
                                                                  'No. Of Host','Network From','Network To','Network From Num','Network To Num','Asigned','Network Region',
                                                                  'Office Size','Status','Type','Vlan No.','Description','DHCP','Allocated','Provisioned','Added By',
                                                                  'Country','Region','Facility','Site')
                        
                        #-------------- Initializing Arrays <END> ----------------------
                        self.data_ready = True
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
                            self.NetworkTable_ssigned_Array.append(self.db.results[i][17].strip())
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
                            self.NetworksTableColumns = ('Network-Slash','Network','Slash','Subnet Mask','IPA','IPB','IPC','IPD','SMA','SMB','SMC','SMD',
                                                                          'No. Of Host','Network From','Network To','Network From Num','Network To Num','Asigned','Network Region',
                                                                          'Office Size','Status','Type','Vlan No.','Description','DHCP','Allocated','Provisioned','Added By',
                                                                          'Country','Region','Facility','Site')
                            if (self.NetworkTable_ssigned_Array[i] == 'YES') and (self.NetworkTable_VLAN_L3_Type_Array[i] != 'Summary'):
                                Network = "     "+self.NetworkTable_Network_Array[i]
                            else:
                                Network = self.NetworkTable_Network_Array[i]
                            item = [
                                    Network,
                                    #self.NetworkTable_Network_Array[i],
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
                            self.NetworksTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            self.NetworksTable = [
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
                                    self.NetworkTable_ssigned_Array[i],
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
                            self.results2.append(self.NetworksTable)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.FindNetworksFrame,title='Find Networks',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.FindNetworksFrame,title='Find Networks',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.FindNetworksFrame,title='Find Networks',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFindNetworks_From_Entry(self,event):
        self.Call_ButtonFindNetworks()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Find Networks"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.NetworksTableArrayColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.FindNetworksFrame,title='Find Networks',
                            message = '!!! The File was saved !!!')
                    self.FindNetworksWindow.title("Find Network                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Find Networks','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.FindNetworksFrame,title='Find Networks List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Find_Networks_Window(self):
        if not self.FindNetworksWindowExist:
            # Set up the Window
            self.FindNetworksWindowExist = True
            self.FindNetworksWindow = Tk()
            self.FindNetworksWindow.geometry('1350x700+350+70')
            self.FindNetworksWindow.title("Find Networks")
            self.FindNetworksWindow.protocol("WM_DELETE_WINDOW", self.on_FindNetworksWindow_quit)
            self.FindNetworksWindow.call('tk', 'scaling', Windows_Scaling)
            self.FindNetworksLabel = Label(self.FindNetworksWindow,text="Helvetica", font=("Helvetica", 13))
            self.FindNetworksLabel["text"] = "Find Networks"
            self.FindNetworksLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.FindNetworksWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_FindNetworksWindow_quit)
            self.FindNetworksWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.FindNetworksFrame = Frame(self.FindNetworksWindow)
            self.FindNetworksFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.FindNetworksFrame.rowconfigure(0, weight=1)
            self.FindNetworksFrame.columnconfigure(0, weight=1)

            # Setup Labels
            self.FindNetworksFrameLabel = Label(self.FindNetworksFrame,text="Helvetica", font=("Helvetica", 10))
            self.FindNetworksFrameLabel["text"] = "Find:"
            self.FindNetworksFrameLabel.place(x=10, y=10)

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

            # Setup ComboBox
            self.NetworksTableFields = [
                                #'Network_Slash',
                                'Network',
                                'Slash',
                                #'Subnet_Mask',
                                #'IPA',
                                #'IPB',
                                #'IPC',
                                #'IPD',
                                #'SMA',
                                #'SMB',
                                #'SMC',
                                #'SMD',
                                #'No_Of_Hosts',
                                #'IP_From',
                                #'IP_To',
                                #'IP_From_Num',
                                #'IP_To_Num',
                                'Assigned',
                                'Network Region ID',
                                'Office Size',
                                'Status',                        
                                'Type',
                                'VLAN No',
                                'Description',
                                'DHCP',
                                'Allocation Date',
                                'Provision Date',
                                #'Executed_by_UserID',
                                'Country ID',
                                'Region ID',
                                'Facility ID',
                                'Site ID'
                                ]

            self.NetworksTableFieldsSQL = [
                                #'Network_Slash',
                                'Network',
                                'Slash',
                                #'Subnet_Mask',
                                #'IPA',
                                #'IPB',
                                #'IPC',
                                #'IPD',
                                #'SMA',
                                #'SMB',
                                #'SMC',
                                #'SMD',
                                #'No_Of_Hosts',
                                #'IP_From',
                                #'IP_To',
                                #'IP_From_Num',
                                #'IP_To_Num',
                                'Assigned',
                                'Network_Region_ID',
                                'Office_Size',
                                'Status',                        
                                'VLAN_L3_Type',
                                'VLAN_L3_Name',
                                'Description',
                                'DHCP',
                                'Allocation_Date',
                                'Provision_Date',
                                #'Executed_by_UserID',
                                'Country_ID',
                                'Region_ID',
                                'Facility_ID',
                                'Site_ID'
                                ]

            self.ComboBoxNetworksTableFields = ttk.Combobox(self.FindNetworksFrame, state='readonly', width = 25)
            self.ComboBoxNetworksTableFields['values'] = self.NetworksTableFields
            self.ComboBoxNetworksTableFields.current(0)
            self.ComboBoxNetworksTableFields.place(x = 50, y = 10)

            # Setup Entry
            self.FindNetworksFrameEntry = Entry(self.FindNetworksFrame)
            self.FindNetworksFrameEntry['width']=50
            self.FindNetworksFrameEntry.place(x=235, y=10)
            self.FindNetworksFrameEntry.bind('<Return>', self.Call_ButtonFindNetworks_From_Entry)
            #FindNetworksFrameEntry['state'] = DISABLED

            if self.db.Connect():                
                # Setup Button
                self.buttonFindNetworksExecute = Button(self.FindNetworksFrame, text = 'Execute', command = self.Call_ButtonFindNetworks, state=ACTIVE)
                self.buttonFindNetworksExecute.place(x = 550, y = 7, width=75, height=25)

                #self.buttonNetworksFacility = Button(self.FindNetworksFrame, text = 'Facility', command = self.Call_Networks_ButtonFacility, state=DISABLED)
                #self.buttonNetworksFacility.place(x = 850, y = 7, width=75, height=25)

                #self.buttonNetworks = Button(self.FindNetworksFrame, text = 'Point of Contacts', command = self.Call_Networks_ButtonNetworks, state=DISABLED)
                #self.buttonNetworks.place(x = 950, y = 7, width=140, height=25)


                #------------------ TREE VIEW For Networks Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.NetworksTreeviewDataColumns = ('Network','Slash','Description','Allocated','Provisioned','Country','Region','Facility','Type','Vlan No.','DHCP','Added By')
                
                self.NetworksTreeview = ttk.Treeview(self.FindNetworksFrame,columns=self.NetworksTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.NetworksTreeviewysb = Scrollbar(self.FindNetworksFrame,orient=VERTICAL, command=self.NetworksTreeview.yview)           # <--- Make sure the frame is correct !!
                self.NetworksTreeviewxsb = Scrollbar(self.FindNetworksFrame,orient=HORIZONTAL, command=self.NetworksTreeview.xview)         # <--- Make sure the frame is correct !!
                self.NetworksTreeview['yscroll'] = self.NetworksTreeviewysb.set
                self.NetworksTreeview['xscroll'] = self.NetworksTreeviewxsb.set

                # setup headings and column
                self.NetworksTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West
                self.NetworksTreeview.heading('#1', text='Network',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Network', False)) # E for East and W for West
                self.NetworksTreeview.heading('#2', text='Slash',         anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Slash', False)) # E for East and W for West
                self.NetworksTreeview.heading('#3', text='Description',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Description', False)) 
                self.NetworksTreeview.heading('#4', text='Allocated',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Allocated', False)) 
                self.NetworksTreeview.heading('#5', text='Provisioned',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Provisioned', False)) 
                self.NetworksTreeview.heading('#6', text='Country',       anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Country', False)) 
                self.NetworksTreeview.heading('#7', text='Region',        anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Region', False)) 
                self.NetworksTreeview.heading('#8', text='Facility',      anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Facility', False)) 
                self.NetworksTreeview.heading('#9', text='Type',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Type', False)) 
                self.NetworksTreeview.heading('#10', text='Vlan No.',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Vlan No.', False)) 
                self.NetworksTreeview.heading('#11',text='DHCP',          anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'DHCP', False)) 
                self.NetworksTreeview.heading('#12', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworksTreeview, 'Added By', False)) 
                    
                self.NetworksTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.NetworksTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.NetworksTreeview.column('#2', stretch=1, width=10)
                self.NetworksTreeview.column('#3', stretch=1, width=10)
                self.NetworksTreeview.column('#4', stretch=1, width=10)
                self.NetworksTreeview.column('#5', stretch=1, width=10)
                self.NetworksTreeview.column('#6', stretch=1, width=10)
                self.NetworksTreeview.column('#7', stretch=1, width=10)
                self.NetworksTreeview.column('#8', stretch=1, width=10)
                self.NetworksTreeview.column('#9', stretch=1, width=10)
                self.NetworksTreeview.column('#10', stretch=1, width=10)
                self.NetworksTreeview.column('#11', stretch=1, width=10)
                self.NetworksTreeview.column('#12', stretch=1, width=10)

                # add tree and scrollbars to frame
                self.NetworksTreeview.grid(row=1, column=0, sticky=NSEW)
                self.NetworksTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.NetworksTreeviewxsb.grid(row=2, column=0, sticky=EW)


                # create fonts and tags
                self.NetworksTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                self.NetworksTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.NetworksTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                self.NetworksTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.NetworksTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.NetworksTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.NetworksTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                self.NetworksTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.NetworksTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.NetworksTreeview.tag_configure('Closed', font=('Helvetica', 8), background='red3') 
                self.NetworksTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                                      
                # Bind the double Click
                #self.NetworksTreeview.bind('<ButtonRelease-1>', self.on_Networks_Tree_select_click) # When Select the Tree
                #self.NetworksTreeview.bind("<Double-1>", self.on_Tree_Networks_double_click)         
                
                #------------------ TREE VIEW For Networks Database <END> -----------------------------------                
                self.FindNetworksWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Find Local Point Of Contacts',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   Networks Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Circuits Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindLPOC = Class_FindNetworks("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindLPOC.Find_Networks_Window()


if __name__ == '__main__':
    Main()
