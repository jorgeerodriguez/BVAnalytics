#==========================================================================
# Program: NetworkTemplatesBreakoutSettings.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: May-23-2018
# Date Last Modified: May-24-2018
# Summary: This is Class to for the NetworkTemplatesBreakoutSettings
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
    Parameter = ['Network Template Breakout Settings','OPEN Window']    
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

class Class_NetworkTemplatesBreakoutSettings:

    def __init__(self,DSN_Name,Windows_Scaling,Parameters,Treeview,Index):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.NetworkTemplatesBreakoutSettingsWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.NetworkTemplatesBreakOutTreeview = Treeview
        self.NetworkTemplatesBreakoutSettingsCostTreeview = Treeview
        self.results2 = Parameters
        self.index = Index
        if (len(Parameters) > 0 ):
            i = 0         
        
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
       
    def on_Sequence_combo_changed(self,event):
        CurItem = self.ComboBoxNetworkTemplatesBreakoutSettingsSequence.current()

        self.ComboBoxNetworkTemplatesBreakoutSettingsSequence['state'] = 'readonly'
        
        self.ComboBoxNetworkTemplatesBreakoutSettingsType['state'] = 'readonly'
        if (self.results2[CurItem][5] == "VLAN"):
            self.ComboBoxNetworkTemplatesBreakoutSettingsType.current(0)
        if (self.results2[CurItem][5] == "Layer3"):
            self.ComboBoxNetworkTemplatesBreakoutSettingsType.current(1)
        if (self.results2[CurItem][5] == "Summary"):
            self.ComboBoxNetworkTemplatesBreakoutSettingsType.current(2)

        self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry['state'] = 'normal'
        self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry.delete(0,END)
        self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry.insert(0,self.results2[CurItem][6])
        
        self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry['state'] = 'normal'
        self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry.delete(0,END)
        self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry.insert(0,self.results2[CurItem][7])

        self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP['state'] = 'readonly'
        if (self.results2[CurItem][8] == "YES"):
            self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP.current(0)
        else:
            self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP.current(1)
            
        self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT['state'] = 'readonly'
        if (self.results2[CurItem][9] == "YES"):
            self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT.current(0)
        else:
            self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT.current(1)
        
        self.ComboBoxNetworkTemplatesBreakoutSettingsSlash['state'] = 'readonly'
        self.ComboBoxNetworkTemplatesBreakoutSettingsSlash.current(self.results2[CurItem][10])

        self.ButtonNetworkTemplatesBreakoutSettingsOK['state'] = ACTIVE        
        self.ButtonNetworkTemplatesBreakoutSettingsCancel['state'] = ACTIVE 

        
    def on_NetworkTemplatesBreakoutSettingsWindow_quit(self):
        if (self.NetworkTemplatesBreakoutSettingsWindowExist):
            self.NetworkTemplatesBreakoutSettingsWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Network Template Breakout Settings','CLOSE Window']    
                Logging.Log(Parameter)
            self.NetworkTemplatesBreakoutSettingsWindow.destroy()


    def on_NetworkTemplatesBreakoutSettings_Table_Refresh(self):
        self.on_Sequence_combo_changed("event")
       


    def Call_Button_NetworkTemplatesBreakoutSettings_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Template Breakout Settings','ADD/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():

            CurItem = self.ComboBoxNetworkTemplatesBreakoutSettingsSequence.current()
            VLAN_L3_Type = self.Types[self.ComboBoxNetworkTemplatesBreakoutSettingsType.current()]
            VLAN_L3_Name = self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry.get()
            Description = self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry.get()
            DHCP = self.DHCP[self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP.current()]
            Split_More = self.SPLIT[self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT.current()]
            Split_More_Slash = self.ComboBoxNetworkTemplatesBreakoutSettingsSlash.current()
            
            self.ButtonNetworkTemplatesBreakoutSettingsOK['state'] = ACTIVE        
            self.ButtonNetworkTemplatesBreakoutSettingsCancel['state'] = ACTIVE 

            Primarykey = self.results2[CurItem][0]
            #print (Primarykey)
    
            sql = """
                    SELECT * FROM Network_Templates_Break_Out
                    WHERE Template_ID_SEQ1_SEQ2 = '%s'
                  """ % (Primarykey)
            if (self.db.Execute(sql)):
                #-------------- EDIT ----------------------
                '''
                    Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                    Template_ID              CHAR(100) NOT NULL,
                    Sequence_1               INT,
                    Sequence_2               INT,
                    Split_in_Slash           INT,
                    VLAN_L3_Type             CHAR(20), # VLAN | Layer3 | Summary
                    VLAN_L3_Name             CHAR(50),
                    Description              CHAR(100),
                    DHCP                     CHAR(10),
                    Split_More               CHAR(10),
                    Split_More_Slash         INT,
                    Executed_by_UserID       CHAR(20) )"""
                '''

                sql = "UPDATE Network_Templates_Break_Out SET VLAN_L3_Type = '%s', VLAN_L3_Name = '%s', Description = '%s', DHCP = '%s', \
                               Split_More = '%s', Split_More_Slash = '%d', Executed_by_UserID = '%s' \
                               WHERE Template_ID_SEQ1_SEQ2 = '%s'" % (VLAN_L3_Type,VLAN_L3_Name,Description,DHCP,Split_More,Split_More_Slash,
                                        self.Username,Primarykey)
                if (self.db.Add_Move_Change_Data(sql)):
                    mbox.showwarning(master=self.NetworkTemplatesBreakoutSettingsFrame,title='Network Template Breakout Settings',
                        message = '*** The Network Template Breakout Settings you entered was Updated ***')
                    
                    # SQL Querry to the NetworkTemplatesBreakOut Table
                    sql = """
                            SELECT * FROM Network_Templates_Break_Out
                            WHERE Template_ID = '%s'
                            ORDER BY Sequence_1 ASC, Sequence_2 ASC
                          """ % (self.results2[self.index][1])
                    
                    #------------------------------- Deleting Tree View <BEGIN> --------
                    x = self.NetworkTemplatesBreakOutTreeview.get_children()
                    if x != '()': # checks if there is something in the first row
                        for child in x:
                            #print (child)
                            self.NetworkTemplatesBreakOutTreeview.delete(child)
                    #------------------------------- Deleting Tree View <END> --------
                                        
                    if (self.db.Execute(sql)):
                        self.TemplateIDSEQ1SEQ2Array = []
                        self.TemplateIDArray = []
                        self.Sequence1Array = []
                        self.Sequence2Array = []
                        self.SplitinSlashArray = []
                        self.VLANL3TypeArray = []
                        self.VLAN_L3_NameArray = []
                        self.DescriptionArray = []
                        self.DHCPArray = []
                        self.SplitMoreArray = []
                        self.SplitMoreSlashArray = []
                        self.ExecutedbyUserIDArray = []
                        self.results = []
                        self.results2 = []
                        i = 0
                        '''

                            0  Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                            1  Template_ID              CHAR(100) NOT NULL,
                            2  Sequence_1               INT,
                            3  Sequence_2               INT,
                            4  Split_in_Slash           INT,
                            5  VLAN_L3_Type             CHAR(20),
                            6  VLAN_L3_Name             CHAR(50),
                            7  Description              CHAR(100),
                            8  DHCP                     CHAR(10),
                            9  Split_More               CHAR(10),
                            10 Split_More_Slash         INT,
                            11 Executed_by_UserID       CHAR(20)
                        '''

                        while (i < len(self.db.results)):
                            num = i + 1
                            self.TemplateIDSEQ1SEQ2Array.append(self.db.results[i][0].strip())
                            self.TemplateIDArray.append(self.db.results[i][1].strip())
                            self.Sequence1Array.append(self.db.results[i][2])
                            self.Sequence2Array.append(self.db.results[i][3])
                            self.SplitinSlashArray.append(self.db.results[i][4])
                            self.VLANL3TypeArray.append(self.db.results[i][5].strip())
                            if (self.db.results[i][6] == None):
                                self.VLAN_L3_NameArray.append("")
                            else:
                                self.VLAN_L3_NameArray.append(self.db.results[i][6].strip())
                            if (self.db.results[i][7] == None):
                                self.DescriptionArray.append("")
                            else:
                                self.DescriptionArray.append(self.db.results[i][7].strip())
                            self.DHCPArray.append(self.db.results[i][8].strip())
                            self.SplitMoreArray.append(self.db.results[i][9].strip())
                            self.SplitMoreSlashArray.append(self.db.results[i][10])
                            self.ExecutedbyUserIDArray.append(self.db.results[i][11].strip())
                            i = i + 1
                           # ('Sequence','Slash','Type','Description','DHCP','Split More','In Slash','Added By')
                        i = 0
                        while (i < len(self.TemplateIDSEQ1SEQ2Array)):
                            num = i + 1
                            tags = self.VLANL3TypeArray[i]                         
                            item = [
                                    (str(self.Sequence1Array[i])+"."+str(self.Sequence2Array[i])),
                                    self.SplitinSlashArray[i],
                                    self.VLANL3TypeArray[i],
                                    self.VLAN_L3_NameArray[i],
                                    self.DescriptionArray[i],
                                    self.DHCPArray[i],
                                    self.SplitMoreArray[i],
                                    self.SplitMoreSlashArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results.append(item)
                            item2 = [
                                    self.TemplateIDSEQ1SEQ2Array[i],
                                    self.TemplateIDArray[i],
                                    self.Sequence1Array[i],
                                    self.Sequence2Array[i],
                                    self.SplitinSlashArray[i],
                                    self.VLANL3TypeArray[i],
                                    self.VLAN_L3_NameArray[i],
                                    self.DescriptionArray[i],
                                    self.DHCPArray[i],
                                    self.SplitMoreArray[i],
                                    self.SplitMoreSlashArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results2.append(item2)
                            self.NetworkTemplatesBreakOutTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                else:
                    mbox.showerror(master=self.NetworkTemplatesBreakoutSettingsFrame,title='Network Template Breakout Settings',
                        message = '*** The Network Template Breakout you entered was NOT Upadted ***')
            else:
                    mbox.showerror(master=self.NetworkTemplatesBreakoutSettingsFrame,title='Network Template Breakout Settings',
                        message = '*** The Network Template Breakout Does not Exist ***')                    
        else:
            mbox.showerror(master=self.NetworkTemplatesBreakoutSettingsFrame,title='Network Template Breakout Settings',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
                          
        self.on_NetworkTemplatesBreakoutSettings_Table_Refresh()


    def Display_NetworkTemplatesBreakoutSettings_Window(self):
        #global NetworkTemplatesBreakoutSettingsLabel
 
        if not self.NetworkTemplatesBreakoutSettingsWindowExist:
            # Set up the Window
            self.NetworkTemplatesBreakoutSettingsWindowExist = True
            self.NetworkTemplatesBreakoutSettingsWindow = Tk()
            self.NetworkTemplatesBreakoutSettingsWindow.geometry('750x270+350+70')
            self.NetworkTemplatesBreakoutSettingsWindow.title("Network Templates Breakout Settings")
            self.NetworkTemplatesBreakoutSettingsWindow.protocol("WM_DELETE_WINDOW", self.on_NetworkTemplatesBreakoutSettingsWindow_quit)
            self.NetworkTemplatesBreakoutSettingsWindow.call('tk', 'scaling', self.Windows_Scaling)
            NetworkTemplatesBreakoutSettingsLabel = Label(self.NetworkTemplatesBreakoutSettingsWindow,text="Helvetica", font=("Helvetica", 19))
            NetworkTemplatesBreakoutSettingsLabel["text"] = "Network Templates Breakout Settings"
            NetworkTemplatesBreakoutSettingsLabel.pack()
                    
            # Setup Frame
            self.NetworkTemplatesBreakoutSettingsFrame = Frame(self.NetworkTemplatesBreakoutSettingsWindow)
            self.NetworkTemplatesBreakoutSettingsFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.NetworkTemplatesBreakoutSettingsFrame.rowconfigure(0, weight=1)
            self.NetworkTemplatesBreakoutSettingsFrame.columnconfigure(0, weight=1)

            if self.db.Connect():

                     # Setup Buttons

                    self.ButtonNetworkTemplatesBreakoutSettingsOK = Button(self.NetworkTemplatesBreakoutSettingsFrame, text = 'UPDATE', command = self.Call_Button_NetworkTemplatesBreakoutSettings_OK, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakoutSettingsOK.place(x = 420, y = 7, width=100, height=25)

                    self.ButtonNetworkTemplatesBreakoutSettingsCancel = Button(self.NetworkTemplatesBreakoutSettingsFrame, text = 'CLOSE', command = self.on_NetworkTemplatesBreakoutSettingsWindow_quit, state=DISABLED)
                    self.ButtonNetworkTemplatesBreakoutSettingsCancel.place(x = 530, y = 7, width=100, height=25)

                    # Setup Labels
                    NetworkTemplatesBreakoutSettingsSequenceFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsSequenceFrameLabel["text"] = "Sequence:"
                    NetworkTemplatesBreakoutSettingsSequenceFrameLabel.place(x=10, y=10)

                    NetworkTemplatesBreakoutSettingsTypeFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsTypeFrameLabel["text"] = "Type:"
                    NetworkTemplatesBreakoutSettingsTypeFrameLabel.place(x=10, y=40)
                    
                    NetworkTemplatesBreakoutSettingsVLANFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsVLANFrameLabel["text"] = "VLAN No.:"
                    NetworkTemplatesBreakoutSettingsVLANFrameLabel.place(x=10, y=70)
                    
                    NetworkTemplatesBreakoutSettingsDescriptionFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsDescriptionFrameLabel["text"] = "Description:"
                    NetworkTemplatesBreakoutSettingsDescriptionFrameLabel.place(x=10, y=100)

                    NetworkTemplatesBreakoutSettingsDHCPFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsDHCPFrameLabel["text"] = "DHCP:"
                    NetworkTemplatesBreakoutSettingsDHCPFrameLabel.place(x=10, y=130)

                    NetworkTemplatesBreakoutSettingsSplitFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsSplitFrameLabel["text"] = "Split More?:"
                    NetworkTemplatesBreakoutSettingsSplitFrameLabel.place(x=10, y=160)

                    NetworkTemplatesBreakoutSettingsSlashFrameLabel = Label(self.NetworkTemplatesBreakoutSettingsFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesBreakoutSettingsSlashFrameLabel["text"] = "New Slash:"
                    NetworkTemplatesBreakoutSettingsSlashFrameLabel.place(x=10, y=190)

                    # SQL Querry to the NetworkTemplatesBreakOut Table
                    sql = """
                            SELECT * FROM Network_Templates_Break_Out
                            WHERE Template_ID = '%s'
                            ORDER BY Sequence_1 ASC, Sequence_2 ASC
                          """ % (self.results2[self.index][1])
                    
                    if (self.db.Execute(sql)):
                        self.TemplateIDSEQ1SEQ2Array = []
                        self.TemplateIDArray = []
                        self.Sequence1Array = []
                        self.Sequence2Array = []
                        self.SplitinSlashArray = []
                        self.VLANL3TypeArray = []
                        self.VLAN_L3_NameArray = []
                        self.DescriptionArray = []
                        self.DHCPArray = []
                        self.SplitMoreArray = []
                        self.SplitMoreSlashArray = []
                        self.ExecutedbyUserIDArray = []
                        self.results = []
                        self.results2 = []
                        i = 0
                        '''

                            0  Template_ID_SEQ1_SEQ2    CHAR(150) NOT NULL PRIMARY KEY,
                            1  Template_ID              CHAR(100) NOT NULL,
                            2  Sequence_1               INT,
                            3  Sequence_2               INT,
                            4  Split_in_Slash           INT,
                            5  VLAN_L3_Type             CHAR(20),
                            6  VLAN_L3_Name             CHAR(50),
                            7  Description              CHAR(100),
                            8  DHCP                     CHAR(10),
                            9  Split_More               CHAR(10),
                            10 Split_More_Slash         INT,
                            11 Executed_by_UserID       CHAR(20)
                        '''

                        while (i < len(self.db.results)):
                            num = i + 1
                            self.TemplateIDSEQ1SEQ2Array.append(self.db.results[i][0].strip())
                            self.TemplateIDArray.append(self.db.results[i][1].strip())
                            self.Sequence1Array.append(self.db.results[i][2])
                            self.Sequence2Array.append(self.db.results[i][3])
                            self.SplitinSlashArray.append(self.db.results[i][4])
                            self.VLANL3TypeArray.append(self.db.results[i][5].strip())
                            if (self.db.results[i][6] == None):
                                self.VLAN_L3_NameArray.append("")
                            else:
                                self.VLAN_L3_NameArray.append(self.db.results[i][6].strip())
                            if (self.db.results[i][7] == None):
                                self.DescriptionArray.append("")
                            else:
                                self.DescriptionArray.append(self.db.results[i][7].strip())
                            self.DHCPArray.append(self.db.results[i][8].strip())
                            self.SplitMoreArray.append(self.db.results[i][9].strip())
                            self.SplitMoreSlashArray.append(self.db.results[i][10])
                            self.ExecutedbyUserIDArray.append(self.db.results[i][11].strip())
                            i = i + 1
                           # ('Sequence','Slash','Type','Description','DHCP','Split More','In Slash','Added By')
                        i = 0
                        while (i < len(self.TemplateIDSEQ1SEQ2Array)):
                            num = i + 1
                            tags = self.VLANL3TypeArray[i]                         
                            item = [
                                    (str(self.Sequence1Array[i])+"."+str(self.Sequence2Array[i])),
                                    self.SplitinSlashArray[i],
                                    self.VLANL3TypeArray[i],
                                    self.VLAN_L3_NameArray[i],
                                    self.DescriptionArray[i],
                                    self.DHCPArray[i],
                                    self.SplitMoreArray[i],
                                    self.SplitMoreSlashArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results.append(item)
                            item2 = [
                                    self.TemplateIDSEQ1SEQ2Array[i],
                                    self.TemplateIDArray[i],
                                    self.Sequence1Array[i],
                                    self.Sequence2Array[i],
                                    self.SplitinSlashArray[i],
                                    self.VLANL3TypeArray[i],
                                    self.VLAN_L3_NameArray[i],
                                    self.DescriptionArray[i],
                                    self.DHCPArray[i],
                                    self.SplitMoreArray[i],
                                    self.SplitMoreSlashArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results2.append(item2)
                            #self.NetworkTemplatesBreakOutTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.NetworkTemplatesBreakoutSettingsFrame,title='Network Templates Breakout Settings',
                            message = '*** The Template ID was not Found ***')                    

                    i = 0
                    seq = []
                    while (i < len(self.results2)):
                        seq.append(str(self.results2[i][2])+"."+str(self.results2[i][3]))
                        i = i + 1

                    # Setup ComboBox
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSequence = ttk.Combobox(self.NetworkTemplatesBreakoutSettingsFrame, state='readonly', width = 37)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSequence.bind("<<ComboboxSelected>>", self.on_Sequence_combo_changed)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSequence.place(x = 165, y = 10)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSequence['values'] = seq
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSequence.current(self.index)

                    self.Types = ["VLAN","Layer3","Summary"]
                    self.ComboBoxNetworkTemplatesBreakoutSettingsType = ttk.Combobox(self.NetworkTemplatesBreakoutSettingsFrame, state='disabled', width = 37)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsType.place(x = 165, y = 40)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsType['values'] = self.Types
                    self.ComboBoxNetworkTemplatesBreakoutSettingsType.current(0)
                        
                    # Setup Entry
                    self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry = Entry(self.NetworkTemplatesBreakoutSettingsFrame)
                    self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry['width']=40
                    self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry.place(x=165, y=70)
                    self.NetworkTemplatesBreakoutSettingsVlanNoFrameEntry['state'] = 'readonly'

                    self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry = Entry(self.NetworkTemplatesBreakoutSettingsFrame)
                    self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry['width']=40
                    self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry.place(x=165, y=100)
                    self.NetworkTemplatesBreakoutSettingsDescriptionFrameEntry['state'] = 'readonly'

                    self.DHCP = ["YES","NO"]
                    self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP = ttk.Combobox(self.NetworkTemplatesBreakoutSettingsFrame, state='disabled', width = 37)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP.place(x = 165, y = 130)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP['values'] = self.DHCP
                    self.ComboBoxNetworkTemplatesBreakoutSettingsDHCP.current(0)


                    self.SPLIT = ["YES","NO"]
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT = ttk.Combobox(self.NetworkTemplatesBreakoutSettingsFrame, state='disabled', width = 37)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT.place(x = 165, y = 160)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT['values'] = self.SPLIT
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSPLIT.current(0)

                    self.slash = []
                    i = 0
                    while (i < 33):
                        self.slash.append("Slash: ["+str(i)+"]")
                        i = i + 1
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSlash = ttk.Combobox(self.NetworkTemplatesBreakoutSettingsFrame, state='disabled', width = 37)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSlash.place(x = 165, y = 190)
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSlash['values'] = self.slash
                    self.ComboBoxNetworkTemplatesBreakoutSettingsSlash.current(0)

                    self.on_Sequence_combo_changed("event")
                    self.NetworkTemplatesBreakoutSettingsWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.NetworkTemplatesBreakoutSettingsFrame,title='NetworkTemplatesBreakoutSettings',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")


################

def on_NetworkTemplatesBreakoutSettingsWindow_quit():
    global NetworkTemplatesBreakoutSettings
    global ModelWindow

    NetworkTemplatesBreakoutSettings.on_NetworkTemplatesBreakoutSettingsWindow_quit()
    ModelWindow.destroy()
    

                
def Display_Model_Window():
    #global NetworkTemplatesBreakoutSettingsLabel
    global NetworkTemplatesBreakoutSettingsCostTreeview
    global ModelWindowExist
    global db
    global NetworkTemplatesBreakoutSettings
    global ModelWindow
 
    if not ModelWindowExist:
        # Set up the Window
        ModelWindowExist = True
        ModelWindow = Tk()
        ModelWindow.geometry('1000x300+350+70')
        ModelWindow.title("NetworkTemplatesBreakoutSettings")
        ModelWindow.protocol("WM_DELETE_WINDOW", on_NetworkTemplatesBreakoutSettingsWindow_quit)
        ModelWindow.call('tk', 'scaling', "1.3")
        ModelLabel = Label(ModelWindow,text="Helvetica", font=("Helvetica", 19))
        ModelLabel["text"] = "Network Templates Settings"
        ModelLabel.pack()
                    
        # Setup Frame
        NetworkTemplatesBreakOutFrame = Frame(ModelWindow)
        NetworkTemplatesBreakOutFrame.pack(side=TOP, fill=BOTH, expand=Y)

        # set frame resizing priorities
        NetworkTemplatesBreakOutFrame.rowconfigure(0, weight=1)
        NetworkTemplatesBreakOutFrame.columnconfigure(0, weight=1)

        if db.Connect():

            #------------------ TREE VIEW For Network Templates Breakout Database <BEGIN> -----------------------------------
            # Create Tree and Scrollbars
            NetworkTemplatesBreakOutTreeviewDataColumns = ('Sequence','Slash','Type','VLAN No.','Description','DHCP','Split More','In Slash','Added By')
                    
            NetworkTemplatesBreakOutTreeview = ttk.Treeview(NetworkTemplatesBreakOutFrame,columns=NetworkTemplatesBreakOutTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
            NetworkTemplatesBreakOutTreeviewysb = Scrollbar(NetworkTemplatesBreakOutFrame,orient=VERTICAL, command=NetworkTemplatesBreakOutTreeview.yview)           # <--- Make sure the frame is correct !!
            NetworkTemplatesBreakOutTreeviewxsb = Scrollbar(NetworkTemplatesBreakOutFrame,orient=HORIZONTAL, command=NetworkTemplatesBreakOutTreeview.xview)         # <--- Make sure the frame is correct !!
            NetworkTemplatesBreakOutTreeview['yscroll'] = NetworkTemplatesBreakOutTreeviewysb.set
            NetworkTemplatesBreakOutTreeview['xscroll'] = NetworkTemplatesBreakOutTreeviewxsb.set

            # setup headings and column 
            NetworkTemplatesBreakOutTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
            NetworkTemplatesBreakOutTreeview.heading('#1', text='Sequence',   anchor=W) # E for East and W for West
            NetworkTemplatesBreakOutTreeview.heading('#2', text='Slash',   anchor=W) # E for East and W for West
            NetworkTemplatesBreakOutTreeview.heading('#3', text='Type',    anchor=W) # E for East and W for West
            NetworkTemplatesBreakOutTreeview.heading('#4', text='VLAN No.',  anchor=W) 
            NetworkTemplatesBreakOutTreeview.heading('#5', text='Description',  anchor=W) 
            NetworkTemplatesBreakOutTreeview.heading('#6', text='DHCP',  anchor=W) 
            NetworkTemplatesBreakOutTreeview.heading('#7', text='Split More',  anchor=W) 
            NetworkTemplatesBreakOutTreeview.heading('#8', text='In Slash',  anchor=W) 
            NetworkTemplatesBreakOutTreeview.heading('#9', text='Added By',     anchor=W) 
                    
            NetworkTemplatesBreakOutTreeview.column('#0', stretch=1, width=3 , anchor=W)
            NetworkTemplatesBreakOutTreeview.column('#1', stretch=1, width=10, anchor=W)
            NetworkTemplatesBreakOutTreeview.column('#2', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#3', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#4', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#5', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#6', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#7', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#8', stretch=1, width=10)
            NetworkTemplatesBreakOutTreeview.column('#9', stretch=1, width=10)

            NetworkTemplatesBreakOutTreeview.grid(row=1, column=0, sticky=NSEW)
            NetworkTemplatesBreakOutTreeviewysb.grid(row=1, column=1, sticky=NS)
            NetworkTemplatesBreakOutTreeviewxsb.grid(row=2, column=0, sticky=EW)

            # create fonts and tags
            NetworkTemplatesBreakOutTreeview.tag_configure('None', font=('Helvetica', 10), background='gray1')
            NetworkTemplatesBreakOutTreeview.tag_configure('Summary', font=('Helvetica', 10), background='tan1')
            NetworkTemplatesBreakOutTreeview.tag_configure('0', font=('Helvetica', 10), background='gray1')
                   
            # Bind the double Click
            #self.NetworkTemplatesBreakOutTreeview.bind('<ButtonRelease-1>', self.on_NetworkTemplatesBreakOut_Tree_select_click) # When Select the Tree
            #self.NetworkTemplatesBreakOutTreeview.bind("<Double-1>", self.On_NetworkTemplatesBreakOut_Tree_Refresh)         

            #------------------ TREE VIEW For Device Model Country Cost Database <END> -----------------------------------                    
            results2 = []
            item=['Slash 21 for Large Offices-1-0','Slash 21 for Large Offices',1,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-2-0','Slash 21 for Large Offices',2,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-3-0','Slash 21 for Large Offices',3,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-4-0','Slash 21 for Large Offices',4,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-5-0','Slash 21 for Large Offices',5,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-6-0','Slash 21 for Large Offices',6,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-7-0','Slash 21 for Large Offices',7,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-8-0','Slash 21 for Large Offices',8,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-9-0','Slash 21 for Large Offices',9,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-10-0','Slash 21 for Large Offices',10,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-11-0','Slash 21 for Large Offices',11,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-12-0','Slash 21 for Large Offices',12,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-13-0','Slash 21 for Large Offices',13,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-14-0','Slash 21 for Large Offices',14,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-15-0','Slash 21 for Large Offices',15,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-16-0','Slash 21 for Large Offices',16,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-17-0','Slash 21 for Large Offices',17,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-18-0','Slash 21 for Large Offices',18,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-19-0','Slash 21 for Large Offices',19,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-20-0','Slash 21 for Large Offices',20,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-21-0','Slash 21 for Large Offices',21,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-22-0','Slash 21 for Large Offices',22,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-23-0','Slash 21 for Large Offices',23,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-24-0','Slash 21 for Large Offices',24,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            item=['Slash 21 for Large Offices-25-0','Slash 21 for Large Offices',25,0,24,'VLAN','0','NO Description','NO','NO',0,'rod90731']
            results2.append(item)
            results = []
            i = 0
            while (i < len(results2)):
                num = i + 1
                tags = results2[i][5]                         
                item = [
                        (str(results2[i][2])+"."+str(results2[i][3])),
                        results2[i][4],
                        results2[i][5],
                        results2[i][6],
                        results2[i][7],
                        results2[i][8],
                        results2[i][9],
                        results2[i][10],
                        results2[i][11]
                        ]
                results.append(item)
                NetworkTemplatesBreakOutTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                i = i + 1            
            Treeview = NetworkTemplatesBreakOutTreeview
            NetworkTemplatesBreakoutSettings = Class_NetworkTemplatesBreakoutSettings("BV",Windows_Scaling,results2,Treeview,4)
            NetworkTemplatesBreakoutSettings.Display_NetworkTemplatesBreakoutSettings_Window()   
            ModelWindow.mainloop()

        else:
            mbox.showerror(master=ModelFrame,title='Device Model',
                            message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                    + 'Please make sure the ODBC DSN Name mathes: ['
                                    + self.ODBC_name + "]")

                
################


def Main():
    global ModelWindowExist
    global db

    db = ODBC("BV")
    print ("Testing the NetworkTemplatesBreakoutSettings Class....:")
    ModelWindowExist = False
    Display_Model_Window()
    #Parameters = []
    #Parameters = ["AVPN","AVPN","USA","EDIT",""]
    #Parameters = ["AVPN","AVPN","","ADD",""]
    #Treeview = NetworkTemplatesBreakoutSettingsCostTreeview
    #NetworkTemplatesBreakoutSettings = Class_NetworkTemplatesBreakoutSettings("BV",Windows_Scaling,Parameters,Treeview)
    #NetworkTemplatesBreakoutSettings.Display_NetworkTemplatesBreakoutSettings_Window()


if __name__ == '__main__':
    Main()

