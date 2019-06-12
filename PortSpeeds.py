#==========================================================================
# Program: PortSpeed.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: April-22-2018
# Date Last Modified: April-22-2018
# Summary: This is Class to for the PortSpeed
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
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
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
    Parameter = ['PortSpeed','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

try:
    from SaveAs import *
    Is_SaveAs_Available = True
    #ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,FileName,"Ver 4.0")
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_SaveAs_Available = False

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_PortSpeed:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.PortSpeedWindowExist = False
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


    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Port Speed List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.PortSpeedTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.PortSpeedIDArray)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.PortSpeedIDArray[i],
                            self.PortSpeedNameArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.PortSpeedFrame,title='Port Speed List',
                            message = '!!! The File was saved !!!')
                    self.PortSpeedWindow.title("Port Speed                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['PortSpeed','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.PortSpeedFrame,title='Port Speed List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def on_PortSpeedWindow_quit(self):
        if (self.PortSpeedWindowExist):
            self.PortSpeedWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['PortSpeed','CLOSE Window']    
                Logging.Log(Parameter) 
            self.PortSpeedWindow.destroy()

    def on_PortSpeed_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():

            self.PortSpeedIDFrameEntry['state'] = 'readonly'
            self.PortSpeedNameFrameEntry['state'] = 'readonly'
            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.PortSpeedTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.PortSpeedTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                    Port_Speed_ID           CHAR(20) NOT NULL PRIMARY KEY,
                    Description             CHAR(50),
                    Executed_by_UserID      CHAR(20))"""
            '''
            # SQL Querry to the PortSpeed Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            # SQL Querry to the PortSpeed Table
            self.data_ready = False
            sql = """
                        SELECT * FROM Port_Speed
                        ORDER BY Description ASC
                  """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.PortSpeedIDArray = []
                self.PortSpeedNameArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                '''
                    0 Port_Speed_ID           CHAR(20) NOT NULL PRIMARY KEY,
                    1 Description             CHAR(50),
                    2 Executed_by_UserID      CHAR(20))"""
                '''
                while (i < len(self.db.results)):
                    num = i + 1
                    self.PortSpeedIDArray.append(self.db.results[i][0].strip())
                    self.PortSpeedNameArray.append(self.db.results[i][1].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][2].strip())
                    i = i + 1
                i = 0
                while (i < len(self.PortSpeedIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            self.PortSpeedIDArray[i],
                            self.PortSpeedNameArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]                    
                    self.PortSpeedTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_PortSpeed_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['PortSpeed','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.PortSpeedIDFrameEntry['state'] = 'normal'
        self.PortSpeedIDFrameEntry.delete(0,END)
        self.PortSpeedIDFrameEntry['state'] = 'readonly'
        self.PortSpeedIDFrameEntry['state'] = 'normal'
            
        self.PortSpeedNameFrameEntry['state'] = 'normal'
        self.PortSpeedNameFrameEntry.delete(0,END)
        self.PortSpeedNameFrameEntry['state'] = 'readonly'
        self.PortSpeedNameFrameEntry['state'] = 'normal'
        
        self.ButtonPortSpeedAdd['state']    = DISABLED
        self.ButtonPortSpeedEdit['state']   = DISABLED
        self.ButtonPortSpeedRemove['state'] = DISABLED
        self.ButtonPortSpeedOK['state']     = ACTIVE
        self.ButtonPortSpeedCancel['state'] = ACTIVE


    def Call_Button_PortSpeed_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['PortSpeed','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.PortSpeedIDFrameEntry['state'] = 'readonly'
        self.PortSpeedNameFrameEntry['state'] = 'normal'

        self.ButtonPortSpeedAdd['state']    = DISABLED
        self.ButtonPortSpeedEdit['state']   = DISABLED
        self.ButtonPortSpeedRemove['state'] = DISABLED
        self.ButtonPortSpeedOK['state']     = ACTIVE
        self.ButtonPortSpeedCancel['state'] = ACTIVE


    def Call_Button_PortSpeed_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['PortSpeed','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.PortSpeedFrame,title='Port Speed',message = 'Are you Sure you want to Remove it?')):
                self.ButtonPortSpeedAdd['state']    = ACTIVE
                self.ButtonPortSpeedEdit['state']   = DISABLED
                self.ButtonPortSpeedRemove['state'] = DISABLED
                self.ButtonPortSpeedOK['state']     = DISABLED
                self.ButtonPortSpeedCancel['state'] = DISABLED                
                PortSpeedID = self.PortSpeedIDFrameEntry.get()
                PortSpeedName = self.PortSpeedNameFrameEntry.get()
                #print (PortSpeedPortSpeedID)
                if (PortSpeedID == "UNKNOWN"):
                    mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                        message = '*** The Port Speed ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM Circuits
                            WHERE Port_Speed = '%s'
                          """ % (PortSpeedID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                                    message = '*** The Port Speed ID you entered CANNOT be Removed, \r\n Becuase is used in Circuits ***')
                            self.on_PortSpeed_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Port_Speed
                                WHERE Port_Speed_ID = '%s'
                              """ % (PortSpeedID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Port_Speed WHERE Port_Speed_ID = '%s'" % (PortSpeedID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.PortSpeedFrame,title='Port Speed',
                                        message = '*** The Port Speed ID you entered was Removed ***')
                                self.on_PortSpeed_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                                        message = '*** Port Speed ID you entered was NOT Removed ***')
                                self.on_PortSpeed_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                            message = '*** The Port Speed ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_PortSpeed_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['PortSpeed','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonPortSpeedAdd['state']    = DISABLED
            self.ButtonPortSpeedEdit['state']   = DISABLED
            self.ButtonPortSpeedRemove['state'] = DISABLED
            self.ButtonPortSpeedOK['state']     = DISABLED 
            self.ButtonPortSpeedCancel['state'] = DISABLED 

            PortSpeedID = self.PortSpeedIDFrameEntry.get()
            PortSpeedName = self.PortSpeedNameFrameEntry.get()
            
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(PortSpeedID) > 0) and (len(PortSpeedName) > 0)):
                    sql = """
                            SELECT * FROM Port_Speed
                            WHERE Port_Speed_ID = '%s'
                          """ % (PortSpeedID)
                    if (self.db.Execute(sql)):
                        self.ButtonPortSpeedOK['state']     = ACTIVE 
                        self.ButtonPortSpeedCancel['state'] = ACTIVE
                        mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                        message = '*** The Port Speed ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Port_Speed(Port_Speed_ID, Description, Executed_by_UserID) \
                                           VALUES('%s', '%s', '%s')" % \
                                   (PortSpeedID, PortSpeedName,self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            mbox.showwarning(master=self.PortSpeedFrame,title='Port Speed',
                                message = '*** The Port Speed ID you entered was Added ***')
                            self.on_PortSpeed_Table_Refresh()
                        else:
                            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                                message = '*** The Port Speed ID you entered was NOT Added ***')
                            self.on_PortSpeed_Table_Refresh()
                else:
                    self.ButtonPortSpeedOK['state']     = ACTIVE 
                    self.ButtonPortSpeedCancel['state'] = ACTIVE
                    mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                        message = '*** The Port Speed ID and Description Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Port_Speed
                        WHERE Port_Speed_ID = '%s'
                      """ % (PortSpeedID)
                if (len(PortSpeedName) == 0):
                    self.ButtonPortSpeedOK['state']     = ACTIVE 
                    self.ButtonPortSpeedCancel['state'] = ACTIVE
                    self.ButtonPortSpeedTypeAdd['state'] = ACTIVE
                    self.ButtonPortSpeedTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                        message = '*** The Port Speed Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):
                        '''
                            Port_Speed_ID           CHAR(20) NOT NULL PRIMARY KEY,
                            Description             CHAR(50),
                            Executed_by_UserID      CHAR(20))"""
                        '''
                        sql = "UPDATE Port_Speed SET Description = '%s', Executed_by_UserID = '%s' WHERE Port_Speed_ID = '%s'" % (PortSpeedName,
                                                                                        self.Username, PortSpeedID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.PortSpeedFrame,title='Port Speed',
                                message = '*** The Port Speed ID you entered was Updated ***')
                            self.on_PortSpeed_Table_Refresh()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                                message = '*** The Port Speed you entered was NOT Upadted ***')
                            self.on_PortSpeed_Table_Refresh()               
                    else:
                        mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                        message = '*** The Port Speed you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.PortSpeedFrame,title='Port Speed',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_PortSpeed_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['PortSpeed','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonPortSpeedAdd['state']    = ACTIVE
        self.ButtonPortSpeedEdit['state']   = DISABLED
        self.ButtonPortSpeedRemove['state'] = DISABLED
        self.ButtonPortSpeedOK['state']     = DISABLED
        self.ButtonPortSpeedCancel['state'] = DISABLED

        self.PortSpeedIDFrameEntry['state'] = 'normal'
        self.PortSpeedIDFrameEntry.delete(0,END)
        self.PortSpeedIDFrameEntry['state'] = 'readonly'
            
        self.PortSpeedNameFrameEntry['state'] = 'normal'
        self.PortSpeedNameFrameEntry.delete(0,END)
        self.PortSpeedNameFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_PortSpeed_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.PortSpeedTreeview.focus()
        dic = self.PortSpeedTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonPortSpeedAdd['state']    = ACTIVE
            self.ButtonPortSpeedEdit['state']   = ACTIVE
            self.ButtonPortSpeedRemove['state'] = ACTIVE
            self.ButtonPortSpeedOK['state']     = DISABLED
            self.ButtonPortSpeedCancel['state'] = DISABLED

            self.PortSpeedIDFrameEntry['state'] = 'normal'
            self.PortSpeedIDFrameEntry.delete(0,END)
            self.PortSpeedIDFrameEntry.insert(0,self.PortSpeedIDArray[curItem])
            self.PortSpeedIDFrameEntry['state'] = 'readonly'
            
            self.PortSpeedNameFrameEntry['state'] = 'normal'
            self.PortSpeedNameFrameEntry.delete(0,END)
            self.PortSpeedNameFrameEntry.insert(0,self.PortSpeedNameArray[curItem])
            self.PortSpeedNameFrameEntry['state'] = 'readonly'


    def Display_PortSpeed_Window(self):
        global PortSpeedLabel
        if not self.PortSpeedWindowExist:
            self.PortSpeedWindowExist = True
            # Set up the Window
            self.PortSpeedWindow = Tk()
            self.PortSpeedWindow.geometry('1300x650+350+70')
            self.PortSpeedWindow.title("Port Speeds")
            self.PortSpeedWindow.protocol("WM_DELETE_WINDOW", self.on_PortSpeedWindow_quit)
            self.PortSpeedWindow.call('tk', 'scaling', self.Windows_Scaling)
            PortSpeedLabel = Label(self.PortSpeedWindow,text="Helvetica", font=("Helvetica", 19))
            PortSpeedLabel["text"] = "Port Speeds"
            PortSpeedLabel.pack()
                    
            # Setup Frame
            self.PortSpeedFrame = Frame(self.PortSpeedWindow)
            self.PortSpeedFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.PortSpeedFrame.rowconfigure(0, weight=1)
            self.PortSpeedFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    PortSpeedIDFrameLabel = Label(self.PortSpeedFrame,text="Helvetica", font=("Helvetica", 10))
                    PortSpeedIDFrameLabel["text"] = "Port Speed:"
                    PortSpeedIDFrameLabel.place(x=10, y=10)                    
                    self.PortSpeedIDFrameEntry = Entry(self.PortSpeedFrame)
                    self.PortSpeedIDFrameEntry['width']=40
                    self.PortSpeedIDFrameEntry.place(x=160, y=10)
                    self.PortSpeedIDFrameEntry['state'] = 'readonly'
                    
                    PortSpeedNameFrameLabel = Label(self.PortSpeedFrame,text="Helvetica", font=("Helvetica", 10))
                    PortSpeedNameFrameLabel["text"] = "Description:"
                    PortSpeedNameFrameLabel.place(x=10, y=40)
                    self.PortSpeedNameFrameEntry = Entry(self.PortSpeedFrame)
                    self.PortSpeedNameFrameEntry['width']=40
                    self.PortSpeedNameFrameEntry.place(x=160, y=40)
                    self.PortSpeedNameFrameEntry['state'] = 'readonly'                    

                   
                    # Setup Buttons
                    self.ButtonPortSpeedAdd = Button(self.PortSpeedFrame, text = 'Add', command = self.Call_Button_PortSpeed_Add, state=DISABLED)
                    self.ButtonPortSpeedAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonPortSpeedAdd['state'] = ACTIVE

                    self.ButtonPortSpeedEdit = Button(self.PortSpeedFrame, text = 'Edit', command = self.Call_Button_PortSpeed_Edit, state=DISABLED)
                    self.ButtonPortSpeedEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonPortSpeedRemove = Button(self.PortSpeedFrame, text = 'Remove', command = self.Call_Button_PortSpeed_Remove, state=DISABLED)
                    self.ButtonPortSpeedRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonPortSpeedOK = Button(self.PortSpeedFrame, text = 'OK / UPDATE', command = self.Call_Button_PortSpeed_OK, state=DISABLED)
                    self.ButtonPortSpeedOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonPortSpeedCancel = Button(self.PortSpeedFrame, text = 'Cancel', command = self.Call_Button_PortSpeed_Cancel, state=DISABLED)
                    self.ButtonPortSpeedCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.PortSpeedFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For PortSpeed Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.PortSpeedTreeviewDataColumns = ('Port Speed','Description','Added/Updated By')
                    
                    self.PortSpeedTreeview = ttk.Treeview(self.PortSpeedFrame,columns=self.PortSpeedTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
                    self.PortSpeedTreeviewysb = Scrollbar(self.PortSpeedFrame,orient=VERTICAL, command=self.PortSpeedTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.PortSpeedTreeviewxsb = Scrollbar(self.PortSpeedFrame,orient=HORIZONTAL, command=self.PortSpeedTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.PortSpeedTreeview['yscroll'] = self.PortSpeedTreeviewysb.set
                    self.PortSpeedTreeview['xscroll'] = self.PortSpeedTreeviewxsb.set

                    # setup headings and column 
                    self.PortSpeedTreeview.heading('#0', text='Item No.',    anchor=W) # E for East and W for West
                    self.PortSpeedTreeview.heading('#1', text='Port Speed',   anchor=W,command=lambda: self.treeview_sort_column(self.PortSpeedTreeview, 'Port Speed', False)) # E for East and W for West
                    self.PortSpeedTreeview.heading('#2', text='Description',        anchor=W,command=lambda: self.treeview_sort_column(self.PortSpeedTreeview, 'Description', False)) 
                    self.PortSpeedTreeview.heading('#3', text='Added/Updated By',    anchor=W,command=lambda: self.treeview_sort_column(self.PortSpeedTreeview, 'Added/Updated By', False)) 
                    
                    self.PortSpeedTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.PortSpeedTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.PortSpeedTreeview.column('#2', stretch=1, width=10)
                    self.PortSpeedTreeview.column('#3', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.PortSpeedTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.PortSpeedTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.PortSpeedTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.PortSpeedTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.PortSpeedTreeview.bind('<ButtonRelease-1>', self.on_PortSpeed_Tree_select_click) # When Select the Tree
                    #self.PortSpeedTreeview.bind("<Double-1>", self.On_PortSpeed_Tree_Refresh)         

                    # SQL Querry to the PortSpeed Table
                    self.data_ready = False
                    self.on_PortSpeed_Table_Refresh()

                    #------------------ MENU <BEGIN> ----------------------------------------------------------
                    menubar = Menu(self.PortSpeedWindow)
                    filemenu = Menu(menubar, tearoff=0)
                    menubar.add_cascade(label="File", menu=filemenu)
                    if Is_SaveAs_Available:
                        filemenu.add_command(label="Save As", command=self.Call_Save_As)
                    filemenu.add_separator()
                    filemenu.add_command(label="Exit", command=self.on_PortSpeedWindow_quit)
                    self.PortSpeedWindow.config(menu=menubar)
                    #------------------ MENU <END> ----------------------------------------------------------
                        
                    self.PortSpeedWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.PortSpeedFrame,title='PortSpeed',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the PortSpeed Class....:")
    PortSpeed = Class_PortSpeed("BV",Windows_Scaling)
    PortSpeed.Display_PortSpeed_Window()


if __name__ == '__main__':
    Main()

