#==========================================================================
# Program: CircuitType.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: April-22-2018
# Date Last Modified: April-22-2018
# Summary: This is Class to for the CircuitType
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
    Parameter = ['CircuitType','OPEN Window']    
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

class Class_CircuitType:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.CircuitTypeWindowExist = False
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
                Tab0 = "Circuit Type List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.CircuitTypeTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.CircuitTypeIDArray)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.CircuitTypeIDArray[i],
                            self.CircuitTypeNameArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.CircuitTypeFrame,title='Circuit Type List',
                            message = '!!! The File was saved !!!')
                    self.CircuitTypeWindow.title("Circuit Type                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['CircuitType','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def on_CircuitTypeWindow_quit(self):
        if (self.CircuitTypeWindowExist):
            self.CircuitTypeWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['CircuitType','CLOSE Window']    
                Logging.Log(Parameter) 
            self.CircuitTypeWindow.destroy()

    def on_CircuitType_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():

            self.CircuitTypeIDFrameEntry['state'] = 'readonly'
            self.CircuitTypeNameFrameEntry['state'] = 'readonly'
            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.CircuitTypeTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.CircuitTypeTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                            Circuit_Type            CHAR(40) NOT NULL PRIMARY KEY,
                            Description             CHAR(50),
                            Executed_by_UserID      CHAR(20))"""
            '''
            # SQL Querry to the CircuitType Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            # SQL Querry to the CircuitType Table
            self.data_ready = False
            sql = """
                        SELECT * FROM Circuit_Type
                        ORDER BY Description ASC
                  """
            if (self.db.Execute(sql)):
                self.data_ready = True
                self.CircuitTypeIDArray = []
                self.CircuitTypeNameArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                '''
                    0 Circuit_Type            CHAR(40) NOT NULL PRIMARY KEY,
                    1 Description             CHAR(50),
                    2 Executed_by_UserID      CHAR(20))"""
                '''
                while (i < len(self.db.results)):
                    num = i + 1
                    self.CircuitTypeIDArray.append(self.db.results[i][0].strip())
                    self.CircuitTypeNameArray.append(self.db.results[i][1].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][2].strip())
                    i = i + 1
                i = 0
                while (i < len(self.CircuitTypeIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire
                    item = [
                            self.CircuitTypeIDArray[i],
                            self.CircuitTypeNameArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]                    
                    self.CircuitTypeTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
            else:
                mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                message = 'No Records found')
        else:
            mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_CircuitType_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['CircuitType','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.CircuitTypeIDFrameEntry['state'] = 'normal'
        self.CircuitTypeIDFrameEntry.delete(0,END)
        self.CircuitTypeIDFrameEntry['state'] = 'readonly'
        self.CircuitTypeIDFrameEntry['state'] = 'normal'
            
        self.CircuitTypeNameFrameEntry['state'] = 'normal'
        self.CircuitTypeNameFrameEntry.delete(0,END)
        self.CircuitTypeNameFrameEntry['state'] = 'readonly'
        self.CircuitTypeNameFrameEntry['state'] = 'normal'
        
        self.ButtonCircuitTypeAdd['state']    = DISABLED
        self.ButtonCircuitTypeEdit['state']   = DISABLED
        self.ButtonCircuitTypeRemove['state'] = DISABLED
        self.ButtonCircuitTypeOK['state']     = ACTIVE
        self.ButtonCircuitTypeCancel['state'] = ACTIVE


    def Call_Button_CircuitType_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['CircuitType','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.CircuitTypeIDFrameEntry['state'] = 'readonly'
        self.CircuitTypeNameFrameEntry['state'] = 'normal'

        self.ButtonCircuitTypeAdd['state']    = DISABLED
        self.ButtonCircuitTypeEdit['state']   = DISABLED
        self.ButtonCircuitTypeRemove['state'] = DISABLED
        self.ButtonCircuitTypeOK['state']     = ACTIVE
        self.ButtonCircuitTypeCancel['state'] = ACTIVE


    def Call_Button_CircuitType_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['CircuitType','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.CircuitTypeFrame,title='Circuit Type',message = 'Are you Sure you want to Remove it?')):
                self.ButtonCircuitTypeAdd['state']    = ACTIVE
                self.ButtonCircuitTypeEdit['state']   = DISABLED
                self.ButtonCircuitTypeRemove['state'] = DISABLED
                self.ButtonCircuitTypeOK['state']     = DISABLED
                self.ButtonCircuitTypeCancel['state'] = DISABLED                
                CircuitTypeID = self.CircuitTypeIDFrameEntry.get()
                CircuitTypeName = self.CircuitTypeNameFrameEntry.get()
                #print (CircuitTypeCircuitTypeID)
                if (CircuitTypeID == "UNKNOWN"):
                    mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                        message = '*** The Circuit Type ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First in Region --------
                    sql = """
                            SELECT * FROM Circuits
                            WHERE Circuit_Type = '%s'
                          """ % (CircuitTypeID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                                    message = '*** The Circuit Type ID you entered CANNOT be Removed, \r\n Becuase is used in Circuits ***')
                            self.on_CircuitType_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Circuit_Type
                                WHERE Circuit_Type = '%s'
                              """ % (CircuitTypeID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Circuit_Type WHERE Circuit_Type = '%s'" % (CircuitTypeID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.CircuitTypeFrame,title='Circuit Type',
                                        message = '*** The Circuit Type ID you entered was Removed ***')
                                self.on_CircuitType_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                                        message = '*** The Circuit Type ID you entered was NOT Removed ***')
                                self.on_CircuitType_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                            message = '*** The Circuit Type ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_CircuitType_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['CircuitType','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonCircuitTypeAdd['state']    = DISABLED
            self.ButtonCircuitTypeEdit['state']   = DISABLED
            self.ButtonCircuitTypeRemove['state'] = DISABLED
            self.ButtonCircuitTypeOK['state']     = DISABLED 
            self.ButtonCircuitTypeCancel['state'] = DISABLED 

            CircuitTypeID = self.CircuitTypeIDFrameEntry.get()
            CircuitTypeName = self.CircuitTypeNameFrameEntry.get()
            
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(CircuitTypeID) > 0) and (len(CircuitTypeName) > 0)):
                    sql = """
                            SELECT * FROM Circuit_Type
                            WHERE Circuit_Type = '%s'
                          """ % (CircuitTypeID)
                    if (self.db.Execute(sql)):
                        self.ButtonCircuitTypeOK['state']     = ACTIVE 
                        self.ButtonCircuitTypeCancel['state'] = ACTIVE
                        mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                        message = '*** The Circuit Type ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Circuit_Type(Circuit_Type, Description, Executed_by_UserID) \
                                           VALUES('%s', '%s', '%s')" % \
                                   (CircuitTypeID, CircuitTypeName,self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            mbox.showwarning(master=self.CircuitTypeFrame,title='Circuit Type',
                                message = '*** The Circuit Type ID you entered was Added ***')
                            self.on_CircuitType_Table_Refresh()
                        else:
                            mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                                message = '*** The Circuit Type ID you entered was NOT Added ***')
                            self.on_CircuitType_Table_Refresh()
                else:
                    self.ButtonCircuitTypeOK['state']     = ACTIVE 
                    self.ButtonCircuitTypeCancel['state'] = ACTIVE
                    mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                        message = '*** The Circuit Type and Description Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Circuit_Type
                        WHERE Circuit_Type = '%s'
                      """ % (CircuitTypeID)
                if (len(CircuitTypeName) == 0):
                    self.ButtonCircuitTypeOK['state']     = ACTIVE 
                    self.ButtonCircuitTypeCancel['state'] = ACTIVE
                    self.ButtonCircuitTypeTypeAdd['state'] = ACTIVE
                    self.ButtonCircuitTypeTypeRefresh['state'] = ACTIVE
                    mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                        message = '*** The Circuit Type Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):
                        '''
                            Circuit_Type            CHAR(40) NOT NULL PRIMARY KEY,
                            Description             CHAR(50),
                            Executed_by_UserID      CHAR(20))"""
                        '''
                        sql = "UPDATE Circuit_Type SET Description = '%s', Executed_by_UserID = '%s' WHERE Circuit_Type = '%s'" % (CircuitTypeName,
                                                                                        self.Username, CircuitTypeID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.CircuitTypeFrame,title='Circuit Type',
                                message = '*** The Circuit Type ID you entered was Updated ***')
                            self.on_CircuitType_Table_Refresh()
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                                message = '*** The Circuit Type you entered was NOT Upadted ***')
                            self.on_CircuitType_Table_Refresh()               
                    else:
                        mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                        message = '*** The Circuit Type you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.CircuitTypeFrame,title='Circuit Type',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_CircuitType_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['CircuitType','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonCircuitTypeAdd['state']    = ACTIVE
        self.ButtonCircuitTypeEdit['state']   = DISABLED
        self.ButtonCircuitTypeRemove['state'] = DISABLED
        self.ButtonCircuitTypeOK['state']     = DISABLED
        self.ButtonCircuitTypeCancel['state'] = DISABLED

        self.CircuitTypeIDFrameEntry['state'] = 'normal'
        self.CircuitTypeIDFrameEntry.delete(0,END)
        self.CircuitTypeIDFrameEntry['state'] = 'readonly'
            
        self.CircuitTypeNameFrameEntry['state'] = 'normal'
        self.CircuitTypeNameFrameEntry.delete(0,END)
        self.CircuitTypeNameFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_CircuitType_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.CircuitTypeTreeview.focus()
        dic = self.CircuitTypeTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonCircuitTypeAdd['state']    = ACTIVE
            self.ButtonCircuitTypeEdit['state']   = ACTIVE
            self.ButtonCircuitTypeRemove['state'] = ACTIVE
            self.ButtonCircuitTypeOK['state']     = DISABLED
            self.ButtonCircuitTypeCancel['state'] = DISABLED

            self.CircuitTypeIDFrameEntry['state'] = 'normal'
            self.CircuitTypeIDFrameEntry.delete(0,END)
            self.CircuitTypeIDFrameEntry.insert(0,self.CircuitTypeIDArray[curItem])
            self.CircuitTypeIDFrameEntry['state'] = 'readonly'
            
            self.CircuitTypeNameFrameEntry['state'] = 'normal'
            self.CircuitTypeNameFrameEntry.delete(0,END)
            self.CircuitTypeNameFrameEntry.insert(0,self.CircuitTypeNameArray[curItem])
            self.CircuitTypeNameFrameEntry['state'] = 'readonly'


    def Display_CircuitType_Window(self):
        global CircuitTypeLabel
        if not self.CircuitTypeWindowExist:
            self.CircuitTypeWindowExist = True
            # Set up the Window
            self.CircuitTypeWindow = Tk()
            self.CircuitTypeWindow.geometry('1300x650+350+70')
            self.CircuitTypeWindow.title("Circuit Types")
            self.CircuitTypeWindow.protocol("WM_DELETE_WINDOW", self.on_CircuitTypeWindow_quit)
            self.CircuitTypeWindow.call('tk', 'scaling', self.Windows_Scaling)
            CircuitTypeLabel = Label(self.CircuitTypeWindow,text="Helvetica", font=("Helvetica", 19))
            CircuitTypeLabel["text"] = "Circuit Types"
            CircuitTypeLabel.pack()
                    
            # Setup Frame
            self.CircuitTypeFrame = Frame(self.CircuitTypeWindow)
            self.CircuitTypeFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.CircuitTypeFrame.rowconfigure(0, weight=1)
            self.CircuitTypeFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    CircuitTypeIDFrameLabel = Label(self.CircuitTypeFrame,text="Helvetica", font=("Helvetica", 10))
                    CircuitTypeIDFrameLabel["text"] = "Circuit Type:"
                    CircuitTypeIDFrameLabel.place(x=10, y=10)                    
                    self.CircuitTypeIDFrameEntry = Entry(self.CircuitTypeFrame)
                    self.CircuitTypeIDFrameEntry['width']=40
                    self.CircuitTypeIDFrameEntry.place(x=160, y=10)
                    self.CircuitTypeIDFrameEntry['state'] = 'readonly'
                    
                    CircuitTypeNameFrameLabel = Label(self.CircuitTypeFrame,text="Helvetica", font=("Helvetica", 10))
                    CircuitTypeNameFrameLabel["text"] = "Description:"
                    CircuitTypeNameFrameLabel.place(x=10, y=40)
                    self.CircuitTypeNameFrameEntry = Entry(self.CircuitTypeFrame)
                    self.CircuitTypeNameFrameEntry['width']=40
                    self.CircuitTypeNameFrameEntry.place(x=160, y=40)
                    self.CircuitTypeNameFrameEntry['state'] = 'readonly'                    

                   
                    # Setup Buttons
                    self.ButtonCircuitTypeAdd = Button(self.CircuitTypeFrame, text = 'Add', command = self.Call_Button_CircuitType_Add, state=DISABLED)
                    self.ButtonCircuitTypeAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonCircuitTypeAdd['state'] = ACTIVE

                    self.ButtonCircuitTypeEdit = Button(self.CircuitTypeFrame, text = 'Edit', command = self.Call_Button_CircuitType_Edit, state=DISABLED)
                    self.ButtonCircuitTypeEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonCircuitTypeRemove = Button(self.CircuitTypeFrame, text = 'Remove', command = self.Call_Button_CircuitType_Remove, state=DISABLED)
                    self.ButtonCircuitTypeRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonCircuitTypeOK = Button(self.CircuitTypeFrame, text = 'OK / UPDATE', command = self.Call_Button_CircuitType_OK, state=DISABLED)
                    self.ButtonCircuitTypeOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonCircuitTypeCancel = Button(self.CircuitTypeFrame, text = 'Cancel', command = self.Call_Button_CircuitType_Cancel, state=DISABLED)
                    self.ButtonCircuitTypeCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.CircuitTypeFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For CircuitType Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.CircuitTypeTreeviewDataColumns = ('Circuit Type','Description','Added/Updated By')
                    
                    self.CircuitTypeTreeview = ttk.Treeview(self.CircuitTypeFrame,columns=self.CircuitTypeTreeviewDataColumns, height=25) # <--- Make sure the frame is correct !!
                    self.CircuitTypeTreeviewysb = Scrollbar(self.CircuitTypeFrame,orient=VERTICAL, command=self.CircuitTypeTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.CircuitTypeTreeviewxsb = Scrollbar(self.CircuitTypeFrame,orient=HORIZONTAL, command=self.CircuitTypeTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.CircuitTypeTreeview['yscroll'] = self.CircuitTypeTreeviewysb.set
                    self.CircuitTypeTreeview['xscroll'] = self.CircuitTypeTreeviewxsb.set

                    # setup headings and column 
                    self.CircuitTypeTreeview.heading('#0', text='Item No.',    anchor=W) # E for East and W for West
                    self.CircuitTypeTreeview.heading('#1', text='Circuit Type',   anchor=W,command=lambda: self.treeview_sort_column(self.CircuitTypeTreeview, 'Circuit Type', False)) # E for East and W for West
                    self.CircuitTypeTreeview.heading('#2', text='Description',        anchor=W,command=lambda: self.treeview_sort_column(self.CircuitTypeTreeview, 'Description', False)) 
                    self.CircuitTypeTreeview.heading('#3', text='Added/Updated By',    anchor=W,command=lambda: self.treeview_sort_column(self.CircuitTypeTreeview, 'Added/Updated By', False)) 
                    
                    self.CircuitTypeTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.CircuitTypeTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.CircuitTypeTreeview.column('#2', stretch=1, width=10)
                    self.CircuitTypeTreeview.column('#3', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.CircuitTypeTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.CircuitTypeTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.CircuitTypeTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.CircuitTypeTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.CircuitTypeTreeview.bind('<ButtonRelease-1>', self.on_CircuitType_Tree_select_click) # When Select the Tree
                    #self.CircuitTypeTreeview.bind("<Double-1>", self.On_CircuitType_Tree_Refresh)         

                    # SQL Querry to the CircuitType Table
                    self.data_ready = False
                    self.on_CircuitType_Table_Refresh()

                    #------------------ MENU <BEGIN> ----------------------------------------------------------
                    menubar = Menu(self.CircuitTypeWindow)
                    filemenu = Menu(menubar, tearoff=0)
                    menubar.add_cascade(label="File", menu=filemenu)
                    if Is_SaveAs_Available:
                        filemenu.add_command(label="Save As", command=self.Call_Save_As)
                    filemenu.add_separator()
                    filemenu.add_command(label="Exit", command=self.on_CircuitTypeWindow_quit)
                    self.CircuitTypeWindow.config(menu=menubar)
                    #------------------ MENU <END> ----------------------------------------------------------
                        
                    self.CircuitTypeWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.CircuitTypeFrame,title='CircuitType',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the CircuitType Class....:")
    CircuitType = Class_CircuitType("BV",Windows_Scaling)
    CircuitType.Display_CircuitType_Window()


if __name__ == '__main__':
    Main()

