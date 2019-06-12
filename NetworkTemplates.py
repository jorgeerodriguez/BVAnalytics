#==========================================================================
# Program: NetworkTemplates.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: May-21-2018
# Date Last Modified: May-21-2018
# Summary: This is Class to for the NetworkTemplates
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
    Parameter = ['Network Teamplates','OPEN Window']    
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

class Class_NetworkTemplates:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.NetworkTemplatesWindowExist = False
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


    def on_NetworkTemplatesWindow_quit(self):
        if (self.NetworkTemplatesWindowExist):
            self.NetworkTemplatesWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Network Teamplates','CLOSE Window']    
                Logging.Log(Parameter) 
            self.NetworkTemplatesWindow.destroy()

    def on_NetworkTemplates_Table_Refresh(self):
        #time.sleep(2)
        if self.db.Connect():            
            #------------------------------- Deleting Tree View <BEGIN> --------
            x = self.NetworkTemplatesTreeview.get_children()
            if x != '()': # checks if there is something in the first row
                for child in x:
                    #print (child)
                    self.NetworkTemplatesTreeview.delete(child)
            #------------------------------- Deleting Tree View <END> --------
            '''
                        Template_ID             CHAR(100) NOT NULL PRIMARY KEY,
                        Template_Name           CHAR(100) NOT NULL,
                        Executed_by_UserID      CHAR(20) )"""
            '''

            # SQL Querry to the NetworkTemplates Table
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            sql = """
                        SELECT * FROM Network_Templates
                        ORDER BY Template_Name ASC
                  """
            if (self.db.Execute(sql)):                        
                self.NetworkTemplatesIDArray = []
                self.NetworkTemplatesNameArray = []
                self.ExecutedbyUserIDArray = []
                i = 0
                while (i < len(self.db.results)):
                    num = i + 1
                    self.NetworkTemplatesIDArray.append(self.db.results[i][0].strip())
                    self.NetworkTemplatesNameArray.append(self.db.results[i][1].strip())
                    self.ExecutedbyUserIDArray.append(self.db.results[i][4].strip())
                    i = i + 1
                i = 0
                while (i < len(self.NetworkTemplatesIDArray)):
                    num = i + 1
                    tags = "-" # To use in the futire                          
                    item = [
                            self.NetworkTemplatesIDArray[i],
                            self.NetworkTemplatesNameArray[i],
                            self.ExecutedbyUserIDArray[i]
                            ]
                    self.NetworkTemplatesTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1
                    self.progress['value'] = i
            else:
                mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                message = 'No Records found')
            #self.db.Disconnect()
        else:
            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")

        
    def Call_Button_NetworkTemplates_Add(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','ADD Button']    
            Logging.Log(Parameter)

        #-- reset the progess bar --
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'add'

        self.NetworkTeamplateIDFrameEntry['state'] = 'normal'
        self.NetworkTeamplateIDFrameEntry.delete(0,END)
        self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
        self.NetworkTeamplateIDFrameEntry['state'] = 'normal'
            
        self.NetworkTeamplateNameFrameEntry['state'] = 'normal'
        self.NetworkTeamplateNameFrameEntry.delete(0,END)
        self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'
        self.NetworkTeamplateNameFrameEntry['state'] = 'normal'

        self.ButtonNetworkTemplatesAdd['state']    = DISABLED
        self.ButtonNetworkTemplatesEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesOK['state']     = ACTIVE
        self.ButtonNetworkTemplatesCancel['state'] = ACTIVE


    def Call_Button_NetworkTemplates_Edit(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','EDIT Button']    
            Logging.Log(Parameter)
            
        self.progress['maximum'] = 100
        self.progress['value'] = 0
        self.Selection = 'edit'
        
        self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
        self.NetworkTeamplateNameFrameEntry['state'] = 'normal'

        self.ButtonNetworkTemplatesAdd['state']    = DISABLED
        self.ButtonNetworkTemplatesEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesOK['state']     = ACTIVE
        self.ButtonNetworkTemplatesCancel['state'] = ACTIVE


    def Call_Button_NetworkTemplates_Remove(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','REMOVE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.progress['maximum'] = 100
            self.progress['value'] = 0
            self.Selection = 'remove'
            
            if (mbox.askyesnocancel(master=self.NetworkTemplatesFrame,title='Network Teamplates',message = 'Are you Sure you want to Remove it?')):
                self.ButtonNetworkTemplatesAdd['state']    = ACTIVE
                self.ButtonNetworkTemplatesEdit['state']   = DISABLED
                self.ButtonNetworkTemplatesRemove['state'] = DISABLED
                self.ButtonNetworkTemplatesOK['state']     = DISABLED
                self.ButtonNetworkTemplatesCancel['state'] = DISABLED                
                NetworkTemplatesID = self.NetworkTeamplateIDFrameEntry.get()
                NetworkTemplatesName = self.NetworkTeamplateNameFrameEntry.get()
                if (NetworkTemplatesID == "UNKNOWN"):
                    mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = '*** The Network Teamplates ID UNKNOWN CANNOT be Removed by Design ***')
                else:
                    #--------------- Validate First Network to make sure the template is not used some place else --------
                    sql = """
                            SELECT * FROM Device_Model
                            WHERE Template_ID = '%s'
                          """ % (NetworkTemplatesID)
                    if (self.db.Execute(sql)):
                            #self.db.Disconnect()
                            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                    message = '*** The Network Teamplates ID you entered CANNOT be Removed, \r\n Becuase is used in a Device Model ***')
                            self.on_NetworkTemplates_Table_Refresh()
                    else:
                        sql = """
                                SELECT * FROM Network_Templates
                                WHERE Template_ID = '%s'
                              """ % (NetworkTemplatesID)
                        if (self.db.Execute(sql)):
                            sql = "DELETE FROM Network_Templates WHERE Template_ID = '%s'" % (NetworkTemplatesID)
                            if (self.db.Add_Move_Change_Data(sql)):
                                #self.db.Disconnect()
                                mbox.showwarning(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                        message = '*** The Network Teamplates ID you entered was Removed ***')
                                self.on_NetworkTemplates_Table_Refresh()
                            else:
                                #self.db.Disconnect()
                                mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                        message = '*** The Network Teamplate ID you entered was NOT Removed ***')
                                self.on_NetworkTemplates_Table_Refresh()               
                        else:
                            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                            message = '*** The Network Teamplate ID you try to Remove Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_NetworkTemplates_OK(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','OK/UPDATE Button']    
            Logging.Log(Parameter)
            
        if self.db.Connect():
            self.ButtonNetworkTemplatesAdd['state']    = DISABLED
            self.ButtonNetworkTemplatesEdit['state']   = DISABLED
            self.ButtonNetworkTemplatesRemove['state'] = DISABLED
            self.ButtonNetworkTemplatesOK['state']     = DISABLED 
            self.ButtonNetworkTemplatesCancel['state'] = DISABLED 

            NetworkTemplatesID = self.NetworkTeamplateIDFrameEntry.get()
            NetworkTemplatesName = self.NetworkTeamplateNameFrameEntry.get()
            #-------------- ADD ----------------------
            if (self.Selection == 'add'):
                if ((len(NetworkTemplatesID) > 0) and (len(NetworkTemplatesName) > 0)):
                    sql = """
                            SELECT * FROM Network_Templates
                            WHERE Template_ID = '%s'
                          """ % (NetworkTemplatesID)
                    if (self.db.Execute(sql)):
                        self.ButtonNetworkTemplatesOK['state']     = ACTIVE 
                        self.ButtonNetworkTemplatesCancel['state'] = ACTIVE
                        mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = '*** The Network Teamplates ID you entered already exist ***')
                    else:
                        sql = "INSERT INTO Network_Templates(Template_ID, \
                                   Template_Name, Base_Network_Slash, Split_in_Slash, Executed_by_UserID) \
                                   VALUES ('%s', '%s', '%d', '%d','%s')" % \
                                   (NetworkTemplatesID, NetworkTemplatesName, 0, 0, self.Username)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                message = '*** The Network Teamplate ID you entered was Added ***')
                            self.on_NetworkTemplates_Table_Refresh()
                            self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
                            self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'
                        else:
                            #print ("Error adding the record, posible dupliated it")
                            #self.db.Disconnect()
                            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                message = '*** The Network Teamplate ID you entered was NOT Added ***')
                            self.on_NetworkTemplates_Table_Refresh()
                            self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
                            self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'
                else:
                    self.ButtonNetworkTemplatesOK['state']     = ACTIVE 
                    self.ButtonNetworkTemplatesCancel['state'] = ACTIVE
                    mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = '*** The Network Teamplate ID and Network Teamplate Name Cannot be BLANK ***') 
            #-------------- EDIT ----------------------
            if (self.Selection == 'edit'):
                sql = """
                        SELECT * FROM Network_Templates
                        WHERE Template_ID = '%s'
                      """ % (NetworkTemplatesID)
                if (len(NetworkTemplatesName) == 0):
                    self.ButtonNetworkTemplatesOK['state']     = ACTIVE 
                    self.ButtonNetworkTemplatesCancel['state'] = ACTIVE
                    mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = '*** The Network Teamplate Name Cannot be BLANK ***')                     
                else:
                    if (self.db.Execute(sql)):

                        sql = "UPDATE Network_Templates SET Template_Name = '%s', Executed_by_UserID = '%s' WHERE Template_ID = '%s'" % (NetworkTemplatesName,
                                                                                 self.Username,NetworkTemplatesID)
                        if (self.db.Add_Move_Change_Data(sql)):
                            #self.db.Disconnect()
                            mbox.showwarning(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                message = '*** The Network Teamplate ID you entered was Updated ***')
                            self.on_NetworkTemplates_Table_Refresh()
                            self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'
                            
                        else:
                            #self.db.Disconnect()
                            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                message = '*** The Network Teamplate ID you entered was NOT Upadted ***')
                            self.on_NetworkTemplates_Table_Refresh()               
                            self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'
                    else:
                        mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = '*** The NetworkTeamplate ID you try to Edit Does not exist Anymore ***')
        else:
            mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")


    def Call_Button_NetworkTemplates_Cancel(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','CANCEL Button']    
            Logging.Log(Parameter)
            
        self.ButtonNetworkTemplatesAdd['state']    = ACTIVE
        self.ButtonNetworkTemplatesEdit['state']   = DISABLED
        self.ButtonNetworkTemplatesRemove['state'] = DISABLED
        self.ButtonNetworkTemplatesOK['state']     = DISABLED
        self.ButtonNetworkTemplatesCancel['state'] = DISABLED

        self.NetworkTeamplateIDFrameEntry['state'] = 'normal'
        self.NetworkTeamplateIDFrameEntry.delete(0,END)
        self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
            
        self.NetworkTeamplateNameFrameEntry['state'] = 'normal'
        self.NetworkTeamplateNameFrameEntry.delete(0,END)
        self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'

        self.Selection = 'cancel'

    def on_NetworkTemplates_Tree_select_click(self,event):
        #print ("Select")
        curItem = self.NetworkTemplatesTreeview.focus()
        dic = self.NetworkTemplatesTreeview.item(curItem)
        values = dic.get('values')        
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1
            #print (values)
            #print (values[0])
            #print (values[1])
            self.ButtonNetworkTemplatesAdd['state']    = ACTIVE
            self.ButtonNetworkTemplatesEdit['state']   = ACTIVE
            self.ButtonNetworkTemplatesRemove['state'] = ACTIVE
            self.ButtonNetworkTemplatesOK['state']     = DISABLED
            self.ButtonNetworkTemplatesCancel['state'] = DISABLED

            self.NetworkTeamplateIDFrameEntry['state'] = 'normal'
            self.NetworkTeamplateIDFrameEntry.delete(0,END)
            self.NetworkTeamplateIDFrameEntry.insert(0,str(values[0]).strip())
            self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
            
            self.NetworkTeamplateNameFrameEntry['state'] = 'normal'
            self.NetworkTeamplateNameFrameEntry.delete(0,END)
            self.NetworkTeamplateNameFrameEntry.insert(0,str(values[1]).strip())
            self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'


    def On_NetworkTemplates_Tree_Refresh(self,event):
        #--- Double Click --
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Network Teamplates','REFRESH Button']    
            Logging.Log(Parameter)            
        self.on_NetworkTemplates_Table_Refresh()

    
    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Network Teamplates List"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.NetworkTemplatesTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                                  #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.NetworkTemplatesFrame,title='Network Teamplates List',
                            message = '!!! The File was saved !!!')
                    self.NetworkTemplatesWindow.title("Network Teamplates                    File: ["+self.file_name+"] SAVED") #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Network Teamplates','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates List',
                            message = '*** The File was not saved, Perhaps It is already open ***')



    def Display_NetworkTemplates_Window(self):
        global NetworkTemplatesLabel
        if not self.NetworkTemplatesWindowExist:
            self.NetworkTemplatesWindowExist = True
            # Set up the Window
            self.NetworkTemplatesWindow = Tk()
            self.NetworkTemplatesWindow.geometry('1000x650+350+70')
            self.NetworkTemplatesWindow.title("Network Teamplates")
            self.NetworkTemplatesWindow.protocol("WM_DELETE_WINDOW", self.on_NetworkTemplatesWindow_quit)
            self.NetworkTemplatesWindow.call('tk', 'scaling', self.Windows_Scaling)
            NetworkTemplatesLabel = Label(self.NetworkTemplatesWindow,text="Helvetica", font=("Helvetica", 19))
            NetworkTemplatesLabel["text"] = "Network Teamplates"
            NetworkTemplatesLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.NetworkTemplatesWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_NetworkTemplatesWindow_quit)
            self.NetworkTemplatesWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                    
            # Setup Frame
            self.NetworkTemplatesFrame = Frame(self.NetworkTemplatesWindow)
            self.NetworkTemplatesFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.NetworkTemplatesFrame.rowconfigure(0, weight=1)
            self.NetworkTemplatesFrame.columnconfigure(0, weight=1)
            if self.db.Connect():        
                    # Setup Labels
                    NetworkTemplatesIDFrameLabel = Label(self.NetworkTemplatesFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesIDFrameLabel["text"] = "Teamplate ID:"
                    NetworkTemplatesIDFrameLabel.place(x=10, y=10)
                    self.NetworkTeamplateIDFrameEntry = Entry(self.NetworkTemplatesFrame)
                    self.NetworkTeamplateIDFrameEntry['width']=50
                    self.NetworkTeamplateIDFrameEntry.place(x=140, y=10)
                    self.NetworkTeamplateIDFrameEntry['state'] = 'readonly'
                    
                    NetworkTemplatesNameFrameLabel = Label(self.NetworkTemplatesFrame,text="Helvetica", font=("Helvetica", 10))
                    NetworkTemplatesNameFrameLabel["text"] = "Teamplate Name:"
                    NetworkTemplatesNameFrameLabel.place(x=10, y=40)
                    self.NetworkTeamplateNameFrameEntry = Entry(self.NetworkTemplatesFrame)
                    self.NetworkTeamplateNameFrameEntry['width']=50
                    self.NetworkTeamplateNameFrameEntry.place(x=140, y=40)
                    self.NetworkTeamplateNameFrameEntry['state'] = 'readonly'                    
                   
                    # Setup Buttons
                    self.ButtonNetworkTemplatesAdd = Button(self.NetworkTemplatesFrame, text = 'Add', command = self.Call_Button_NetworkTemplates_Add, state=DISABLED)
                    self.ButtonNetworkTemplatesAdd.place(x = 450, y = 7, width=75, height=25)
                    self.ButtonNetworkTemplatesAdd['state'] = ACTIVE

                    self.ButtonNetworkTemplatesEdit = Button(self.NetworkTemplatesFrame, text = 'Edit', command = self.Call_Button_NetworkTemplates_Edit, state=DISABLED)
                    self.ButtonNetworkTemplatesEdit.place(x = 550, y = 7, width=75, height=25)

                    self.ButtonNetworkTemplatesRemove = Button(self.NetworkTemplatesFrame, text = 'Remove', command = self.Call_Button_NetworkTemplates_Remove, state=DISABLED)
                    self.ButtonNetworkTemplatesRemove.place(x = 650, y = 7, width=75, height=25)

                    self.ButtonNetworkTemplatesOK = Button(self.NetworkTemplatesFrame, text = 'OK / UPDATE', command = self.Call_Button_NetworkTemplates_OK, state=DISABLED)
                    self.ButtonNetworkTemplatesOK.place(x = 750, y = 7, width=100, height=25)

                    self.ButtonNetworkTemplatesCancel = Button(self.NetworkTemplatesFrame, text = 'Cancel', command = self.Call_Button_NetworkTemplates_Cancel, state=DISABLED)
                    self.ButtonNetworkTemplatesCancel.place(x = 875, y = 7, width=75, height=25)

                    # Create Progress Bar
                    self.progress = ttk.Progressbar(self.NetworkTemplatesFrame, orient="horizontal",length=500, mode="determinate")
                    self.progress['maximum'] = 100
                    self.progress['value'] = 0
                    self.progress.place(x=450, y=38)


                    #------------------ TREE VIEW For NetworkTemplates Database <BEGIN> -----------------------------------
                    # Create Tree and Scrollbars
                    self.NetworkTemplatesTreeviewDataColumns = ('Network Teamplates ID','Network Teamplates Name','Added By')
                    
                    self.NetworkTemplatesTreeview = ttk.Treeview(self.NetworkTemplatesFrame,columns=self.NetworkTemplatesTreeviewDataColumns, height=24) # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesTreeviewysb = Scrollbar(self.NetworkTemplatesFrame,orient=VERTICAL, command=self.NetworkTemplatesTreeview.yview)           # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesTreeviewxsb = Scrollbar(self.NetworkTemplatesFrame,orient=HORIZONTAL, command=self.NetworkTemplatesTreeview.xview)         # <--- Make sure the frame is correct !!
                    self.NetworkTemplatesTreeview['yscroll'] = self.NetworkTemplatesTreeviewysb.set
                    self.NetworkTemplatesTreeview['xscroll'] = self.NetworkTemplatesTreeviewxsb.set

                    # setup headings and column 
                    self.NetworkTemplatesTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West
                    self.NetworkTemplatesTreeview.heading('#1', text='Network Teamplates ID',   anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesTreeview, 'Network Teamplates ID', False)) # E for East and W for West
                    self.NetworkTemplatesTreeview.heading('#2', text='Network Teamplates Name', anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesTreeview, 'Network Teamplates Name', False)) 
                    self.NetworkTemplatesTreeview.heading('#3', text='Added By',     anchor=W,command=lambda: self.treeview_sort_column(self.NetworkTemplatesTreeview, 'Added By', False)) 
                    
                    self.NetworkTemplatesTreeview.column('#0', stretch=1, width=3 , anchor=W)
                    self.NetworkTemplatesTreeview.column('#1', stretch=1, width=10, anchor=W)
                    self.NetworkTemplatesTreeview.column('#2', stretch=1, width=10)
                    self.NetworkTemplatesTreeview.column('#3', stretch=1, width=10)

                    # add tree and scrollbars to frame
                    self.NetworkTemplatesTreeview.grid(row=1, column=0, sticky=NSEW)
                    self.NetworkTemplatesTreeviewysb.grid(row=1, column=1, sticky=NS)
                    self.NetworkTemplatesTreeviewxsb.grid(row=2, column=0, sticky=EW)


                    # create fonts and tags
                    self.NetworkTemplatesTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                    # Bind the double Click
                    self.NetworkTemplatesTreeview.bind('<ButtonRelease-1>', self.on_NetworkTemplates_Tree_select_click) # When Select the Tree
                    self.NetworkTemplatesTreeview.bind("<Double-1>", self.On_NetworkTemplates_Tree_Refresh)         

                    # SQL Querry to the NetworkTemplates Table
                    sql = """
                                SELECT * FROM Network_Templates
                                ORDER BY Template_Name ASC
                              """
                    if (self.db.Execute(sql)):                        
                        self.NetworkTemplatesIDArray = []
                        self.NetworkTemplatesNameArray = []
                        self.ExecutedbyUserIDArray = []
                        self.results = []
                        i = 0
                        self.data_ready = True
                        while (i < len(self.db.results)):
                            num = i + 1
                            self.NetworkTemplatesIDArray.append(self.db.results[i][0].strip())
                            self.NetworkTemplatesNameArray.append(self.db.results[i][1].strip())
                            self.ExecutedbyUserIDArray.append(self.db.results[i][4].strip())
                            i = i + 1
                        i = 0
                        while (i < len(self.NetworkTemplatesIDArray)):
                            num = i + 1
                            tags = "-" # To use in the futire                          
                            item = [
                                    self.NetworkTemplatesIDArray[i],
                                    self.NetworkTemplatesNameArray[i],
                                    self.ExecutedbyUserIDArray[i]
                                    ]
                            self.results.append(item)
                            self.NetworkTemplatesTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                        message = 'No Records found')
                        self.data_ready = False
                    #------------------ TREE VIEW For NetworkTemplates Database <END> -----------------------------------
                    #self.db.Disconnect()
                    self.NetworkTemplatesWindow.mainloop()
                #else:
                    #self.db.Disconnect()
            else:
                mbox.showerror(master=self.NetworkTemplatesFrame,title='Network Teamplates',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + self.ODBC_name + "]")

def Main():
    print ("Testing the NetworkTemplates Class....:")
    NetworkTemplates = Class_NetworkTemplates("BV",Windows_Scaling)
    NetworkTemplates.Display_NetworkTemplates_Window()


if __name__ == '__main__':
    Main()

