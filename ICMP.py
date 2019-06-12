#==========================================================================
# Program: ICMP.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-21-2018
# Date Last Modified: Jan-21-2018
# Summary: This is Class to for the ICMP
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
from threading import Thread
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

try:
    from Logging import *
    Is_logging_Available = True
    Parameter = []
    Parameter = ['ICMP','OPEN Window']    
    Logging = Class_Logging(ODBC_DSN_name,Parameter)
    Logging.Log(Parameter)
except:
    print ("********************************************************************************** \n")
    print ("*** NO Logging Library Found, please download it in order to access the Databases *** \n")
    print ("********************************************************************************** \n")
    Is_logging_Available = False

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


class Class_ICMP:

    def __init__(self,DSN_Name,Windows_Scaling,Location):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPWindowExist = False
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.Go_To_Location = False
        if (len(Location) > 0):
            self.Init_SQL_Field = Location[0]
            self.Init_Value = Location[1]
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


    def on_ICMPWindow_quit(self):
        if (self.ICMPWindowExist):
            self.ICMPWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['ICMP','CLOSE Window']    
                Logging.Log(Parameter)
            self.ICMPWindow.destroy()

    def Traceroute_Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Traceroute Times History"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.TracerouteTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results3)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results3[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.TracerouteFrame,title='Traceroute Response Times History List',
                            message = '!!! The File was saved !!!')
                    self.TracerouteWindow .title("Traceroute Response Times History                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['Traceroute Response Times History','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.TracerouteFrame,title='Traceroute Response Times History List',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Display_Traceroute_Window(self,event):
        if self.db.Connect():
            # Set up the Window
            self.TracerouteWindow = Tk()
            self.TracerouteWindow.geometry('1350x410+350+70')
            self.TracerouteWindow.title("Traceroute Response Times History")
            self.TracerouteWindow.call('tk', 'scaling', Windows_Scaling)
            self.TracerouteLabel = Label(self.TracerouteWindow,text="Helvetica", font=("Helvetica", 13))
            self.TracerouteLabel["text"] = "Traceroute Response Times History"
            self.TracerouteLabel.pack()

            self.data_ready = False
            #------------------ MENU <BEGIN> ----------------------------------------------------------
            menubar = Menu(self.TracerouteWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Traceroute_Call_Save_As)
            #filemenu.add_separator()
            #filemenu.add_command(label="Exit", command=self.on_VendorWindow_quit)
            self.TracerouteWindow.config(menu=menubar)
            #------------------ MENU <END> ----------------------------------------------------------
                
            # Setup Frame
            self.TracerouteFrame = Frame(self.TracerouteWindow)
            self.TracerouteFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.TracerouteFrame.rowconfigure(0, weight=1)
            self.TracerouteFrame.columnconfigure(0, weight=1)

            #------------------ TREE VIEW For Traceroute Database <BEGIN> -----------------------------------
            # Create Tree and Scrollbars
            self.TracerouteTreeviewDataColumns = ('Date',
                               'Time',
                               'IP Address',
                               'Hop No',
                               'Next Hop',
                               '1st. Response (ms)',
                               '2nd. Response (ms)',
                               '3rd. Response (ms)')
                
            self.TracerouteTreeview = ttk.Treeview(self.TracerouteFrame,columns=self.TracerouteTreeviewDataColumns, height=16)     # <--- Make sure the frame is correct !!
            self.TracerouteTreeviewysb = Scrollbar(self.TracerouteFrame,orient=VERTICAL, command=self.TracerouteTreeview.yview)    # <--- Make sure the frame is correct !!
            self.TracerouteTreeviewxsb = Scrollbar(self.TracerouteFrame,orient=HORIZONTAL, command=self.TracerouteTreeview.xview)        # <--- Make sure the frame is correct !!
            self.TracerouteTreeview['yscroll'] = self.TracerouteTreeviewysb.set
            self.TracerouteTreeview['xscroll'] = self.TracerouteTreeviewxsb.set

            # setup headings and column 
            self.TracerouteTreeview.heading('#0', text='Item No.',           anchor=W) # E for East and W for West
            self.TracerouteTreeview.heading('#1', text='Date',               anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, 'Date', False)) # E for East and W for West
            self.TracerouteTreeview.heading('#2', text='Time',               anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, 'Time', False)) 
            self.TracerouteTreeview.heading('#3', text='IP Address',         anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, 'IP Address', False))
            self.TracerouteTreeview.heading('#4', text='Hop No',             anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, 'Hop No', False))
            self.TracerouteTreeview.heading('#5', text='Next Hop',           anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, 'Next Hop', False))
            self.TracerouteTreeview.heading('#6', text='1st. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, '1st. Response (ms)', False))
            self.TracerouteTreeview.heading('#7', text='2nd. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, '2nd. Response (ms)', False))
            self.TracerouteTreeview.heading('#8', text='3rd. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.TracerouteTreeview, '3rd. Response (ms)', False))
                
            self.TracerouteTreeview.column('#0', stretch=1, width=3 , anchor=W)
            self.TracerouteTreeview.column('#1', stretch=1, width=10, anchor=W)
            self.TracerouteTreeview.column('#2', stretch=1, width=10)
            self.TracerouteTreeview.column('#3', stretch=1, width=20, anchor=W)
            self.TracerouteTreeview.column('#4', stretch=1, width=3)
            self.TracerouteTreeview.column('#5', stretch=1, width=20)
            self.TracerouteTreeview.column('#6', stretch=1, width=20)
            self.TracerouteTreeview.column('#7', stretch=1, width=20)
            self.TracerouteTreeview.column('#8', stretch=1, width=20)

            # add tree and scrollbars to frame
            self.TracerouteTreeview.grid(row=1, column=0, sticky=NSEW)
            self.TracerouteTreeviewysb.grid(row=1, column=1, sticky=NS)
            self.TracerouteTreeviewxsb.grid(row=2, column=0, sticky=EW)

            # create fonts and tags for future use to compare Region and delay
            self.TracerouteTreeview.tag_configure('ATT', font=('Helvetica', 8), background='LightBlue1')
            self.TracerouteTreeview.tag_configure('SPRINT', font=('Helvetica', 8), background='coral1')
            self.TracerouteTreeview.tag_configure('Asterisc', font=('Helvetica', 8), background='pink')
                
            # Bind the double Click
            #self.TracerouteTreeview.bind('<ButtonRelease-1>', on_ICMP_Tree_select_click)
            #self.TracerouteTreeview.bind("<Double-1>", on_ICMP_Tree_double_click)

            curItem = self.ICMPtreeview.focus()
            dic = self.ICMPtreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                #print (values)
                date = values[0]
                time = values[1]            
                ipaddress = values[2]
                # SQL Querry to the ICMP Table
                if self.db.Connect():
                    sql = """
                            SELECT * FROM TRACEROUTE 
                            WHERE Date_String = '%s' AND Time_String = '%s' AND Device_IP = '%s' ORDER BY Hop_No ASC
                          """ % (date,time,ipaddress)
                    if (self.db.Execute(sql)):
                        i = 0
                        self.data_ready = True
                        self.results3 = []
                        while (i < len(self.db.results)):
                            num = i + 1
                            tags = "regions" # more work on in for RTTD > 250 ms.
                            if ((self.db.results[i][12].find("*",0) != -1) or (self.db.results[i][13].find("*",0) != -1) or (self.db.results[i][14].find("*",0) != -1)):
                                tags = "Asterisc"
                            if ((self.db.results[i][11].find("10.50.",0) != -1)): 
                                tags = "ATT"
                            if ((self.db.results[i][11].find("10.64.",0) != -1)):
                                tags = "SPRINT"
                                
                            item = [self.db.results[i][2],
                                    self.db.results[i][3],
                                    self.db.results[i][1],
                                    self.db.results[i][4],
                                    self.db.results[i][11],
                                    self.db.results[i][13],
                                    self.db.results[i][14],
                                    self.db.results[i][12]]
                            self.results3.append(item)
                            self.TracerouteTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.TracerouteFrame,title='Traceroute Response Time History',
                        message = 'No Records found')
                        self.data_ready = False
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.TracerouteFrame,title='ICMP Response Time History',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")            
                #db.Disconnect()
                self.TracerouteWindow.mainloop()
            #else:
                #db.Disconnect()
        else:
            mbox.showerror(master=root,title='ICMP Response Time History',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + ODBC_DSN_name + "]")


    def on_ICMP_Tree_select_click(self,event):    # <- Good for Testing new functions !!!!!
            #print ("Select Click ICMP ......")
            #print (self.tree.get_children())
            #iid = event.widget.focus() # <--------- it Works
            #iid = tree.focus()
            #print (tree.index(iid))
            #curItem = tree.focus() # <---- this works to but tree need to be global variable
            #print (event.x)
            curItem = event.widget.focus()
            #print (tree.item(curItem)) # The Data is in Dictionary format
            # dic = tree.item(curItem) <---- this works to but tree need to be global variable
            dic = event.widget.item(curItem)
            #print (dic)
            #print (dic.get('values'))
            values = dic.get('values')
            #print (values)
            #print (len(values))
            #if (len(values) > 0):
                #print ("repose time")
                #print ('IP Address => ' + values[7])
            self.buttonICMPLOG['state'] = ACTIVE
            self.buttonTraceroute['state'] = ACTIVE




    def on_ICMP_Tree_Device_select_click(self,event):
            #print ("Select Click Device.....")
            #print (self.tree.get_children())
            #iid = event.widget.focus() # <--------- it Works
            #iid = event.widget.focus()
            #print (event.widget.index(iid))
            #curItem = tree.focus() # <---- this works to but tree need to be global variable
            #print (event.x)
            curItem = event.widget.focus()
            #print (tree.item(curItem)) # The Data is in Dictionary format
            # dic = tree.item(curItem) <---- this works to but tree need to be global variable
            dic = event.widget.item(curItem)
            #print (dic)
            #print (dic.get('values'))
            values = dic.get('values')
            #print (values)
            #print (len(values))
            #if (len(values) > 0):
            #    print ('IP Address => ' + values[7])
            self.buttonICMPLOG['state'] = ACTIVE
            self.buttonTraceroute['state'] = DISABLED

    def on_ICMP_Tree_Device_double_click(self,event):
           
            curItem = self.ICMPtreeviewDevice.focus()
            dic = self.ICMPtreeviewDevice.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                ipaddress = values[7]
                self.results2 = []
                # SQL Querry to the ICMP Table
                if self.db.Connect():
                    sql = """SELECT * FROM ICMP 
                            WHERE Device_IP = '%s' ORDER BY Year desc, Month desc, Day desc,
                            Hour desc, Minute desc, Second desc , Size_of_Ping desc """ % (ipaddress)
                    if (self.db.Execute(sql)):
                        #------------------------------- Deleting Begin --------
                        x = self.ICMPtreeview.get_children()
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.ICMPtreeview.delete(child)
                        #------------------------------- Deleting End --------
                        i = 0
                        while (i < len(self.db.results)):
                            num = i + 1
                            tags = "regions" # more work on in for RTTD > 250 ms.
                            item = [self.db.results[i][2],
                                    self.db.results[i][3],
                                    self.db.results[i][1],
                                    self.db.results[i][10],
                                    self.db.results[i][11],
                                    self.db.results[i][13],
                                    self.db.results[i][14],
                                    self.db.results[i][12]]
                            self.results2.append(item)
                            self.ICMPtreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.ICMPFrame,title='ICMP Response Time History',
                        message = 'No Records found')
                    #db.Disconnect()
                else:
                    mbox.showerror(master=self.ICMPFrame,title='ICMP Response Time History',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")

    def Call_ButtonICMPLOG(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['ICMP','ICMP LOG']    
            Logging.Log(Parameter)
        self.on_ICMP_Tree_Device_double_click('event')

    def Call_ButtonTreaceroute(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['ICMP','TRACEROUTE Window']    
            Logging.Log(Parameter)
        self.Display_Traceroute_Window('event')

    def Call_ButtonFind(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['ICMP','FIND Button']    
            Logging.Log(Parameter)        
        #print ("Execute SQL")
        value = self.ICMPFrameEntry.get() # Value to querry
        index = (self.ComboBoxICMPTableFields.current())
        #field = self.ComboBoxICMPTableFields.get())
        field = (self.DeviceTableFieldsSQL[index])
        # I need to work on this querry today!!!!
        #data = Build_fake_data()
        #Populate_tree(self.ICMPtreeviewDevice,data)
        '''
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
        '''
        self.buttonICMPLOG['state'] = DISABLED
        self.buttonTraceroute['state'] = DISABLED

        if (len(value) > 0):
                if self.db.Connect():
                    ## SELECT * FROM Customers ORDER BY Country ASC, CustomerName DESC
                    if (self.Go_To_Location):
                        sql = """SELECT * FROM DEVICES 
                                WHERE %s = '%s' ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Device_ID, IP4_Address, Status """ % (field,value)
                    else:
                        sql = """SELECT * FROM DEVICES 
                                WHERE %s LIKE '%s' ORDER BY Region_ID,Country_ID, Facility_ID, Site_ID,
                                Device_ID, IP4_Address, Status """ % (field,('%'+value+'%'))
                    #print (sql)
                    if (self.db.Execute(sql)):
                        #print ("No of Rows:"+str(len(self.db.results)))
                        #print (self.db.results)
                        #print (self.db.results[0][2])

                        self.results = []
                        self.results2 = []
                        
                        #------------------------------- Deleting ICMP Begin --------
                        x = self.ICMPtreeview.get_children()
                        #print ('get_children values: ', x ,'\n')
                        #self.ICMPtreeview.delete(*tree.get_children())
                        #self.ICMPtreeview.delete(x)
                        #self.ICMPtreeview.destroy() <- removes the widget
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.ICMPtreeview.delete(child)
                        #------------------------------- Deleting ICMP End --------

                        #------------------------------- Deleting ICMP Device Begin --------
                        x = self.ICMPtreeviewDevice.get_children()
                        #print ('get_children values: ', x ,'\n')
                        #self.ICMPtreeview.delete(*tree.get_children())
                        #self.ICMPtreeview.delete(x)
                        #self.ICMPtreeview.destroy() <- removes the widget
                        if x != '()': # checks if there is something in the first row
                            for child in x:
                                #print (child)
                                self.ICMPtreeviewDevice.delete(child)
                        #------------------------------- Deleting ICMP Device End --------
                        self.data_ready = True
                        i = 0
                        while (i < len(self.db.results)):
                            num = i + 1
                            item = [self.db.results[i][0],
                                    self.db.results[i][1],
                                    self.db.results[i][2],
                                    self.db.results[i][3],
                                    self.db.results[i][4],
                                    self.db.results[i][5],
                                    self.db.results[i][6],
                                    self.db.results[i][7],
                                    self.db.results[i][12]]
                            #print (item)
                            #print (self.db.results[i][12])
                            self.results.append(item)
                            if (item[8] == None):
                                tags = "None"
                            else:
                                tags = item[8]
                            #print ("tag-> "+tag)
                            self.ICMPtreeviewDevice.insert('', END, text='%3d'%num, values=item, tags=tags)
                            i = i + 1
                    else:
                        mbox.showerror(master=self.ICMPFrame,title='ICMP Response Time History',
                        message = 'No Records found')
                        self.data_ready = False
                    #self.db.Disconnect()
                else:
                    mbox.showerror(master=self.ICMPFrame,title='ICMP Response Time History',
                    message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")
        else:
            mbox.showerror(master=self.ICMPFrame,title='ICMP Response Time History',
            message = '*** ERROR *** - The Vlaue Field Cannot be Blank - ***, \r\n')            

        
    def Call_ButtonFind_From_Entry(self,event):
        self.Call_ButtonFind()

    def ICMP_Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Device List"
                Tab1 = "ICMP Log"
                Tabs = [Tab0,Tab1]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.ICMPtreeviewDeviceDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                ExcellFile.Add_DataToWorksheet(self.ICMPtreeviewDataColumns,Row,Column,Tab1,"Fill",lightsalmon_1,14,'Bold')   #<--- Columns
                Row = 3
                i = 0
                while (i < len(self.results)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                Row = 3
                i = 0
                while (i < len(self.results2)):                                                                            #<------ Data for the Columns
                    num = i + 1
                    item = self.results2[i]                          
                    ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab1,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1

                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ICMPFrame,title='ICMP and Traceroute Response Times History List',
                            message = '!!! The File was saved !!!')
                    self.ICMPWindow.title("ICMP and Traceroute Response Times History                    File: ["+self.file_name+"] SAVED")                  #<---- Window Name Change
                    if Is_logging_Available:
                        Parameter = []
                        Parameter = ['ICMP and Traceroute Response Times History','SAVE AS']    
                        Logging.Log(Parameter)
                else:
                    mbox.showerror(master=self.ICMPFrame,title='ICMP and Traceroute Response Times History List',
                            message = '*** The File was not saved, Perhaps It is already open ***')

    def Display_ICMP_Window(self):
        if self.db.Connect():
            if not self.ICMPWindowExist:
                # Set up the Window
                self.ICMPWindowExist = True
                self.ICMPWindow = Tk()
                self.ICMPWindow.geometry('1350x700+350+70')
                self.ICMPWindow.title("ICMP and Traceroute Response Times History")
                self.ICMPWindow.protocol("WM_DELETE_WINDOW", self.on_ICMPWindow_quit)
                self.ICMPWindow.call('tk', 'scaling', Windows_Scaling)
                self.ICMPLabel = Label(self.ICMPWindow,text="Helvetica", font=("Helvetica", 13))
                self.ICMPLabel["text"] = "ICMP and Traceroute Response Times History"
                self.ICMPLabel.pack()

                self.data_ready = False
                #------------------ MENU <BEGIN> ----------------------------------------------------------
                menubar = Menu(self.ICMPWindow)
                filemenu = Menu(menubar, tearoff=0)
                menubar.add_cascade(label="File", menu=filemenu)
                if Is_SaveAs_Available:
                    filemenu.add_command(label="Save As", command=self.ICMP_Call_Save_As)
                filemenu.add_separator()
                filemenu.add_command(label="Exit", command=self.on_ICMPWindow_quit)
                self.ICMPWindow.config(menu=menubar)
                #------------------ MENU <END> ----------------------------------------------------------
                
                # Setup Frame
                self.ICMPFrame = Frame(self.ICMPWindow)
                self.ICMPFrame.pack(side=TOP, fill=BOTH, expand=Y)

                # set frame resizing priorities
                self.ICMPFrame.rowconfigure(0, weight=1)
                self.ICMPFrame.columnconfigure(0, weight=1)

                # Setup Labels
                self.ICMPFrameLabel = Label(self.ICMPFrame,text="Helvetica", font=("Helvetica", 10))
                self.ICMPFrameLabel["text"] = "Find:"
                self.ICMPFrameLabel.place(x=10, y=10)

                # Setup ComboBox
                self.DeviceTableFields = ['Device Name',
                                'Description',
                                'Country',
                                'Region',
                                'Facility',
                                'Site',
                                'Device Type',
                                'IP Address',
                                'Status']
                self.DeviceTableFieldsSQL = ['Device_ID',
                                'Device_Description',
                                'Country_ID',
                                'Region_ID',
                                'Facility_ID',
                                'Site_ID',
                                'Device_Type_ID',
                                'IP4_Address',
                                'Status']

                self.ComboBoxICMPTableFields = ttk.Combobox(self.ICMPFrame, state='readonly', width = 25)
                self.ComboBoxICMPTableFields['values'] = self.DeviceTableFields
                self.ComboBoxICMPTableFields.current(0)
                self.ComboBoxICMPTableFields.place(x = 50, y = 10)
                #ComboBoxButtonWEBOSName['state'] = DISABLED

                # Setup Entry
                self.ICMPFrameEntry = Entry(self.ICMPFrame)
                self.ICMPFrameEntry['width']=50
                self.ICMPFrameEntry.place(x=235, y=10)
                self.ICMPFrameEntry.bind('<Return>', self.Call_ButtonFind_From_Entry)
                #self.ICMPFrameEntry['state'] = DISABLED
                
                # Setup Button
                self.buttonFind = Button(self.ICMPFrame, text = 'Execute', command = self.Call_ButtonFind, state=ACTIVE)
                self.buttonFind.place(x = 550, y = 7, width=75, height=25)

                self.buttonICMPLOG = Button(self.ICMPFrame, text = 'ICMP LOG', command = self.Call_ButtonICMPLOG, state=DISABLED)
                self.buttonICMPLOG.place(x = 650, y = 7, width=75, height=25)

                self.buttonTraceroute = Button(self.ICMPFrame, text = 'Traceroute LOG', command = self.Call_ButtonTreaceroute, state=DISABLED)
                self.buttonTraceroute.place(x = 750, y = 7, width=100, height=25)


                '''
                # This is s List Box setup wiht Scrollbars that I can use as a template!!!
                # Setup ListBox and the Scrollbars
                self.ICMPListBox = Listbox(self.ICMPFrame, width=120, height=15, font=("Helvetica", 9))        
                yscrollbar = Scrollbar(self.ICMPFrame, orient=VERTICAL, command=self.ICMPListBox.yview)
                xscrollbar = Scrollbar(self.ICMPFrame, orient=HORIZONTAL, command=self.ICMPListBox.xview )
                self.ICMPListBox['yscroll'] = yscrollbar.set
                self.ICMPListBox['xscroll'] = xscrollbar.set
                self.ICMPListBox.grid(row=1, column=0, sticky=NSEW)
                yscrollbar.grid(row=1, column=1, sticky=NS)
                xscrollbar.grid(row=2, column=0, sticky=EW)
                # ---- Need to do more research on Scrollbar with place function ---
                #self.ICMPListBox.place(x=10, y=20)
                #yscrollbar.place(x=1000, y=20)
                #xscrollbar.place(x=10, y=100)

                # Building Fake Data for testing
                for n in range(40):
                    self.ICMPListBox.insert(END,"Opening Database....."+str(n))

                self.ICMPListBox.select_set(0)
                self.ICMPListBox.bind("<Double-Button-1>",OnDouble_Click_ICMP) # Double Clcik
    #            self.ICMPListBox.bind("<<ListboxSelect>>", OnSelect_ICMP(self.ICMPListBox))      # Single Click or Keys up and down too.
    #            self.ICMPListBox.bind("<<ListboxSelect>>", OnSelect('icmp'))      # Single Click or Keys up and down too.

                '''

                #------------------ TREE VIEW For ICMP Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.ICMPtreeviewDataColumns = ('Date',
                               'Time',
                               'IP Address',
                               'Ping Size',
                               '% Loss',
                               'Min. Response (ms)',
                               'Ave. Response (ms)',
                               'Max. Response (ms)')
                
                self.ICMPtreeview = ttk.Treeview(self.ICMPFrame,columns=self.ICMPtreeviewDataColumns, height=12) # <--- Make sure the frame is correct !!
                self.ICMPtreeviewysb = Scrollbar(self.ICMPFrame,orient=VERTICAL, command=self.ICMPtreeview.yview)           # <--- Make sure the frame is correct !!
                self.ICMPtreeviewxsb = Scrollbar(self.ICMPFrame,orient=HORIZONTAL, command=self.ICMPtreeview.xview)         # <--- Make sure the frame is correct !!
                self.ICMPtreeview['yscroll'] = self.ICMPtreeviewysb.set
                self.ICMPtreeview['xscroll'] = self.ICMPtreeviewxsb.set

                # setup headings and column 
                self.ICMPtreeview.heading('#0', text='Item No.',           anchor=W) # E for East and W for West
                self.ICMPtreeview.heading('#1', text='Date',               anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Date', False)) # E for East and W for West
                self.ICMPtreeview.heading('#2', text='Time',               anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Time', False)) 
                self.ICMPtreeview.heading('#3', text='IP Address',         anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'IP Address', False))
                self.ICMPtreeview.heading('#4', text='Ping Size',          anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Ping Size', False))
                self.ICMPtreeview.heading('#5', text='% Loss',             anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, '% Loss', False))
                self.ICMPtreeview.heading('#6', text='Min. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Min. Response (ms)', False))
                self.ICMPtreeview.heading('#7', text='Ave. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Ave. Response (ms)', False))
                self.ICMPtreeview.heading('#8', text='Max. Response (ms)', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeview, 'Max. Response (ms)', False))
                
                self.ICMPtreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.ICMPtreeview.column('#1', stretch=1, width=10, anchor=W)
                self.ICMPtreeview.column('#2', stretch=1, width=10)
                self.ICMPtreeview.column('#3', stretch=1, width=20, anchor=W)
                self.ICMPtreeview.column('#4', stretch=1, width=10)
                self.ICMPtreeview.column('#5', stretch=1, width=10)
                self.ICMPtreeview.column('#6', stretch=1, width=20)
                self.ICMPtreeview.column('#7', stretch=1, width=20)
                self.ICMPtreeview.column('#8', stretch=1, width=20)

                # add tree and scrollbars to frame
                self.ICMPtreeview.grid(row=1, column=0, sticky=NSEW)
                self.ICMPtreeviewysb.grid(row=1, column=1, sticky=NS)
                self.ICMPtreeviewxsb.grid(row=2, column=0, sticky=EW)


                # create fonts and tags
                self.ICMPtreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                #self.ICMPtreeview.tag_configure('production', font=('Helvetica', 8), background='green2')
                self.ICMPtreeview.tag_configure('obsoleted', font=('Helvetica', 8), background='ivory3') # 'red'
                self.ICMPtreeview.tag_configure('obsolete', font=('Helvetica', 8), background='ivory3')  # 'red'
                self.ICMPtreeview.tag_configure('implementation', font=('Helvetica', 8), background='green3')
                self.ICMPtreeview.tag_configure('pre-production', font=('Helvetica', 8), background='yellow2')
                self.ICMPtreeview.tag_configure('up', font=('Helvetica', 8), background='green2')
                self.ICMPtreeview.tag_configure('down', font=('Helvetica', 8), background='pink')
                

                # Bind the double Click
                self.ICMPtreeview.bind('<ButtonRelease-1>', self.on_ICMP_Tree_select_click)
                self.ICMPtreeview.bind("<Double-1>", self.Display_Traceroute_Window)

                #data = Build_fake_data()
                #Populate_tree(tree,data)
            
                #------------------ TREE VIEW For Device Database <ENDS> -----------------------------------


                
                #------------------ TREE VIEW For Device Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.ICMPtreeviewDeviceDataColumns = ('Device Name',
                               'Description',
                               'Country',
                               'Region',
                               'Facility',
                               'Site',
                               'Device Type',
                               'IP Address',
                               'Status')
                
                self.ICMPtreeviewDevice = ttk.Treeview(self.ICMPFrame,columns=self.ICMPtreeviewDeviceDataColumns, height=15)    # <--- Make sure the frame is correct !!
                self.ICMPtreeviewDeviceysb = Scrollbar(self.ICMPFrame,orient=VERTICAL, command=self.ICMPtreeviewDevice.yview)      # <--- Make sure the frame is correct !!
                self.ICMPtreeviewDevicexsb = Scrollbar(self.ICMPFrame,orient=HORIZONTAL, command=self.ICMPtreeviewDevice.xview)    # <--- Make sure the frame is correct !!
                self.ICMPtreeviewDevice['yscroll'] = self.ICMPtreeviewDeviceysb.set
                self.ICMPtreeviewDevice['xscroll'] = self.ICMPtreeviewDevicexsb.set

                # setup headings and column 
                self.ICMPtreeviewDevice.heading('#0', text='Item No.',    anchor=W) # E for East and W for West
                self.ICMPtreeviewDevice.heading('#1', text='Device Name', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Device Name', False)) # E for East and W for West
                self.ICMPtreeviewDevice.heading('#2', text='Description', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Description', False)) 
                self.ICMPtreeviewDevice.heading('#3', text='Country',     anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Country', False))
                self.ICMPtreeviewDevice.heading('#4', text='Region',      anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Region', False))
                self.ICMPtreeviewDevice.heading('#5', text='Facility',    anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Facility', False))
                self.ICMPtreeviewDevice.heading('#6', text='Site',        anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Site', False))
                self.ICMPtreeviewDevice.heading('#7', text='Device Type', anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Device Type', False))
                self.ICMPtreeviewDevice.heading('#8', text='IP Address',  anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'IP Address', False))
                self.ICMPtreeviewDevice.heading('#9', text='Status',      anchor=W,command=lambda: self.treeview_sort_column(self.ICMPtreeviewDevice, 'Status', False))
                
                self.ICMPtreeviewDevice.column('#0', stretch=1, width=3 , anchor=W)
                self.ICMPtreeviewDevice.column('#1', stretch=1, width=20, anchor=W)
                self.ICMPtreeviewDevice.column('#2', stretch=1, width=50)
                self.ICMPtreeviewDevice.column('#3', stretch=1, width=30, anchor=W)
                self.ICMPtreeviewDevice.column('#4', stretch=1, width=30)
                self.ICMPtreeviewDevice.column('#5', stretch=1, width=30)
                self.ICMPtreeviewDevice.column('#6', stretch=1, width=30)
                self.ICMPtreeviewDevice.column('#7', stretch=1, width=20)
                self.ICMPtreeviewDevice.column('#8', stretch=1, width=30)
                self.ICMPtreeviewDevice.column('#9', stretch=1, width=15)

                # add tree and scrollbars to frame
                self.ICMPtreeviewDevice.grid(row=11, column=0, sticky=NSEW)
                self.ICMPtreeviewDeviceysb.grid(row=11, column=1, sticky=NS)
                self.ICMPtreeviewDevicexsb.grid(row=12, column=0, sticky=EW)

                # create fonts and tags
    #           normal   = tkfont.Font(family='Consolas', size=6,)
    #           self.ICMPtreeviewDevice.tag_configure('Connected', font=normal)
                self.ICMPtreeviewDevice.tag_configure('None', font=('Helvetica', 8), background='azure')
                #self.ICMPtreeviewDevice.tag_configure('production', font=('Helvetica', 8), background='SpringGreen2')
                self.ICMPtreeviewDevice.tag_configure('obsolete', font=('Helvetica', 8), background='ivory3') # 'OrangeRed2'
                self.ICMPtreeviewDevice.tag_configure('implementation', font=('Helvetica', 8), background='green3')
                self.ICMPtreeviewDevice.tag_configure('pre-production', font=('Helvetica', 8), background='yellow2')
                self.ICMPtreeviewDevice.tag_configure('unmanaged', font=('Helvetica', 8), background='plum1')
                self.ICMPtreeviewDevice.tag_configure('up', font=('Helvetica', 8), background='SpringGreen2')
                self.ICMPtreeviewDevice.tag_configure('down', font=('Helvetica', 8), background='pink')

                # Bind the double Click
                self.ICMPtreeviewDevice.bind('<ButtonRelease-1>', self.on_ICMP_Tree_Device_select_click)
                self.ICMPtreeviewDevice.bind("<Double-1>", self.on_ICMP_Tree_Device_double_click)

            
                #------------------ TREE VIEW For Device Database <ENDS> -----------------------------------
                #self.db.Disconnect()
                if (self.Go_To_Location):
                    i = 0
                    while (i < len(self.DeviceTableFields)):
                        if (self.DeviceTableFields[i] == self.Init_SQL_Field):
                            self.ComboBoxICMPTableFields.current(i)
                            self.ICMPFrameEntry.insert(0,self.Init_Value)
                            self.Call_ButtonFind()
                            i = len(self.DeviceTableFields) + 5
                        else:
                            i = i + 1
                self.ICMPWindow.mainloop()
        else:
            mbox.showerror(master=root,title='ICMP Response Time History',
                             message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + ODBC_DSN_name + "]")


def Main():
    print ("Testing the ICMP Class....:")
    Location = []
    #Location = ['IP Address','10.1.1.1']
    ICMP = Class_ICMP("BV",Windows_Scaling,Location)
    ICMP.Display_ICMP_Window()


if __name__ == '__main__':
    Main()
