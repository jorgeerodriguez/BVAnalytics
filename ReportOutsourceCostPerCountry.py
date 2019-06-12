#==========================================================================
# Program: ReportOutsourcePerCountry.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: March-21-2018
# Date Last Modified: MAr-31-2018
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
    Parameter = ['Report Outsource Cost Per Country','OPEN Window']    
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


#*********************************************************************************************************************************************
#                                   ReportOutsourcePerCountry Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************
class Class_ReportOutsourcePerCountry:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPSummaryWindowExist = False
        self.FindCircuitWindowExist = False
        self.FindDeviceWindowExist = False
        self.CircuitsWindowExist = False
        self.ReportOutsourcePerCountryWindowExist = False
        self.Username = os.getlogin()
        self.date = ""
        self.Windows_Scaling = Windows_Scaling
        self.data_ready = False

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


    def on_ReportOutsourcePerCountryWindow_quit(self):
        if (self.ReportOutsourcePerCountryWindowExist):
            self.ReportOutsourcePerCountryWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Report Outsource Cost Per Country','CLOSED Window']    
                Logging.Log(Parameter)
            self.ReportOutsourcePerCountryWindow.destroy()


    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.ReportOutsourcePerCountryTreeview.focus()
        dic = self.ReportOutsourcePerCountryTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.ReportOutsourcePerCountryTreeview.next(curItem)
        self.ReportOutsourcePerCountryTreeview.selection_set((NextItem, NextItem))
        self.ReportOutsourcePerCountryTreeview.focus(NextItem)

        NextItem = self.ReportOutsourcePerCountryTreeview.focus()
        dic = self.ReportOutsourcePerCountryTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
            
    def on_ReportOutsourcePerCountry_Tree_select_click(self,event):
            curItem = self.ReportOutsourcePerCountryTreeview.focus()  
            dic = self.ReportOutsourcePerCountryTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1

    def on_Tree_ReportOutsourcePerCountry_double_click(self,event):
            curItem = self.ReportOutsourcePerCountryTreeview.focus()  
            dic = self.ReportOutsourcePerCountryTreeview.item(curItem)
            values = dic.get('values')
            if (len(values) > 0):
                curItem = int(dic.get('text')) - 1


    def Call_ButtonReportOutsourcePerCountry(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Report Outsource Cost Per Country','Run Report Button']    
            Logging.Log(Parameter)
        self.data_ready = False
        if self.db.Connect():
            sql = """SELECT * FROM Device_Model 
                            ORDER BY Device_Model_ID
                  """
            #print (sql)
            if (self.db.Execute(sql)):
                self.data_ready = True
                #------------------------------- Deleting Tree View <BEGIN> --------
                x = self.ReportOutsourcePerCountryTreeview.get_children()
                if x != '()': # checks if there is something in the first row
                    for child in x:
                        #print (child)
                        self.ReportOutsourcePerCountryTreeview.delete(child)
                #------------------------------- Deleting Tree View <END> --------
                #-------------- Initializing Arrays <BEGIN> ----------------------
                self.DeviceModelTableDeviceModelIDArray =  []
                self.DeviceModelTableDeviceTypeIDArray =  []
                #-------------- Initializing Arrays <END> ----------------------                        
                i = 0
                self.progress['maximum'] = len(self.db.results)
                self.progress['value'] = 0
                while (i < len(self.db.results)):
                    self.DeviceModelTableDeviceModelIDArray.append(self.db.results[i][0].strip())
                    self.DeviceModelTableDeviceTypeIDArray.append(self.db.results[i][5].strip())
                    i = i + 1
                    self.progress['value'] = i

                sql = """SELECT * FROM Country 
                                ORDER BY Country_Name
                      """
                #print (sql)
                if (self.db.Execute(sql)):
                    self.CountryTableCountryIDArray =  []
                    self.CountryTableCountryNameArray =  []
                    self.progress['maximum'] = len(self.db.results)
                    self.progress['value'] = 0
                    i = 0
                    while (i < len(self.db.results)):
                        self.CountryTableCountryIDArray.append(self.db.results[i][0].strip())
                        self.CountryTableCountryNameArray.append(self.db.results[i][1].strip())
                        i = i + 1
                        self.progress['value'] = i

                ################ Array of Country and Device DONE ###################
                self.report_index = 0
                self.ReportOutsourceCostPerCountryQuantity = []
                self.ReportOutsourceCostPerCountryDeviceModel = []
                self.ReportOutsourceCostPerCountryDeviceType = []
                self.ReportOutsourceCostPerCountryCountryID = []
                self.ReportOutsourceCostPerCountryRegionID = []
                self.ReportOutsourceCostPerCountryTotalCost = []
                i = 0
                self.progress['value'] = 0
                while (i < len(self.CountryTableCountryIDArray)):
                    #print (i)
                    j = 0
                    total_cost_per_Country = 0
                    total_devices_per_Country = 0
                    Country = self.CountryTableCountryNameArray[i]
                    Region = "NONE"
                    while (j < len(self.DeviceModelTableDeviceModelIDArray)):
                        #print (j)
                        sql = """SELECT * FROM Devices 
                                    WHERE Device_Model_ID = '%s' AND Country_ID = '%s'
                               """ %(self.DeviceModelTableDeviceModelIDArray[j],self.CountryTableCountryIDArray[i])
                        #print (sql)
                        if (self.db.Execute(sql)):
                            #print (self.db.results)
                            total_devices_per_Country = total_devices_per_Country + len(self.db.results)
                            self.ReportOutsourceCostPerCountryQuantity.append(len(self.db.results))
                            self.ReportOutsourceCostPerCountryDeviceModel.append(self.db.results[0][33].strip())
                            self.ReportOutsourceCostPerCountryDeviceType.append(self.db.results[0][6].strip())
                            self.ReportOutsourceCostPerCountryCountryID.append(self.CountryTableCountryNameArray[i])
                            self.ReportOutsourceCostPerCountryRegionID.append(self.db.results[0][3].strip())
                            Country = self.CountryTableCountryNameArray[i]
                            Region = self.db.results[0][3].strip()
                            total_cost = 0
                            k = 0
                            while (k < len(self.db.results)):
                                total_cost = total_cost + self.db.results[k][10]
                                k = k + 1
                            self.ReportOutsourceCostPerCountryTotalCost.append(total_cost)
                            total_cost_per_Country = total_cost_per_Country + total_cost
                            self.report_index = self.report_index + 1
                        j = j + 1
                    i = i + 1
                    self.progress['value'] = i
                    self.ReportOutsourceCostPerCountryQuantity.append(total_devices_per_Country)
                    self.ReportOutsourceCostPerCountryDeviceModel.append(" Totals ")
                    self.ReportOutsourceCostPerCountryDeviceType.append(" Totals ")
                    self.ReportOutsourceCostPerCountryCountryID.append(Country)
                    self.ReportOutsourceCostPerCountryRegionID.append(Region)
                    self.ReportOutsourceCostPerCountryTotalCost.append(total_cost_per_Country)
                    self.report_index = self.report_index + 1
                    
                i = 0
                while (i < self.report_index):
                    num = i + 1
                    if (self.ReportOutsourceCostPerCountryDeviceModel[i] == " Totals "):
                        tags = 'Total'
                    else:
                        tags = self.ReportOutsourceCostPerCountryDeviceModel[i] # To use in the futire
                    item = [
                            self.ReportOutsourceCostPerCountryQuantity[i],
                            self.ReportOutsourceCostPerCountryDeviceModel[i],
                            self.ReportOutsourceCostPerCountryDeviceType[i],
                            self.ReportOutsourceCostPerCountryCountryID[i],
                            self.ReportOutsourceCostPerCountryRegionID[i],
                            self.ReportOutsourceCostPerCountryTotalCost[i]
                            ]
                    self.ReportOutsourcePerCountryTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                    i = i + 1                    
            else:
                mbox.showerror(master=self.ReportOutsourcePerCountryFrame,title='Report Outsource Cost Per Country',
                message = 'No Records found')
                #db.Disconnect()
        else:
            mbox.showerror(master=self.ReportOutsourcePerCountryFrame,title='Report Outsource Cost Per Country',
            message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                            + 'Please make sure the ODBC DSN Name mathes: ['
                            + ODBC_DSN_name + "]")

        
    def Call_ButtonReportOutsourcePerCountry_From_Entry(self,event):
        self.Call_ButtonReportOutsourcePerCountry()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Outsource Cost Per Country"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.ReportOutsourcePerCountryTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')
                Row = 3
                Column = 1
                i = 0
                while (i < self.report_index):
                    item = [
                            self.ReportOutsourceCostPerCountryQuantity[i],
                            self.ReportOutsourceCostPerCountryDeviceModel[i],
                            self.ReportOutsourceCostPerCountryDeviceType[i],
                            self.ReportOutsourceCostPerCountryCountryID[i],
                            self.ReportOutsourceCostPerCountryRegionID[i],
                            self.ReportOutsourceCostPerCountryTotalCost[i]
                            ]
                    if (self.ReportOutsourceCostPerCountryDeviceModel[i] == " Totals "):
                        ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,13,'Bold')
                    else:
                        ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ReportOutsourcePerCountryFrame,title='Report Outsource Cost Per Country',
                            message = '!!! The File was saved !!!')
                    self.ReportOutsourcePerCountryWindow.title("Report Outsource Cost Per Country                    File: ["+self.file_name+"] SAVED")
                else:
                    mbox.showerror(master=self.ReportOutsourcePerCountryFrame,title='Report Outsource Cost Per Country',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Find_ReportOutsourcePerCountry_Window(self):
        if not self.ReportOutsourcePerCountryWindowExist:
            # Set up the Window
            self.ReportOutsourcePerCountryWindowExist = True
            self.ReportOutsourcePerCountryWindow = Tk()
            self.ReportOutsourcePerCountryWindow.geometry('1350x700+350+70')
            self.ReportOutsourcePerCountryWindow.title("Report Outsource Cost Per Country")
            self.ReportOutsourcePerCountryWindow.protocol("WM_DELETE_WINDOW", self.on_ReportOutsourcePerCountryWindow_quit)
            self.ReportOutsourcePerCountryWindow.call('tk', 'scaling', Windows_Scaling)
            self.ReportOutsourcePerCountryLabel = Label(self.ReportOutsourcePerCountryWindow,text="Helvetica", font=("Helvetica", 13))
            self.ReportOutsourcePerCountryLabel["text"] = "Report Outsource Cost Per Country"
            self.ReportOutsourcePerCountryLabel.pack()
            
            menubar = Menu(self.ReportOutsourcePerCountryWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_ReportOutsourcePerCountryWindow_quit)
            self.ReportOutsourcePerCountryWindow.config(menu=menubar)
            
                
            # Setup Frame
            self.ReportOutsourcePerCountryFrame = Frame(self.ReportOutsourcePerCountryWindow)
            self.ReportOutsourcePerCountryFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.ReportOutsourcePerCountryFrame.rowconfigure(0, weight=1)
            self.ReportOutsourcePerCountryFrame.columnconfigure(0, weight=1)


            if self.db.Connect():                
                # Setup Button
                self.buttonReportOutsourcePerCountryExecute = Button(self.ReportOutsourcePerCountryFrame, text = 'Run Report', command = self.Call_ButtonReportOutsourcePerCountry, state=ACTIVE)
                self.buttonReportOutsourcePerCountryExecute.place(x = 10, y = 7, width=75, height=25)

                self.progress = ttk.Progressbar(self.ReportOutsourcePerCountryWindow, orient="horizontal",length=500, mode="determinate")
                Total = 100
                self.progress['maximum'] = Total
                self.progress['value'] = 0
                self.progress.place(x=100, y=35)

                #------------------ TREE VIEW For ReportOutsourcePerCountry Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.ReportOutsourcePerCountryTreeviewDataColumns = ('Quantity','Device Model','Device Type','Country','Region ID','Total Cost')
                    
                self.ReportOutsourcePerCountryTreeview = ttk.Treeview(self.ReportOutsourcePerCountryFrame,columns=self.ReportOutsourcePerCountryTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerCountryTreeviewysb = Scrollbar(self.ReportOutsourcePerCountryFrame,orient=VERTICAL, command=self.ReportOutsourcePerCountryTreeview.yview)           # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerCountryTreeviewxsb = Scrollbar(self.ReportOutsourcePerCountryFrame,orient=HORIZONTAL, command=self.ReportOutsourcePerCountryTreeview.xview)         # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerCountryTreeview['yscroll'] = self.ReportOutsourcePerCountryTreeviewysb.set
                self.ReportOutsourcePerCountryTreeview['xscroll'] = self.ReportOutsourcePerCountryTreeviewxsb.set

                # setup headings and column
                self.ReportOutsourcePerCountryTreeview.heading('#0', text='Item No.',      anchor=W) # E for East and W for West and CENTER
                self.ReportOutsourcePerCountryTreeview.heading('#1', text='Quantity',    anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Quantity', False)) # E for East and W for West and CENTER
                self.ReportOutsourcePerCountryTreeview.heading('#2', text='Device Model',   anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Device Model', False)) # E for East and W for West and CENTER
                self.ReportOutsourcePerCountryTreeview.heading('#3', text='Device Type',    anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Device Type', False)) # E for East and W for West
                self.ReportOutsourcePerCountryTreeview.heading('#4', text='Country',     anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Country', False)) # E for East and W for West
                self.ReportOutsourcePerCountryTreeview.heading('#5', text='Region ID',     anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Region ID', False)) 
                self.ReportOutsourcePerCountryTreeview.heading('#6', text='Total Cost', anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerCountryTreeview, 'Total Cost', False)) 
                                   
                self.ReportOutsourcePerCountryTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.ReportOutsourcePerCountryTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.ReportOutsourcePerCountryTreeview.column('#2', stretch=1, width=10)
                self.ReportOutsourcePerCountryTreeview.column('#3', stretch=1, width=10)
                self.ReportOutsourcePerCountryTreeview.column('#4', stretch=1, width=10)
                self.ReportOutsourcePerCountryTreeview.column('#5', stretch=1, width=10)
                self.ReportOutsourcePerCountryTreeview.column('#6', stretch=1, width=10)

                # add tree and scrollbars to frame
                self.ReportOutsourcePerCountryTreeview.grid(row=1, column=0, sticky=NSEW)
                self.ReportOutsourcePerCountryTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.ReportOutsourcePerCountryTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.

                self.ReportOutsourcePerCountryTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.ReportOutsourcePerCountryTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Total', font=('Helvetica', 10), background='salmon')
                self.ReportOutsourcePerCountryTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.ReportOutsourcePerCountryTreeview.tag_configure('Closed', font=('Helvetica', 8), background='red3') 
                self.ReportOutsourcePerCountryTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                # Bind the double Click
                self.ReportOutsourcePerCountryTreeview.bind('<ButtonRelease-1>', self.on_ReportOutsourcePerCountry_Tree_select_click) # When Select the Tree
                self.ReportOutsourcePerCountryTreeview.bind("<Double-1>", self.on_Tree_ReportOutsourcePerCountry_double_click)         
                
                #------------------ TREE VIEW For ReportOutsourcePerCountry Database <END> -----------------------------------                
                self.ReportOutsourcePerCountryWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Report Outsource Cost Per Country',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   ReportOutsourcePerCountry Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Report Outsource Cost per Country Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindLPOC = Class_ReportOutsourcePerCountry("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindLPOC.Find_ReportOutsourcePerCountry_Window()


if __name__ == '__main__':
    Main()
