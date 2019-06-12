#==========================================================================
# Program: ReportOutsourcePerFacility.py ---- CLASS ---
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
    Parameter = ['Report Outsource Cost Per Facility','OPEN Window']    
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
#                                   ReportOutsourcePerFacility Section <BEGIN>                                                                                   *
#*********************************************************************************************************************************************
class Class_ReportOutsourcePerFacility:

    def __init__(self,DSN_Name,Windows_Scaling):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.db2 = ODBC(self.ODBC_name)
        self.version = self.db.Get_Version()
        self.ICMPSummaryWindowExist = False
        self.FindCircuitWindowExist = False
        self.FindDeviceWindowExist = False
        self.CircuitsWindowExist = False
        self.ReportOutsourcePerFacilityWindowExist = False
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


    def on_ReportOutsourcePerFacilityWindow_quit(self):
        if (self.ReportOutsourcePerFacilityWindowExist):
            self.ReportOutsourcePerFacilityWindowExist = False
            self.db.Disconnect()
            if Is_logging_Available:
                Parameter = []
                Parameter = ['Report Outsource Cost Per Facility','CLOSED Window']    
                Logging.Log(Parameter)
            self.ReportOutsourcePerFacilityWindow.destroy()


    def Show_Selection_on_Tree(self): # Pretty Cool
        curItem = self.ReportOutsourcePerFacilityTreeview.focus()
        dic = self.ReportOutsourcePerFacilityTreeview.item(curItem)
        #print(dic)
        position = int(dic.get('text'))
        #print ("Before => " + str(position))
    
        NextItem = self.ReportOutsourcePerFacilityTreeview.next(curItem)
        self.ReportOutsourcePerFacilityTreeview.selection_set((NextItem, NextItem))
        self.ReportOutsourcePerFacilityTreeview.focus(NextItem)

        NextItem = self.ReportOutsourcePerFacilityTreeview.focus()
        dic = self.ReportOutsourcePerFacilityTreeview.item(NextItem)
        position = int(dic.get('text'))
        #print ("After => "+ str(position))
            
    def on_ReportOutsourcePerFacility_Tree_select_click(self,event):
        curItem = self.ReportOutsourcePerFacilityTreeview.focus()  
        dic = self.ReportOutsourcePerFacilityTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1

    def on_Tree_ReportOutsourcePerFacility_double_click(self,event):
        curItem = self.ReportOutsourcePerFacilityTreeview.focus()  
        dic = self.ReportOutsourcePerFacilityTreeview.item(curItem)
        values = dic.get('values')
        if (len(values) > 0):
            curItem = int(dic.get('text')) - 1

    def Call_ButtonReportOutsourcePerFacility(self):
        if Is_logging_Available:
            Parameter = []
            Parameter = ['Report Outsource Cost Per Facility','Run Report Button']    
            Logging.Log(Parameter)
        self.data_ready = False
        if (mbox.askyesnocancel(master=self.ReportOutsourcePerFacilityFrame,title='Region',message = 'This Report MIGHT take Several Minutes, Do you want to Continue?')):
            if self.db.Connect():
                sql = """SELECT * FROM Device_Model 
                                ORDER BY Device_Model_ID
                      """
                #print (sql)
                if (self.db.Execute(sql)):
                    self.data_ready = True
                    #------------------------------- Deleting Tree View <BEGIN> --------
                    x = self.ReportOutsourcePerFacilityTreeview.get_children()
                    if x != '()': # checks if there is something in the first row
                        for child in x:
                            #print (child)
                            self.ReportOutsourcePerFacilityTreeview.delete(child)
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

                    sql = """SELECT * FROM Facility 
                                    ORDER BY Facility_Name
                          """
                    #print (sql)
                    if (self.db.Execute(sql)):
                        self.FacilityTableFacilityIDArray =  []
                        self.FacilityTableFacilityNameArray =  []
                        self.progress['maximum'] = len(self.db.results)
                        self.progress['value'] = 0
                        i = 0
                        while (i < len(self.db.results)):
                            self.FacilityTableFacilityIDArray.append(self.db.results[i][3].strip())
                            self.FacilityTableFacilityNameArray.append(self.db.results[i][4].strip())
                            i = i + 1
                            self.progress['value'] = i

                    ################ Array of Country and Device DONE ###################
                    self.report_index = 0
                    self.ReportOutsourceCostPerFacilityQuantity = []
                    self.ReportOutsourceCostPerFacilityDeviceModel = []
                    self.ReportOutsourceCostPerFacilityDeviceType = []
                    self.ReportOutsourceCostPerFacilityCountryID = []
                    self.ReportOutsourceCostPerFacilityRegionID = []
                    self.ReportOutsourceCostPerFacilityFacilityName = []
                    self.ReportOutsourceCostPerFacilityTotalCost = []
                    i = 0
                    self.progress['value'] = 0
                    while (i < len(self.FacilityTableFacilityIDArray)):
                        #print (i)
                        j = 0
                        total_cost_per_Facility = 0
                        total_devices_per_Facility = 0
                        Country = "NONE"
                        Region = "NONE"
                        Facility = self.FacilityTableFacilityNameArray[i]
                        while (j < len(self.DeviceModelTableDeviceModelIDArray)):
                            #print (j)
                            sql = """SELECT * FROM Devices 
                                        WHERE Device_Model_ID = '%s' AND Facility_ID = '%s'
                                   """ %(self.DeviceModelTableDeviceModelIDArray[j],self.FacilityTableFacilityIDArray[i])
                            #print (sql)
                            if (self.db.Execute(sql)):
                                #print (self.db.results)
                                total_devices_per_Facility = total_devices_per_Facility + len(self.db.results)
                                self.ReportOutsourceCostPerFacilityQuantity.append(len(self.db.results))
                                self.ReportOutsourceCostPerFacilityDeviceModel.append(self.db.results[0][33].strip())
                                self.ReportOutsourceCostPerFacilityDeviceType.append(self.db.results[0][6].strip())
                                self.ReportOutsourceCostPerFacilityCountryID.append(self.db.results[0][2].strip())
                                self.ReportOutsourceCostPerFacilityRegionID.append(self.db.results[0][3].strip())
                                self.ReportOutsourceCostPerFacilityFacilityName.append(self.FacilityTableFacilityNameArray[i])
                                Country = self.db.results[0][2].strip()
                                Region = self.db.results[0][3].strip()
                                Facility = self.FacilityTableFacilityNameArray[i]
                                total_cost = 0
                                k = 0
                                while (k < len(self.db.results)):
                                    total_cost = total_cost + self.db.results[k][10]
                                    k = k + 1
                                self.ReportOutsourceCostPerFacilityTotalCost.append(total_cost)
                                total_cost_per_Facility = total_cost_per_Facility + total_cost
                                self.report_index = self.report_index + 1
                            j = j + 1
                        i = i + 1
                        self.progress['value'] = i
                        self.ReportOutsourceCostPerFacilityQuantity.append(total_devices_per_Facility)
                        self.ReportOutsourceCostPerFacilityDeviceModel.append(" Totals ")
                        self.ReportOutsourceCostPerFacilityDeviceType.append(" Totals ")
                        self.ReportOutsourceCostPerFacilityCountryID.append(Country)
                        self.ReportOutsourceCostPerFacilityRegionID.append(Region)
                        self.ReportOutsourceCostPerFacilityFacilityName.append(Facility)
                        self.ReportOutsourceCostPerFacilityTotalCost.append(total_cost_per_Facility)
                        self.report_index = self.report_index + 1
                        
                    i = 0
                    while (i < self.report_index):
                        num = i + 1
                        if (self.ReportOutsourceCostPerFacilityDeviceModel[i] == " Totals "):
                            tags = 'Total'
                        else:
                            tags = self.ReportOutsourceCostPerFacilityDeviceModel[i] # To use in the futire
                        item = [
                                self.ReportOutsourceCostPerFacilityQuantity[i],
                                self.ReportOutsourceCostPerFacilityDeviceModel[i],
                                self.ReportOutsourceCostPerFacilityDeviceType[i],
                                self.ReportOutsourceCostPerFacilityCountryID[i],
                                self.ReportOutsourceCostPerFacilityRegionID[i],
                                self.ReportOutsourceCostPerFacilityFacilityName[i],
                                self.ReportOutsourceCostPerFacilityTotalCost[i]
                                ]
                        self.ReportOutsourcePerFacilityTreeview.insert('', END, text='%3d'%num, values=item, tags=tags)
                        i = i + 1
                else:
                    mbox.showerror(master=self.ReportOutsourcePerFacilityFrame,title='Report Outsource Cost Per Facility',
                    message = 'No Records found')
                    #db.Disconnect()
            else:
                mbox.showerror(master=self.ReportOutsourcePerFacilityFrame,title='Report Outsource Cost Per Facility',
                message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                + 'Please make sure the ODBC DSN Name mathes: ['
                                + ODBC_DSN_name + "]")

        
    def Call_ButtonReportOutsourcePerFacility_From_Entry(self,event):
        self.Call_ButtonReportOutsourcePerFacility()

    def Call_Save_As(self):
        #self.data_ready = True
        if (self.data_ready):
            #print ("Save as")
            input_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 2010 Documents", "*.xlsx"),("All Files", "*.*")])
            if input_file_name != "":
                self.file_name = input_file_name
                ExcellFile = Class_SaveAs(ODBC_DSN_name,Windows_Scaling,self.file_name,self.version)
                Tab0 = "Outsource Cost Per Facility"
                Tabs = [Tab0]
                ExcellFile.Call_Write_to_File(Tabs)
                Row = 2    # 3
                Column = 1 # A
                ExcellFile.Add_DataToWorksheet(self.ReportOutsourcePerFacilityTreeviewDataColumns,Row,Column,Tab0,"Fill",lightsalmon_1,14,'Bold')
                Row = 3
                Column = 1
                i = 0
                while (i < self.report_index):
                    item = [
                            self.ReportOutsourceCostPerFacilityQuantity[i],
                            self.ReportOutsourceCostPerFacilityDeviceModel[i],
                            self.ReportOutsourceCostPerFacilityDeviceType[i],
                            self.ReportOutsourceCostPerFacilityCountryID[i],
                            self.ReportOutsourceCostPerFacilityRegionID[i],
                            self.ReportOutsourceCostPerFacilityFacilityName[i],
                            self.ReportOutsourceCostPerFacilityTotalCost[i]
                           ]
                    if (self.ReportOutsourceCostPerFacilityDeviceModel[i] == " Totals "):
                        ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,13,'Bold')
                    else:
                        ExcellFile.Add_DataToWorksheet(item,Row,Column,Tab0,"No-Fill",lightsalmon_1,12,'No-Bold')
                    i = i + 1
                    Row = Row + 1
                    Column = 1
                if (ExcellFile.Save_File()):
                    mbox.showinfo(master=self.ReportOutsourcePerFacilityFrame,title='Report Outsource Cost Per Facility',
                            message = '!!! The File was saved !!!')
                    self.ReportOutsourcePerFacilityWindow.title("Report Outsource Cost Per Facility                    File: ["+self.file_name+"] SAVED")
                else:
                    mbox.showerror(master=self.ReportOutsourcePerFacilityFrame,title='Report Outsource Cost Per Facility',
                            message = '*** The File was not saved, Perhaps It is already open ***')


    def Find_ReportOutsourcePerFacility_Window(self):
        if not self.ReportOutsourcePerFacilityWindowExist:
            # Set up the Window
            self.ReportOutsourcePerFacilityWindowExist = True
            self.ReportOutsourcePerFacilityWindow = Tk()
            self.ReportOutsourcePerFacilityWindow.geometry('1350x700+350+70')
            self.ReportOutsourcePerFacilityWindow.title("Report Outsource Cost Per Facility")
            self.ReportOutsourcePerFacilityWindow.protocol("WM_DELETE_WINDOW", self.on_ReportOutsourcePerFacilityWindow_quit)
            self.ReportOutsourcePerFacilityWindow.call('tk', 'scaling', Windows_Scaling)
            self.ReportOutsourcePerFacilityLabel = Label(self.ReportOutsourcePerFacilityWindow,text="Helvetica", font=("Helvetica", 13))
            self.ReportOutsourcePerFacilityLabel["text"] = "Report Outsource Cost Per Facility"
            self.ReportOutsourcePerFacilityLabel.pack()
            
            menubar = Menu(self.ReportOutsourcePerFacilityWindow)
            filemenu = Menu(menubar, tearoff=0)
            menubar.add_cascade(label="File", menu=filemenu)
            if Is_SaveAs_Available:
                filemenu.add_command(label="Save As", command=self.Call_Save_As)
            filemenu.add_separator()
            filemenu.add_command(label="Exit", command=self.on_ReportOutsourcePerFacilityWindow_quit)
            self.ReportOutsourcePerFacilityWindow.config(menu=menubar)
            
                
            # Setup Frame
            self.ReportOutsourcePerFacilityFrame = Frame(self.ReportOutsourcePerFacilityWindow)
            self.ReportOutsourcePerFacilityFrame.pack(side=TOP, fill=BOTH, expand=Y)

            # set frame resizing priorities
            self.ReportOutsourcePerFacilityFrame.rowconfigure(0, weight=1)
            self.ReportOutsourcePerFacilityFrame.columnconfigure(0, weight=1)


            if self.db.Connect():                
                # Setup Button
                self.buttonReportOutsourcePerFacilityExecute = Button(self.ReportOutsourcePerFacilityFrame, text = 'Run Report', command = self.Call_ButtonReportOutsourcePerFacility, state=ACTIVE)
                self.buttonReportOutsourcePerFacilityExecute.place(x = 10, y = 7, width=75, height=25)

                self.progress = ttk.Progressbar(self.ReportOutsourcePerFacilityWindow, orient="horizontal",length=500, mode="determinate")
                Total = 100
                self.progress['maximum'] = Total
                self.progress['value'] = 0
                self.progress.place(x=100, y=35)
                #self.progress.start()

                #------------------ TREE VIEW For ReportOutsourcePerFacility Database <BEGIN> -----------------------------------
                # Create Tree and Scrollbars
                self.ReportOutsourcePerFacilityTreeviewDataColumns = ('Quantity','Device Model','Device Type','Country ID','Region ID','Facility','Total Cost')
                    
                self.ReportOutsourcePerFacilityTreeview = ttk.Treeview(self.ReportOutsourcePerFacilityFrame,columns=self.ReportOutsourcePerFacilityTreeviewDataColumns, height=29) # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerFacilityTreeviewysb = Scrollbar(self.ReportOutsourcePerFacilityFrame,orient=VERTICAL, command=self.ReportOutsourcePerFacilityTreeview.yview)           # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerFacilityTreeviewxsb = Scrollbar(self.ReportOutsourcePerFacilityFrame,orient=HORIZONTAL, command=self.ReportOutsourcePerFacilityTreeview.xview)         # <--- Make sure the frame is correct !!
                self.ReportOutsourcePerFacilityTreeview['yscroll'] = self.ReportOutsourcePerFacilityTreeviewysb.set
                self.ReportOutsourcePerFacilityTreeview['xscroll'] = self.ReportOutsourcePerFacilityTreeviewxsb.set

                # setup headings and column
                self.ReportOutsourcePerFacilityTreeview.heading('#0', text='Item No.',     anchor=W) # E for East and W for West and CENTER
                self.ReportOutsourcePerFacilityTreeview.heading('#1', text='Quantity',     anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Quantity', False)) # E for East and W for West and CENTER
                self.ReportOutsourcePerFacilityTreeview.heading('#2', text='Device Model', anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Device Model', False)) # E for East and W for West and CENTER
                self.ReportOutsourcePerFacilityTreeview.heading('#3', text='Device Type',  anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Device Type', False)) # E for East and W for West
                self.ReportOutsourcePerFacilityTreeview.heading('#4', text='Country ID',   anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Country ID', False)) # E for East and W for West
                self.ReportOutsourcePerFacilityTreeview.heading('#5', text='Region ID',    anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Region ID', False)) 
                self.ReportOutsourcePerFacilityTreeview.heading('#6', text='Facility',     anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Facility', False)) 
                self.ReportOutsourcePerFacilityTreeview.heading('#7', text='Total Cost',   anchor=W,command=lambda: self.treeview_sort_column(self.ReportOutsourcePerFacilityTreeview, 'Total Cost', False)) 
                                   
                self.ReportOutsourcePerFacilityTreeview.column('#0', stretch=1, width=3 , anchor=W)
                self.ReportOutsourcePerFacilityTreeview.column('#1', stretch=1, width=10, anchor=W)
                self.ReportOutsourcePerFacilityTreeview.column('#2', stretch=1, width=10)
                self.ReportOutsourcePerFacilityTreeview.column('#3', stretch=1, width=10)
                self.ReportOutsourcePerFacilityTreeview.column('#4', stretch=1, width=10)
                self.ReportOutsourcePerFacilityTreeview.column('#5', stretch=1, width=10)
                self.ReportOutsourcePerFacilityTreeview.column('#6', stretch=1, width=10)
                self.ReportOutsourcePerFacilityTreeview.column('#7', stretch=1, width=10)

                # add tree and scrollbars to frame
                self.ReportOutsourcePerFacilityTreeview.grid(row=1, column=0, sticky=NSEW)
                self.ReportOutsourcePerFacilityTreeviewysb.grid(row=1, column=1, sticky=NS)
                self.ReportOutsourcePerFacilityTreeviewxsb.grid(row=2, column=0, sticky=EW)

                # create fonts and tags
                # Use later to mark Business Units per color.

                self.ReportOutsourcePerFacilityTreeview.tag_configure('YES', font=('Helvetica', 8), background='green2')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('NO', font=('Helvetica', 8), background='orange')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Total', font=('Helvetica', 10), background='salmon')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('New_Site', font=('Helvetica', 8), background='royal blue1')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Quote_Requested', font=('Helvetica', 8), background='cyan2')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Quote_Received', font=('Helvetica', 8), background='salmon')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Quote_Approved', font=('Helvetica', 8), background='gold')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Active', font=('Helvetica', 8), background='lawn green')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Renovation', font=('Helvetica', 8), background='plum1')
                self.ReportOutsourcePerFacilityTreeview.tag_configure('Closed', font=('Helvetica', 8), background='red3') 
                self.ReportOutsourcePerFacilityTreeview.tag_configure('None', font=('Helvetica', 8), background='gray1')
                   
                # Bind the double Click
                self.ReportOutsourcePerFacilityTreeview.bind('<ButtonRelease-1>', self.on_ReportOutsourcePerFacility_Tree_select_click) # When Select the Tree
                self.ReportOutsourcePerFacilityTreeview.bind("<Double-1>", self.on_Tree_ReportOutsourcePerFacility_double_click)

                #self.On_Run_Progres_Bar(5)
                #self.ReportOutsourcePerFacilityProgressBarWindow.destroy()
                
                #------------------ TREE VIEW For ReportOutsourcePerFacility Database <END> -----------------------------------                
                self.ReportOutsourcePerFacilityWindow.mainloop()
            else:
                mbox.showerror(master=root,title='Report Outsource Cost Per Facility',
                                 message = '*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                           + 'Please make sure the ODBC DSN Name mathes: ['
                                           + ODBC_DSN_name + "]")


#*********************************************************************************************************************************************
#                                   ReportOutsourcePerFacility Section <END>                                                                                   *
#*********************************************************************************************************************************************


        
def Main():
    print ("Testing the Report Outsource Cost Per Facility Class....:")
    location = []
    #location = ['UNKNOWN','UNKNOWN','UNKNOWN','UNKNOWN']
    FindLPOC = Class_ReportOutsourcePerFacility("BV",Windows_Scaling)
    #FindCEF.Find_Circuits_Window()
    #FindCEF.Find_Devices_Window()
    FindLPOC.Find_ReportOutsourcePerFacility_Window()


if __name__ == '__main__':
    Main()
