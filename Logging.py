#==========================================================================
# Program: Logging.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-30-2017
# Date Last Modified: Feb-18-2017
# Summary: This is Class to for the Logging
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

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_Logging:

    def __init__(self,DSN_Name,Parameters):      
        self.ODBC_name = DSN_Name
        self.db = ODBC(self.ODBC_name)
        self.Username = os.getlogin()
        self.Windows_Scaling = Windows_Scaling
        self.now = datetime.datetime.now()
        self.day = self.now.day
        self.month = self.now.month
        self.year = self.now.year
        self.hour = self.now.hour
        self.minute = self.now.minute
        self.second = self.now.second
        self.date = str(self.month) + "/" + str(self.day) + "/" + str(self.year)
        self.time = str(self.hour) + ":" + str(self.minute) + ":" + str(self.second)
        self.Go_To_Parameters = False
        if (len(Parameters) > 0):
            self.Module = Parameters[0]
            self.Function = Parameters[1]
            self.Go_To_Parameters = True
            
    def Log(self,Parameters):
        if self.db.Connect():
            self.Go_To_Parameters = False
            if (len(Parameters) > 0):
                self.Module = Parameters[0]
                self.Function = Parameters[1]
                self.Go_To_Parameters = True
            if (self.Go_To_Parameters):
                self.now = datetime.datetime.now()
                self.day = self.now.day
                self.month = self.now.month
                self.year = self.now.year
                self.hour = self.now.hour
                self.minute = self.now.minute
                self.second = self.now.second
                self.date = str(self.month) + "/" + str(self.day) + "/" + str(self.year)
                self.time = str(self.hour) + ":" + str(self.minute) + ":" + str(self.second)
                sql = "INSERT INTO Logging(User_ID, Log_Date, Log_Day, Log_Month, Log_Year, \
                                   Log_Time, Log_Hour, Log_Min, Log_Sec, Log_Module, Log_Function, Executed_by_UserID) \
                                   VALUES ('%s','%s','%d','%d','%d','%s','%d','%d','%d','%s','%s','%s')" % (self.Username, self.date, self.day, self.month,
                                        self.year, self.time, self.hour, self.minute, self.second, self.Module, self.Function, self.Username)
                #print (sql)
                        
                '''
                sql = """CREATE TABLE LOGGING (
                    User_ID              CHAR(20) NOT NULL,
                    Log_Date             CHAR(20) NOT NULL,
                    Log_Day              INT,
                    Log_Month            INT,
                    Log_Year             INT,
                    Log_Time             CHAR(20) NOT NULL,
                    Log_Hour             INT,
                    Log_Min              INT,
                    Log_Sec              INT,
                    Log_Module           CHAR(50) NOT NULL,
                    Log_Function         CHAR(50) NOT NULL,
                    Executed_by_UserID   CHAR(20))"""

                '''
                if (self.db.Add_Move_Change_Data(sql)):
                    x = 0
                    #print ('*** The Logging ID you entered was Added ***')
                else:
                    x = 1
                    #print ('*** The Logging ID you entered was NOT Added ***')
        else:
            message = ('*** ERROR *** - THE ODBC Connection was NOT Succesful, \r\n'
                                       + 'Please make sure the ODBC DSN Name mathes: ['
                                       + self.ODBC_name + "]")
            #print (message)


def Main():
    print ("Testing the Logging Class....:")
    #root = widget.winfo_toplevel()
    #varpower = IntVar()
    Parameter = []
    Parameter = ['Module','Function','UNKNOWN','NO']
    Parameter = ['Module','Function'] 
    Logging = Class_Logging("BV",Parameter)
    Logging.Log(Parameter)
    Parameter = ['Module1','Function'] 
    Logging.Log(Parameter)
    Parameter = ['Module2','Function'] 
    Logging.Log(Parameter)
    Parameter = ['Module3','Function'] 
    Logging.Log(Parameter)
    Parameter = ['Module4','Function'] 
    Logging.Log(Parameter)

if __name__ == '__main__':
    Main()

