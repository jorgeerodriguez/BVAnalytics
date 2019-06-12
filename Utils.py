#==========================================================================
# Program: Utils.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Dec-22-2017
# Date Last Modified: Dec-23-2017
# Summary: This is Class to for the Utils
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

#*************************************************************
# ==================== Libraries Required <END> =============*
#*************************************************************

class Class_Utils:

    def __init__(self):      
        self.ODBC_name = "BV"
        self.db = ODBC(self.ODBC_name)
        self.UtilsWindowExist = False
        self.Username = os.getlogin()
        self.BackendServer = "NONE"


    def Get_Values(self):
        #time.sleep(2)
        self.Database_Version = "0.0"
        self.ODBC_DNS_Name = "BV"
        self.Created_Date = "Dec-26-1966"
        self.Last_Time_Updated = "Dec-26-1966"
        self.Updated_By_User = "manual"
        self.BackendServer = "NONE"
        self.Enviroment = "NONE"
        if self.db.Connect():            
            # SQL Querry to the Utils Table
            sql = """ SELECT * FROM Variables """
            if (self.db.Execute(sql)):
                i = 0
                try:
                    while (i < len(self.db.results)):
                        num = i + 1
                        self.Database_Version = self.db.results[i][0].strip()
                        self.ODBC_DNS_Name = self.db.results[i][1].strip()
                        self.Created_Date = self.db.results[i][2].strip()
                        self.Last_Time_Updated = self.db.results[i][3].strip()
                        self.Updated_By_User = self.db.results[i][4].strip()
                        if (self.Database_Version != "1.0"):
                            self.Windows_Scaling = self.db.results[i][5] # <- Did not exit on DB Ver 1.0
                        else:
                            self.Windows_Scaling = 1.0
                        i = i + 1
                    self.BackendServer = self.db.Get_BackendServer()
                    self.Enviroment = self.db.Get_Enviroment()
                except:
                    self.db.Disconnect()
            self.db.Disconnect()

    def Get_DB_Ver(self):
        return self.Database_Version

    def Get_ODBC_Name(self):
        return self.ODBC_DNS_Name

    def Get_Windows_Scaling(self):
        return self.Windows_Scaling

    def Get_BackendServer(self):
        return self.BackendServer

    def Get_Enviroment(self):
        return self.Enviroment


def Main():
    print ("Testing the Utils Class....:")
    Utils = Class_Utils()
    Utils.Get_Values()
    print (Utils.Get_DB_Ver()+"....")
    print (Utils.Get_ODBC_Name())
    print (Utils.Get_Windows_Scaling())
    print (Utils.Get_BackendServer())
    print (Utils.Get_Enviroment())

if __name__ == '__main__':
    Main()

