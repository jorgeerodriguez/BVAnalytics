#==========================================================================
# Program: Password.py ---- CLASS ---
# Author:  Jorge E. Rodriguez
# Date Created: Jan-29-2018
# Date Last Modified: Jan-29-2018
# Summary: This is Class to for the Password to remove deivces or anything
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
#

class Class_Password:
    
    def __init__(self,frame):
        #super().__init__(master)
        self.GetPasswordWindowsExists = False
        self.GetPasswordWindow = frame
        self.PassLogin = False
        

    def try_login(self):
        self.GetPasswordWindowsExists = True 
        if self.password_guess.get() == "BeCareful":
            self.PassLogin = True
            self.GetPasswordWindow.destroy()
            #Calll the Rutine
            
        else:
            mbox.showerror(master=self.GetPasswordWindow,title='Username and Password',
                             message = '*** ERROR *** - Please Enter a Valid Information')
            self.PassLogin = False
            self.GetPasswordWindow.destroy()
            
    def try_login_Enter(self,event):
        self.try_login()
        
    def on_GetPasswordWindow_quit(self):
        self.GetPasswordWindowsExists = False
        self.PassLogin = False
        self.GetPasswordWindow.destroy()
        return self.PassLogin
        

    def Get_Usernanme_and_Password(self):
        if not self.GetPasswordWindowsExists:
            self.password = ""
            self.username = ""
            self.GetPasswordWindowsExists = True
            #self.GetPasswordWindow = Tk()
            #GetPasswordWindow.title("Get Password to Access Proteus")
            self.GetPasswordWindow.resizable(width=FALSE, height=FALSE)
            self.GetPasswordWindow.protocol("WM_DELETE_WINDOW", self.on_GetPasswordWindow_quit)
            self.GetPasswordWindow.title("Log-In")
            self.GetPasswordWindow.geometry("200x150")
            #Creating the username & password entry boxes
            self.username_text = Label(self.GetPasswordWindow, text="Username:")
            self.username_guess = Entry(self.GetPasswordWindow)
            self.password_text = Label(self.GetPasswordWindow, text="Password:")
            self.password_guess = Entry(self.GetPasswordWindow, show="*")
            self.password_guess.bind('<Return>',self.try_login_Enter)
            self.attempt_login = Button(self.GetPasswordWindow,text="Login", command = self.try_login)                 
            self.username_text.pack()
            self.username_guess.pack()
            self.password_text.pack()
            self.password_guess.pack()
            self.attempt_login.pack()
            #Main Starter
            #self.GetPasswordWindow.mainloop()
            #return self.PassLogin      


def Main():
    print ("Testing the Password Class....:")
    root = Tk()
    Password = Class_Password(root)
    Password.Get_Usernanme_and_Password()
    #print (Password.PassLogin)
    root.mainloop()
    if (Password.PassLogin):
        print ("Pass")
    else:
        print ("Fail")

        
if __name__ == '__main__':
    Main()

