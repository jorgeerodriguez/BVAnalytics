####################################################################
#
# IPAddressCalculator.py
#
# Created by Jorge Rodriguez, May-26-2018
# Date Last Modified: May-26-2018
####################################################################


#---------------------- Gloval Variables ---------------------

#---------------------- < IPAddressCalculator CLASS BEGIN >-----------------------
class IPAddressCalculator:

    def __init__(self,IP_Address,Slash):
        self.IP_Address = IP_Address
        self.Slash = Slash
        self.Subnetmask = []
        self.Subnetmask.append("0.0.0.0")   #0
        self.Subnetmask.append("128.0.0.0") #1
        self.Subnetmask.append("192.0.0.0") #2
        self.Subnetmask.append("224.0.0.0") #3
        self.Subnetmask.append("240.0.0.0") #4
        self.Subnetmask.append("248.0.0.0") #5
        self.Subnetmask.append("252.0.0.0") #6
        self.Subnetmask.append("254.0.0.0") #7
        self.Subnetmask.append("255.0.0.0") #8
        
        self.Subnetmask.append("255.128.0.0") #9
        self.Subnetmask.append("255.192.0.0") #10
        self.Subnetmask.append("255.224.0.0") #11
        self.Subnetmask.append("255.240.0.0") #12
        self.Subnetmask.append("255.248.0.0") #13
        self.Subnetmask.append("255.252.0.0") #14
        self.Subnetmask.append("255.254.0.0") #15
        self.Subnetmask.append("255.255.0.0") #16
        
        self.Subnetmask.append("255.255.128.0") #17
        self.Subnetmask.append("255.255.192.0") #18
        self.Subnetmask.append("255.255.224.0") #19
        self.Subnetmask.append("255.255.240.0") #20
        self.Subnetmask.append("255.255.248.0") #21
        self.Subnetmask.append("255.255.252.0") #22
        self.Subnetmask.append("255.255.254.0") #23
        self.Subnetmask.append("255.255.255.0") #24
        
        self.Subnetmask.append("255.255.255.128") #25
        self.Subnetmask.append("255.255.255.192") #26
        self.Subnetmask.append("255.255.255.224") #27
        self.Subnetmask.append("255.255.255.240") #28
        self.Subnetmask.append("255.255.255.248") #29
        self.Subnetmask.append("255.255.255.252") #30
        self.Subnetmask.append("255.255.255.254") #31
        self.Subnetmask.append("255.255.255.255") #32

    def IPFormatCheck(self,ip_str):
        if len(ip_str.split()) == 1:
               ipList = ip_str.split('.')
               if len(ipList) == 4:
                   for i, item in enumerate(ipList):
                       try:
                           ipList[i] = int(item)
                       except:
                           return False
                       if not isinstance(ipList[i], int):
                           return False
                   if max(ipList) < 256:
                       return True
                   else:
                       return False
               else:
                   return False
        else:
               return False


    def Is_Valid(self):
        if (self.IPFormatCheck(self.IP_Address)):
            return True
        else:
            return False
    
    def Get_Hosts(self):
        results = [-1]
        if (self.IPFormatCheck(self.IP_Address)):
            self.IP_Address_Octects = self.IP_Address.split(".")
            self.Hosts = 2 ** (32 - self.Slash)
            results = [self.Hosts] 
            return results
        else:
            return results
            
            
    def Get_Networks(self):
        results = [-1]
        if (self.IPFormatCheck(self.IP_Address)):
            self.IP_Address_Octects = self.IP_Address.split(".")
            if (self.Slash < 8):
                self.Networks = (2 ** (self.Slash)) 
            else:
                if (self.Slash >= 8 and self.Slash < 16):
                    self.Networks = (2 ** (self.Slash - 8)) 
                else:
                    if (self.Slash >= 16 and self.Slash < 24):
                        self.Networks = (2 ** (self.Slash - 16)) 
                    else:
                        if (self.Slash >= 24 and self.Slash <= 32):
                            self.Networks = (2 ** (self.Slash - 24))

            results = [self.Networks]
            return results
        else:
            return results
            
    def Get_Network_Range(self):
        results = [-1]
        if (self.IPFormatCheck(self.IP_Address)):
            Networks = self.Get_Networks()
            Hosts = self.Get_Hosts()

            self.IP_Address_Octects = self.IP_Address.split(".")
            self.Subnetmask_Octects = self.Subnetmask[self.Slash].split(".")
            self.From_IPA = int(self.IP_Address_Octects[0]) & int(self.Subnetmask_Octects[0])
            self.From_IPB = int(self.IP_Address_Octects[1]) & int(self.Subnetmask_Octects[1])
            self.From_IPC = int(self.IP_Address_Octects[2]) & int(self.Subnetmask_Octects[2])
            self.From_IPD = int(self.IP_Address_Octects[3]) & int(self.Subnetmask_Octects[3])
            
            self.To_IPA = 255 - int(self.Subnetmask_Octects[0]) + self.From_IPA
            self.To_IPB = 255 - int(self.Subnetmask_Octects[1]) + self.From_IPB
            self.To_IPC = 255 - int(self.Subnetmask_Octects[2]) + self.From_IPC
            self.To_IPD = 255 - int(self.Subnetmask_Octects[3]) + self.From_IPD

            results = [self.From_IPA,self.From_IPB,self.From_IPC,self.From_IPD,self.To_IPA,self.To_IPB,self.To_IPC,self.To_IPD]
            return results
        else:
            return results


    def Get_Network_Range_All(self,Total_Number_of_Subnets):
        results = ['-1','-1','-1','-1']
        results_range = []
        if (self.IPFormatCheck(self.IP_Address)):
            Networks = self.Get_Networks()
            Hosts = self.Get_Hosts()
            Original_Network = self.Get_Network_Range()
            Original_From_IPA = self.From_IPA
            Original_From_IPB = self.From_IPB
            Original_From_IPC = self.From_IPC
            Original_From_IPD = self.From_IPD
            Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
            self.IP_Address = Original_IP_Address
            Number_Of_Subnets = 0
            #print (Original_IP_Address)
            if (self.Slash < 8):
                l = 0
                while ((l < Networks[0]) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                    results = self.Get_Network_Range()
                    if (results[0] != -1):
                        results_range.append(self.Get_Network_Range())
                        #print (self.Get_Network_Range())
                        Number_Of_Subnets = Number_Of_Subnets + 1
                        IPA = self.To_IPA + 1
                        IPB = self.To_IPB
                        IPC = self.To_IPC
                        IPD = self.To_IPD
                        self.IP_Address = str(IPA)+"."+str(IPB)+"."+str(IPC)+"."+str(IPD)
                    l = l + 1
                    #print ("----- 1st Octects -----")
            else:
                if (self.Slash >= 8 and self.Slash < 16):
                    i = 0
                    while (((Original_From_IPA) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                        l = 0
                        while ((l < Networks[0]) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                            results = self.Get_Network_Range()
                            if (results[0] != -1):
                                results_range.append(self.Get_Network_Range())
                                #print (self.Get_Network_Range())
                                Number_Of_Subnets = Number_Of_Subnets + 1
                                IPA = self.To_IPA
                                IPB = self.To_IPB + 1
                                IPC = self.To_IPC
                                IPD = self.To_IPD
                                self.IP_Address = str(IPA)+"."+str(IPB)+"."+str(IPC)+"."+str(IPD)
                            l = l + 1
                        i = i + 1
                        Original_From_IPD = 0
                        Original_From_IPC = 0
                        Original_From_IPB = 0
                        Original_From_IPA = Original_From_IPA + 1
                        if (Original_From_IPA == 255):
                            Original_From_IPB = 255
                            Original_From_IPC = 255
                        self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                        #print ("----- 1st Octects -----")
                else:
                    if (self.Slash >= 16 and self.Slash < 24):
                        i = 0
                        while (((Original_From_IPA) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                            j = 0
                            while (((Original_From_IPB) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                l = 0
                                while ((l < Networks[0]) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                    results = self.Get_Network_Range()
                                    if (results[0] != -1):
                                        results_range.append(self.Get_Network_Range())
                                        #print (self.Get_Network_Range())
                                        Number_Of_Subnets = Number_Of_Subnets + 1
                                        IPA = self.To_IPA
                                        IPB = self.To_IPB
                                        IPC = self.To_IPC + 1
                                        IPD = self.To_IPD
                                        self.IP_Address = str(IPA)+"."+str(IPB)+"."+str(IPC)+"."+str(IPD)
                                    l = l + 1
                                j = j + 1
                                Original_From_IPD = 0
                                Original_From_IPC = 0
                                Original_From_IPB = Original_From_IPB + 1
                                if (Original_From_IPB == 255):
                                    Original_From_IPC = 255
                                self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                                #print ("----- 2nd Octects -----")
                            i = i + 1
                            Original_From_IPD = 0
                            Original_From_IPC = 0
                            Original_From_IPB = 0
                            Original_From_IPA = Original_From_IPA + 1
                            if (Original_From_IPA == 255):
                                Original_From_IPB = 255
                                Original_From_IPC = 255
                            self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                            #print ("----- 1st Octects -----")
                    else:
                        if (self.Slash >= 24 and self.Slash <= 32): # ----------------------> 4th Octtect breakout <-----------------------
                            i = 0
                            while (((Original_From_IPA) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                j = 0
                                while (((Original_From_IPB) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                    k = 0
                                    while (((Original_From_IPC) <= 255) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                        l = 0
                                        while ((l < Networks[0]) and (Number_Of_Subnets < Total_Number_of_Subnets)):
                                            results = self.Get_Network_Range()
                                            if (results[0] != -1):
                                                results_range.append(self.Get_Network_Range())
                                                #print (self.Get_Network_Range())
                                                Number_Of_Subnets = Number_Of_Subnets + 1
                                                IPA = self.To_IPA
                                                IPB = self.To_IPB
                                                IPC = self.To_IPC
                                                IPD = self.To_IPD + 1
                                                self.IP_Address = str(IPA)+"."+str(IPB)+"."+str(IPC)+"."+str(IPD)
                                            l = l + 1
                                        #print ("----- 3rd Octect -----")
                                        Original_From_IPD = 0
                                        Original_From_IPC = Original_From_IPC + 1
                                        k = k + 1
                                        self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                                    j = j + 1
                                    Original_From_IPD = 0
                                    Original_From_IPC = 0
                                    Original_From_IPB = Original_From_IPB + 1
                                    if (Original_From_IPB == 255):
                                        Original_From_IPC = 255
                                    self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                                    #print ("----- 2nd Octects -----")
                                i = i + 1
                                Original_From_IPD = 0
                                Original_From_IPC = 0
                                Original_From_IPB = 0
                                Original_From_IPA = Original_From_IPA + 1
                                if (Original_From_IPA == 255):
                                    Original_From_IPB = 255
                                    Original_From_IPC = 255
                                self.IP_Address = Original_IP_Address = str(Original_From_IPA)+"."+str(Original_From_IPB)+"."+str(Original_From_IPC)+"."+str(Original_From_IPD)
                                #print ("----- 1st Octects -----")

            return results_range
        else:
            results_range.append(results)
            return results
             
    def Get_Version(self):
        return "6.0"

        


#---------------------- < IPAddressCalculatorCLASS ENDS >-----------------------
            
#except connection.Error as e:
#    print("Error %d: %s" % (e.args[0], e.args[1]))
    #sys.exit(1)
    # Rollback in case there is any error
#    print ("duplicate")
#    connection.rollback()
    
def Main():
    print ("Testing the IPAddressCalculator Class....:")
    i = 1
    while i <= 32:
        IPCalc = IPAddressCalculator("10.0.0.0",i)
        print ("Slash: " + str(i))
        print ("No. of Hosts:")
        print (IPCalc.Get_Hosts())
        print ("No. of Possible Networks:")
        print (IPCalc.Get_Networks())
        print ("Network Range:")
        print (IPCalc.Get_Network_Range())
        print ("-------------------------")
        i = i + 1
    Network = "10.255.254.129"
    Network = "10.0.0.0"
    
    Network = "10.208.32.0"
    Slash = 21
    Brake_into_Slash = 24
    Number_of_Consecutive_Networks = 4
    Number_of_Consecutive_Networks = 2**(Brake_into_Slash - Slash)
    
    IPCalc = IPAddressCalculator(Network,Brake_into_Slash)
    results = IPCalc.Get_Network_Range_All(Number_of_Consecutive_Networks)
    #print (results)
    i = 0
    while (i < len(results)):
        if (results[i][:len(results[i])] == '-1'):
            print ("Error")
            i = i + len(results)
        else:
            #print (len(results[i]))
            print (results[i])
            #print (results[i][7])
            i = i + 1
        
    

if __name__ == '__main__':
    Main()

