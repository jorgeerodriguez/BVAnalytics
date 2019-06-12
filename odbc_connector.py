####################################################################
#
# ODBC Connectio to Access and MS SQL
# Create Tables, Select Fields, etc
# Created by Jorge Rodriguez, Nov-3-2017
# Date Last Modified: Nov-18-2017
####################################################################

try:
    import pyodbc
except:
    print ("\n----------------------------------------------------\n")
    print ("  | You don't have the pyodbc library installed. |")
    print ("\n----------------------------------------------------\n")
    sys.exit(1)    

#---------------------- Gloval Variables ---------------------

#---------------------- < ODBC CLASS BEGIN >-----------------------
class ODBC:

        def __init__(self,DSN_Name):
            self.ODBC_name = DSN_Name
            self.TestEnviroment  = False # True or False
            self.BackendServer = 'ACCESS' # Values = ['SQL' or 'ACCESS']

        def Get_BackendServer(self):
            return self.BackendServer

        def Get_Enviroment(self):
            if (self.TestEnviroment):
                return "*** TEST Enviroment ***"
            else:
                return "Production Enviroment"
            
        def Get_Version(self):
            return "4.0"

        
        def Connect(self):
            try:
                if (self.BackendServer == 'SQL'):
                    # SQL Azure Connector:
                    if (self.TestEnviroment):
                        database = 'TestBVAnalytics'
                    else:
                        database = 'BVAnalytics'                    
                    server = 'bvinfrastructure.database.windows.net'
                    username = 'infrastructure@bvinfrastructure'
                    password = 'Password'
                    driver= '{ODBC Driver 13 for SQL Server}'
                    connection = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
                else:
                    # Access Connector:
                    connection = pyodbc.connect("DSN="+self.ODBC_name)
                self.db = connection.cursor()
                self.db2 = connection.cursor()
                return True
            except:
                return False

        def Disconnect(self):
            try:
                self.db.close()
                self.db2.close()
                return True
            except:
                return False
            
        def Get_Table_Names(self):
            Table_Names = []
            Table_Index = 0
            for rows in self.db.tables():
               if rows.table_type == "TABLE":
                   #print ("Table Name:"+rows.table_name)
                   Table_Names.insert(Table_Index,rows.table_name)
                   Table_Index += 1
            return (Table_Names)

        def Get_All_Tables_Fields(self):
            tblCount = 0
            Table_Names_and_Fields = []
            Table_Index = 0
            for rows in self.db.tables():
               if rows.table_type == "TABLE":
                   tblCount += 1
                   #print ("Table Name:"+rows.table_name)
                   for fld in self.db2.columns(rows.table_name):
                       #print("Table Name:"+fld.table_name + "["+fld.column_name+"]")
                       Table_Names_and_Fields.insert(Table_Index,(fld.table_name + ":"+fld.column_name))
                       Table_Index += 1
            return (Table_Names_and_Fields)


        def Get_Table_Fields(self,Table_Name):
            tblCount = 0
            Table_Fields = []
            Table_Index = 0
            for rows in self.db.tables():
               if rows.table_type == "TABLE":
                   tblCount += 1
                   if (rows.table_name == Table_Name):
                        #print ("Table Name:"+rows.table_name)
                        for fld in self.db2.columns(rows.table_name):
                            #print("Table Name:"+fld.table_name + "["+fld.column_name+"]")
                            Table_Fields.insert(Table_Index,fld.column_name)
                            Table_Index += 1
            return (Table_Fields)
        
        def Execute(self,SQL):
            try:
                self.db.execute(SQL)
                results = []
                results_index = 0
                for row in self.db.fetchall():
                    results.insert(results_index,row)
                    results_index += 1
                    #print (row)
                    #print (row[1])
                #print ("the Result Index => " + str(results_index))
                if (results_index > 0):
                    self.results = results
                    return True
                else:
                    return False
            except:
                return False

        def Create_Index(self,SQL):
            try:
                self.db.execute(SQL)
                self.db.commit()
                return True
            except:
                return False


        def Drop_Index(self,SQL):
            try:
                self.db.execute(SQL)
                self.db.commit()
                return True
            except:
                return False


        def Drop_Table(self,Table_Name):
            try:
                self.db.execute("DROP TABLE "+Table_Name)
                self.db.commit()
                return True
            except:
                return False
                
        def Create_Table(self,SQL):
            try:
                self.db.execute(SQL)
                self.db.commit()
                return True
            except:
                return False    # Table alredy exist

        def Add_Move_Change_Data(self,SQL):
            try:
               self.db.execute(SQL)
               self.db.commit()
               return True
            except:
               # Rollback in case there is any error or Duplicate it
               self.db.rollback()
               return False

        def Commit(self):
            self.db.commit()

        def Alter_Table_Add_Field(self,Table_Name,Field_Name,Type):
            try:
                #SQL = ("ALTER TABLE " + Table_Name + 
                #      " ADD COLUMN " + Field_Name +
                #      " " + Type)
                # on Feb 27 found that the WORD COLUMN is not required in MSSQL
                SQL = ("ALTER TABLE " + Table_Name + 
                      " ADD " + Field_Name + 
                      " " + Type)
                self.db.execute(SQL)
                self.db.commit()
                return True
            except:
                return False

        def Alter_Table_Drop_Field(self,Table_Name,Field_Name):
            try:
                
                SQL = ("ALTER TABLE " + Table_Name +
                      " DROP COLUMN " + Field_Name)
                #print (SQL)
                self.db.execute(SQL)
                self.db.commit()
                return True
            except:
                return False


#---------------------- < ODBC CLASS ENDS >-----------------------
            
#except connection.Error as e:
#    print("Error %d: %s" % (e.args[0], e.args[1]))
    #sys.exit(1)
    # Rollback in case there is any error
#    print ("duplicate")
#    connection.rollback()
    
def Main():
    print ("Testing the ODBC Class....:")
    db = ODBC("BV")
    if db.Connect():
        print ("Success")

        #------------- CREATE TABLES <BEGIN>-----------------------
        '''
        sql = """CREATE TABLE ICMP (
                Device_IP_Date_Time_Size_of_Ping CHAR(100) NOT NULL PRIMARY KEY,
                Device_IP                        CHAR(20) NOT NULL,
                Date_String                      CHAR(20) NOT NULL,
                Time_String                      CHAR(20) NOT NULL,
                Day                              INT,
                Month                            INT,
                Year                             INT,
                Hour                             INT,
                Minute                           INT,
                Second                           INT,
                Size_of_Ping                     CHAR(10) NOT NULL,
                Percentage_Loss                  INT,
                Response_Time_Max                INT,
                Response_Time_Min                INT,
                Response_Time_Avg                INT,
                Response_Status                  CHAR(10),
                Executed_by_UserID               CHAR(20) )"""
        '''
        '''
        sql = """CREATE TABLE EMPLOYEE (
                FIRST_NAME  CHAR(20) NOT NULL PRIMARY KEY,
                LAST_NAME  CHAR(20),
                AGE INT,  
                SEX CHAR(1),
                INCOME FLOAT )"""
        
        if (db.Create_Table(sql)):
            print ("Table Created")
        else:
            print ("Table Already Exist")
        
        sql = """CREATE TABLE TestTable2(
                    symbol varchar(15),
                    leverage double,
                    shares integer,
                    price double)"""
        if (db.Create_Table(sql)):
            print ("Table Created")
        else:
            print ("Table Already Exist")
        #------------- CREATE TABLES <END>--------------------------
        
        #------------- DROP TABLES <BEGIN>-----------------------
        table_name = ("TestTable")
        if (db.Drop_Table(table_name)):
            print ("Table TestTable Droped")
        else:
            print ("Table Does NOT Exist")
        #------------- DROP TABLES <END>-----------------------

        #------------- ADD Fields to TABLES <BEGIN>-------------------
        if (db.Alter_Table_Add_Field("EMPLOYEE","ADDRESS","CHAR(20)")):
            print ("Field Added")
        else:
            print ("Error to Add the Field")
        #------------- ADD Fields to TABLES <END>-------------------

        #------------- DROP Fields to TABLES <BEGIN>-------------------
        if (db.Alter_Table_Drop_Field("EMPLOYEE","ADDRESS")):
            print ("Field droped")
        else:
            print ("Error to Add the Field")
        #------------- DROP Fields to TABLES <END>-------------------

        #------------- GET all the Tables & Fields <BEGIN>-----------
        all_tables = db.Get_All_Tables_Fields() 
        print (all_tables)
        print (all_tables[0].split(":"))
        print ("----------------------------------------------------")
        #------------- GET all the Tables & Fields <END>------------

        #------------- GET All Table Names <BEGIN>------------------
        print ("----------- Table Names ----------------------------")
        print (db.Get_Table_Names())
        print ("----------------------------------------------------")
        #------------- GET All Table Names <END>--------------------
        
        #------------- GET All Fields from a Table <BEGIN>----------
        print ("------------Fields in Table Country ----------------")
        print (db.Get_Table_Fields("Country"))
        print ("----------------------------------------------------")
        #------------- GET All Fields from a Table <END>----------

        #------------- SQL SELECT <BEGIN> ------------------------
        sql = "SELECT * FROM Country WHERE CountryID = '%s'" % ("USA")
        print (db.Execute(sql))
        print ("----------------------------------------------------")
        
        sql = 'SELECT * from country'
        if (db.Execute(sql)):
            print ("No of Rows:"+str(len(db.results)))
            print (db.results)
            print (db.results[0][2])

        sql = "SELECT * FROM EMPLOYEE WHERE FIRST_NAME = '%s'" % ("Mac")
        if (db.Execute(sql)):
            print ("No of Rows:"+str(len(db.results)))
            print (db.results)
            print (db.results[0][0])
        else:
            print ("Record not Found")
        #------------- SQL SELECT <END> ------------------------
            
        #------------- INSERT data into a Table <BEGIN> ---------
        sql = """INSERT INTO EMPLOYEE(FIRST_NAME,
                     LAST_NAME, AGE, SEX, INCOME)
                     VALUES ('Daniela1', 'Rodriguez-Chavez', 17, 'F', 10000)"""
        if (db.Add_Move_Change_Data(sql)):
            print ("Record Added it....!!!")
        else:
            print ("Error adding the record, posible dupliated it")
            
        sql = "INSERT INTO EMPLOYEE(FIRST_NAME, \
                   LAST_NAME, AGE, SEX, INCOME) \
                   VALUES ('%s', '%s', '%d', '%c', '%d' )" % \
                   ('Jorge1', 'Mohan', 20, 'M', 2000)
        if (db.Add_Move_Change_Data(sql)):
            print ("Record Added it....!!!")
        else:
            print ("Error adding the record, posible dupliated it")
        #------------- INSERT data into a Table <END> ---------
            
        #------------- DELETE data from a Table <BEGIN> ---------
        sql = "SELECT * FROM EMPLOYEE WHERE FIRST_NAME = '%s'" % ("Daniela1")
        if (db.Execute(sql)):
            sql = "DELETE * FROM EMPLOYEE WHERE FIRST_NAME = '%s'" % ("Daniela1")
            if (db.Add_Move_Change_Data(sql)):
                print ("Record Deleted it....")
            else:
                print ("Error adding the record, posible dupliated it")
        else:
            print ("record not Found")
        #------------- DELETE data from a Table <END> ---------

        #------------- UPDATE data from a Table <BEGIN> ---------
        sql = "SELECT * FROM EMPLOYEE WHERE FIRST_NAME = '%s'" % ("Daniela1")
        if (db.Execute(sql)):
            sql = "UPDATE EMPLOYEE SET LAST_NAME = '%s' WHERE FIRST_NAME = '%s'" % ("Rodriguez-Chavez","Daniela1")
            if (db.Add_Move_Change_Data(sql)):
                print ("Record updated it....")
            else:
                print ("Error updating the record, posible dupliated it")
        else:
            print ("record not Found")
        #------------- UPDATE data from a Table <END> ---------
        '''
        '''
        if (db.Alter_Table_Add_Field("VARIABLES","Window","FLOAT")):
            print ("!!!!!")
        else:
            print ("NO")
        '''
        print ("..............")
        db.Disconnect()
    else:
        print ("Failure")

if __name__ == '__main__':
    Main()

