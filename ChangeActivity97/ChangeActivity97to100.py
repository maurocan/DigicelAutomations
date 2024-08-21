#Oracle connection Library
import cx_Oracle  #install cx_Oracle:     python -m pip install cx_Oracle --upgrade
cx_Oracle.init_oracle_client(lib_dir=r"C:\instantclient_21_9")
import time

#RobotFramework library
import subprocess
import sys
from robot.run import run_cli

import pandas as pd         #pip install numpy --upgrade
                            #pip install pandas --upgrade
                            #pip install matplotlib --upgrade
from openpyxl import workbook


#Function to create SQL query 
def BD_Conncetion():
    #credentials for DB connection
    host="my_host_ip_adress"
    user="user_name"
    passw="user_pass"
    tsname="ts_name"

    try:
        connection = cx_Oracle.connect(user, passw, host + "/" + tsname)

    except Exception as error:
            print("Unable to connect to the Data Base. Error: " + str(error))
            

    else:
            
            with open("C:\\my_computer_adress\\query2run.txt", "r") as archivo:                
                list101 = archivo.read()

            #Executing query and storing in variable
            cursor01= connection.cursor()
            cursor01.execute(list101)
            columns = [desc[0] for desc in cursor01.description]

            result= cursor01.fetchall()

            # creating Excel file        
            df = pd.DataFrame(list(result), columns=columns)            
           
            df.to_excel('C:\\my_computer_adress\\pandas80to90.xlsx', index=False)
                        
            # Close connection to the Data Base            
            connection.close()

def pandas():
            # Waiting to read the Excel file
            time.sleep(2)
            print("creating pandas")
            # Reading Excel File
            df = pd.read_excel('C:C:\\my_computer_adress\\pandas80to90.xlsx')
            # Select column first file
            last_line = df.iloc[:,2]
            text02 = str(last_line)  
            print (text02)
            last_line.to_csv('C:\\my_computer_adress\\ChangeActivity97\\Paste_OrderID_Here.txt', index=False, header=False)

            # Select column second file
            last_line = df.iloc[:,0]
            text02 = str(last_line)  
            print (text02)
            last_line.to_csv('C:\\my_computer_adress\\ChangeActivity97\\Paste_ProcessID_Here.txt', index=False, header=False)


def update():
    with open("C:\\my_computer_adress\\ChangeActivity97\\Paste_ProcessID_Here.txt", "r") as archivo:    
        lista = list(map(str.rstrip, archivo))
        
    with open("C:\\my_computer_adress\\ChangeActivity97\\updateQuerry.txt", "w") as file:
       first_element = True
       for elemento in lista:
          if first_element:
             file.write(str(elemento))
             first_element = False
          else:             
             file.write("," + str(elemento))

def Update_query() :
    print("creating last paste file")

    with open("C:\\my_computer_adress\\ChangeActivity97\\updateQuerry.txt", "r") as archivo01:    
        for linea in archivo01:
            amigo = linea
        #list(map(str.rstrip, archivo01))

    file1 = open("C:\\my_computer_adress\\ChangeActivity97\\updateQuerry.txt", "w") 
    file1.write("Update CWPACTIVITY Set status = 4 Where process_id in ( \n " 
    + str(amigo) + "\n" 
    + ") And status =1")
    file1.close 

def insert_query():    

    pasodeseado = 90

    with open("C:\\my_computer_adress\\ChangeActivity97\\Paste_OrderID_Here.txt", "r") as archivo:       
        lista = list(map(str.rstrip, archivo))
        
    with open("C:\\my_computer_adress\\ChangeActivity97\\output.txt", "w") as file:
     for elemento in lista:
        file.write("INSERT INTO CWPACTIVITY (PROCESS_ID, ACTIVITY_INDEX, LONG_DATA_ID, LONG_DATA_ID_SENT, START_TIME, START_TIME_MS, COMPLETION_TIME, END_TIME_MS, STATUS, PARALLEL_COUNT, REPEAT_COUNT, MESSAGE_TYPE) VALUES(" + elemento + ", " + str(pasodeseado) + ", 0, -1, SYSDATE, 0, NULL, 0, 1, 0, 0, NULL)" '\n')

def resume_issue():
    with open("C:\\my_computer_adress\\ChangeActivity97\\Paste_OrderID_Here.txt", "r") as archivo:    
        lista = list(map(str.rstrip, archivo))

    file1 = open("C:\\my_computer_adress\\Inputs_Outputs\\Template_Output.txt", "w")
    file1.write("var orders = " + str(lista)+ ";\n"+ "\n" 
    + "var resp = new Array(); \n\n" 
    + "for (var i = 0; i < orders.length; i++) { \n" 
    + "     var order = Order.getOrderById(orders[i]); \n" 
    + "     if (order) { \n" 
    + "         var processId = order.orderInstance.processId; \n" 
    + "         Process.resumeProcess(processId); \n"
    + "         resp.push(processId); \n"
    + "     } \n"
    + "     bundleMigration.sleepFor(300);"
    + "} \n\n"
    + "'(' + resp.length + ') Process Resumed: ' + resp")
    file1.close  

def update_connection():
    print('updating Activity...')
    #credentials for DB connection
   host="my_host_ip_adress"
    user="user_name"
    passw="user_pass"
    tsname="ts_name"

    connection = cx_Oracle.connect(user, passw, host + "/" + tsname)

    with open("C:\\my_computer_adress\\ChangeActivity97\\updateQuerry.txt", "r") as file:
        query102 = file.read()

    #Executing query and storing in variable
    cursor02= connection.cursor()
    cursor02.execute(query102)
    connection.commit()       
                      
    # end connection to the Data Base
    cursor02.close()
    connection.close()  

def output_connection():
    print('Inserting a new Activity...')
    #credentials for DB connection
    host="my_host_ip_adress"
    user="user_name"
    passw="user_pass"
    tsname="ts_name"

    connection = cx_Oracle.connect(user, passw, host + "/" + tsname)

    with open("C:\\my_computer_adress\\ChangeActivity97\\output.txt", "r") as file:
        query102 = file.read()

    sql_lines = query102.splitlines()

    #Executing query line by line
    for line in sql_lines:
        cursor02= connection.cursor()
        cursor02.execute(line)
        cursor02.close()

    # stop connection to the Data Base    
    connection.close()  

def robot():
     print("Runing Robot...")
     ruta_archivo = "C:\\my_computer_adress\\Terminate_Resume_Robot\\tasks.robot"
     run_cli([ruta_archivo])


BD_Conncetion()
pandas()
update()
insert_query()
resume_issue()
Update_query()
update_connection()
output_connection()
robot()
