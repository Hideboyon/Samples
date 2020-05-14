import pandas as pd
import pyodbc
import pandas.io.sql


#Connection script to SQL server
server = 'tcp:xx.xx.xx.xx'
database = 'xxxxxxxxxxx'
username = 'xxxxxxxxxxx'
password = 'xxxxxxxxxxx'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

#Get dataagit
df = pd.read_sql(
                #Chose colum and Select DB which is excuted some conditions to excerpt data. 
                '''SELECT TOP 150000 DATE_WRONG,DDI,CallingNumber,Passcode\
                FROM AntiFraud.dbo.ConferenceWrongAttempt where CallingNumber not like '+81%' '''
                #Connect SQL
                ,cnxn
                )

