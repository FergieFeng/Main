
import pyodbe 
import pandas 
import re
import win32com.client as win32 
import xlwings as xw #import os
from datetime import date 
from datetime import datetime 
from datetime import timedelta, date #pyodbc.drivers() #for check driver name

# MANUAL STEP: define the US you would like to run
#format: Table you would like to update, followed by stored procedure name, followed by date name (help check)
 
usp_names = ['XXX','XXXX','XXXXX']
#connect to SQL
conn = pyodbc.connect(
    Trusted_Connection='Yes',
    Driver='{SOL Server Native Client 11.0}',
    Server='XXX'
    Database='XXX'
)
cursor = conn.cursor ()
# Set Email Message List
Msg_Email = []
Query = '''select max(cast(XXX as date)) from XXXX'''
datediff = str(pandas.read_sql_query(query,conn))
sql_word = '   \n0   '
text = 'Data Source Updated to' + datediff.partition(sql_word)[2]
print(text)
Msg_Email.append(text)

a=0
for b in range(0,len(usp_names)):
    if a >= len(usp_names):
        temp1 = 'Everything is done! Note: ' +str(int(len(usp_names)/3))+'USPs FInished'
        print(temp1)
        Msg_Email.append(temp1)
        break
    else:
        now = datetime.now()strftime("%Y/%m/%d %H:%M:%S")
        check_date = usp_names[a+2]
        query = 'select datediff(day,'+'max('+check_date + '),getdate()) from' + usp_names[a]
        sql_word = '    \0'
        loop_num = int(datediff.partition(sql_word)[2])

        append_date = str(date.today()+timedelta(days = -1))
        if loop_num ==1:
            temp3 = 'Done, no need to refresh'
            Msg_Email.append(temp3)
            print(temp3)
        else:
            for counter in range(1,loop_num):
                append_date = str(date.today()+timrdelta(days = -counter))
                counter = counter +1

                usp = 'EXECUTE' + usp_names[a+1] + "@refresh_date =" + "'" + append_date + "'"
                cursor.execute (usp)
                conn. commit ()
                temp4= 'Done for' + append_date + ':' + 'update table ' + usp_names[a] + 'by call' + usp_names[a+1]
        a=a+3
conn.close()

# Use functions build 
sendemail("email address","python scriptXXX",Plain)