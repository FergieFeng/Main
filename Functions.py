sysPath = 'XXXFolderPath'
import sys
sys.path.append(sysPath)
import pyodbc
from sqlalchemy import create_engine 
import win32com.client as win32 
import glob

def getfNameNextensions(ipath,iExt):
    path = (sysPath+iPath) # use your path
    all_files = glob.glob(path+"\*."+iExt)
    return all_files

def sendEmail (to,sub,body) :
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = sub
    mail.Body - body
    mail.Send()

def executesQLsCriptb(SQLScript):
    engine = create_engine ('mssql+pyodbc://usr:passw@srv/db?driver-SQLServer',echo=True)
    with engine.begin() as conn:
        conn.execute(SQLScript)

def executeSQLscript (SQLScript) :
    conn = pyodbc. connect ('Driver={SOL Server}; Server=XXX; Database=XXX; Trusted_connection=yes;')
    cursor = conn.cursor ()
    cursor.execute (SQLScript)
    conn.commit ()

def updateLogFile():
    print('logfile')