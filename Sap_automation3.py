# %%
import pandas as pd
import numpy as np
import pyodbc
import openpyxl
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import datetime as dt
import re
import glob
import os
from pathlib import Path
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook, styles, formatting
from pickle import TRUE
import sys
import win32com.client
import subprocess
from datetime import datetime
import time
import shutil
# %%
from calendar import month
from datetime import date, timedelta

# %%
shell = win32com.client.Dispatch("WScript.Shell")
# Log in SAP QA for now / will be changed to R3 with system=P01
subprocess.check_call(['C:\Program Files (x86)\SAP\FrontEnd\SAPgui\\sapshcut.exe', '-system=**', '-client=**', '-user=**', '-pw=**', 
'-command=V/LD', 
# '-command=ZPPRMRP01', 
'-type=Transaction', '-max'])
time.sleep(5)
def main():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not isinstance(SapGuiAuto, win32com.client.CDispatch):
            return

        application = SapGuiAuto.GetScriptingEngine
        if not isinstance(application, win32com.client.CDispatch):
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not isinstance(connection, win32com.client.CDispatch):
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not isinstance(session, win32com.client.CDispatch):
            connection = None
            application = None
            SapGuiAuto = None
            return
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "V/LD"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtRV14A-KONLI").text = "ZA"
        session.findById("wnd[0]/usr/ctxtRV14A-KONLI").caretPosition = 1
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/usr/btn%_P_1_%_APP_%-VALU_PUSH").press
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1100"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1400"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[0]/usr/chkPAR_DAT").selected = true
        session.findById("wnd[0]/usr/ctxtP_2-LOW").text = "TR"
        session.findById("wnd[0]/usr/txtP_3-LOW").text = "9"
        session.findById("wnd[0]/usr/ctxtL_1-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtKSCHL-LOW").text = "ZD10"
        session.findById("wnd[0]/usr/ctxtDATUM-LOW").text = "01/10/2022"
        session.findById("wnd[0]/usr/txtMAX_LINE").text = "0"
        session.findById("wnd[0]/usr/txtMAX_LINE").setFocus
        session.findById("wnd[0]/usr/txtMAX_LINE").caretPosition = 7
        session.findById("wnd[0]/tbar[1]/btn[8]").press
    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


if __name__ == "__main__":
    main()

# %%
name=['a1','a2','a3','a4','a5','a6','a7','a8','a9','a10','a11','a12','a13','a14','a15','a16','a17','a18','a19','a20','a21']
col=["CnTy","Material","a1","Scale quantity","UoM","Amount","UoM","Valid From","NaN","Valid to"]
df=pd.read_csv(r"\promotion.txt",delimiter="	",names=name,usecols=list(range(1,21,+1)),header=None,engine='python')
df=df[4:]

# %%
df.head()

# %%
df=df[["a2","a3","a9","a10","a12","a13","a18","a19","a20","a21"]]
df.drop([4],inplace=True)
df=df.reset_index(drop=True)
df.rename(columns={"a2":"CnTy","a3":"Material","a9":"a1","a10":"Scale quantity","a12":"UoM1","a13":"Amount","a18":"UoM2","a19":"Valid From","a20":"a2","a21":"Valid to"},inplace=True)

# %%
df.head()

# %%
df["Scale quantity"]=df["Scale quantity"].fillna(df["a1"])
df["Amount"]=df["Amount"].fillna(df["UoM1"])
df["Valid From"]=df["Valid From"].fillna(df["UoM2"])
df["Valid to"]=df["Valid to"].fillna(df["a2"])
df["Valid From"]=pd.to_datetime(df["Valid From"],errors='coerce').dt.strftime('%Y-%m-%d')
df["Valid to"]=pd.to_datetime(df["Valid to"],errors='coerce').dt.strftime('%Y-%m-%d')

# %%
# def func(x):
#     if x["Scale quantity"] is np.nan:
#         x["Scale quantity"]=x["Scale quantity"].fillna(x["a1"])
#     elif x["Amount"] == np.nan:
#         x["Amount"]=x["Amount"].fillna(x["UoM1"])
#     elif x["Valid From"] is np.nan:
#         x["Valid From"]=x["Valid From"].fillna(x["UoM2"])
#     elif x["Valid to"] is np.nan:
#         x["Valid to"]=x["Valid to"].fillna(x["a2"])
#     else:
#         return x

# %%
def read_csv():
    df["CnTy"]=df["CnTy"].astype('string')
    df["Material"]=df["Material"].astype('string')
    df["Scale quantity"]=pd.to_numeric(df["Scale quantity"],errors='coerce')
    df["Amount"]=df["Amount"].astype('string')
    return df
df_final=read_csv()
df_final=df_final[["CnTy","Material","Scale quantity","Amount","Valid From","Valid to"]]

# %%
df["Valid From"]=pd.to_datetime(df["Valid From"], format="%Y-%m-%d",errors='coerce').dt.strftime('%Y-%m-%d')
df["Valid to"]=pd.to_datetime(df["Valid to"],format="%Y-%m-%d",errors='coerce').dt.strftime('%Y-%m-%d')
df["Material"]=df["Material"].fillna(method='ffill')
df["CnTy"]=df["CnTy"].fillna(method='ffill')
df["Valid From"]=df["Valid From"].fillna(method='ffill')
df["Scale quantity"]=df["Scale quantity"].fillna(1)
df["Valid to"]=df["Valid to"].fillna(method='ffill')
df_final=df[["CnTy", "Material","Scale quantity","Amount","Valid From","Valid to"]]
#df_final.insert(loc=4,column='Unit', value=['%' for i in range(df_final.shape[0])])
df_final["Valid From - Copy"]=pd.to_datetime(df_final["Valid From"], format="%Y-%m-%d",errors='coerce').dt.strftime('%Y-%m-01')
df_final["Valid to - Copy"]=pd.to_datetime(df_final["Valid to"], format="%Y-%m-%d",errors='coerce').dt.strftime('%Y-%m-01')
df_final["Valid From - Copy"]=pd.to_datetime(df_final["Valid From - Copy"], format = '%Y-%m-%d')
df_final["Valid to - Copy"]=pd.to_datetime(df_final["Valid to - Copy"], format = '%Y-%m-%d')

# %%
df_final=df_final[~df_final["Amount"].isna()]
df_final["Amount"]=pd.to_numeric(df_final["Amount"],errors='coerce')
df_final=df_final[df_final["Amount"]!=0.0]
df_final=df_final[df_final["CnTy"]!='CnTy']
df_final["cur_date"]=datetime.now()

# %%
df_final.info()
# %%
df_final

# %%
df_final["valid_to"].unique()

# %%
server = '**' 
database = '**' 
username = '**' 
password = '**' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)
print("connected")

# %%
with engine.connect() as con:
          con.execute("DELETE FROM **")
df_final.to_sql('**', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("\n Successfully Transported")


