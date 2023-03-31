# %%
import adodbapi
import numpy as np
import pandas as pd
import re
import datetime as dt
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
# %%
#Cleans up raw data format into DataFrame
def get_df(data):
    ar = np.array(data.ado_results)
    df = pd.DataFrame(ar).transpose() 
    df.columns = data.columnNames.keys()
    return df

#User inputs
Dataset_Link = 'powerbi://api.powerbi.com/v1.0/myorg/SOM_DB_Management'
# Tried loop Failed lol..
# Dataset_Name = ['BO', 'BL', 'SO', 'RT']
# Table_Name = ['BO', 'BL', 'SO', 'RT']
# %%
#Connecting to dataset and executing query
# connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog='+Dataset_Name
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=BO'
query = 'EVALUATE BO'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
# %%
df['act_date']=df['act_date'].dt.tz_convert(None)
df_bo = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=BL'
query = 'EVALUATE BL'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
# %%
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['act_date']=df['act_date'].dt.tz_convert(None)
df_bl = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=SO'
query = 'EVALUATE SO'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['act_date']=df['act_date'].dt.tz_convert(None)
df_so = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=RT'
query = 'EVALUATE RT'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['act_date']=df['act_date'].dt.tz_convert(None)
df_rt = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=VT'
query = 'EVALUATE VT'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['act_date']=df['act_date'].dt.tz_convert(None)
df_vt = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=AR'
query = 'EVALUATE AR'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)        
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['doc_date']=df['doc_date'].dt.tz_convert(None)
df_ar = df.copy()
# %%
connstr = 'Provider=MSOLAP.8; Data Source='+Dataset_Link+'; Initial Catalog=OO'
query = 'EVALUATE OO'
with adodbapi.connect(connstr) as conn:
    with conn.cursor() as cur:              
        cur.execute(query)  
        data = cur.fetchall()
        df = get_df(data)
    #Post-processing DataFrame column headers        
for column_name in list(df.columns.values):
    newname = re.findall(r'\[.*?\]',column_name)[0]
    for i in ["[","]"]:
        newname = newname.replace(i,"")
    df.rename(columns={column_name:newname},inplace=True)
df['act_date']=df['act_date'].dt.tz_convert(None)
df_oo = df.copy()
# %%   
bl_md = "'"+df_bl['act_date'].min().strftime("%Y-%m-%d")+"'"
bl_ld =  "'"+df_bl['act_date'].max().strftime("%Y-%m-%d")+"'"
so_md = "'"+df_so['act_date'].min().strftime("%Y-%m-%d")+"'"
so_ld =  "'"+df_so['act_date'].max().strftime("%Y-%m-%d")+"'"
bo_md = "'"+df_bo['act_date'].min().strftime("%Y-%m-%d")+"'"
bo_ld =  "'"+df_bo['act_date'].max().strftime("%Y-%m-%d")+"'"
rt_md = "'"+df_rt['act_date'].min().strftime("%Y-%m-%d")+"'"
rt_ld =  "'"+df_rt['act_date'].max().strftime("%Y-%m-%d")+"'"
vt_md = "'"+df_vt['act_date'].min().strftime("%Y-%m-%d")+"'"
vt_ld =  "'"+df_vt['act_date'].max().strftime("%Y-%m-%d")+"'"
oo_md = "'"+df_oo['act_date'].min().strftime("%Y-%m-%d")+"'"
oo_ld =  "'"+df_oo['act_date'].max().strftime("%Y-%m-%d")+"'"
# %%
# Server Connection
server = '10.1.3.25' 
database = 'KIRA' 
username = 'kiradba' 
password = 'Kiss!234!' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url, fast_executemany=True)
# %%
# Specify dtype for dataframe
# Trim Columns
#########################################################################################################
df_so.drop(columns=['ordreason_desc'], inplace=True)
df_so.rename(columns= {'gross': 'gross_amt', 'ordreason': 'ord_reason'}, inplace=True)
# %%
# DELETE existing data from Sales Order table and INSERT NEW DATA INTO Sales Order Table
query_so = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.order] WHERE act_date BETWEEN {0} AND {1}'.format(so_md, so_ld)
with engine.connect() as con:    
    con.execute(query_so)
df_so.to_sql('ivy.sd.fact.order', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("Sales Order Completed")
# %%
#########################################################################################################
# DELETE existing data from Sales Billing table and INSERT NEW DATA INTO Sales Billing Table
query_bl = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.bill] WHERE act_date BETWEEN {0} AND {1}'.format(bl_md, bl_ld)
with engine.connect() as con:    
    con.execute(query_bl)
df_bl.to_sql('ivy.sd.fact.bill', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("Sales Billing Completed")
# %%
#########################################################################################################
# DELETE existing data from BO table and INSERT NEW DATA INTO BO Table
query_bo = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.bo] WHERE act_date BETWEEN {0} AND {1}'.format(bo_md, bo_ld)
with engine.connect() as con:    
    con.execute(query_bo)
df_bo.to_sql('ivy.sd.fact.bo', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("BO Completed")
# %%
#########################################################################################################
# DELETE existing data from Return table and INSERT NEW DATA INTO Return Table
query_rt = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.return] WHERE act_date BETWEEN {0} AND {1}'.format(rt_md, rt_ld)
with engine.connect() as con:    
    con.execute(query_rt)
df_rt.to_sql('ivy.sd.fact.return', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("Return Completed")
# %%
#########################################################################################################
# %%
# DELETE existing data from Return table and INSERT NEW DATA INTO Return Table
query_vt = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.visit] WHERE act_date BETWEEN {0} AND {1}'.format(vt_md, vt_ld)
with engine.connect() as con:    
    con.execute(query_vt)
df_vt.to_sql('ivy.sd.fact.visit', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("Visit Completed")
# %%
#########################################################################################################
# %%
# DELETE existing data from AR table and INSERT NEW DATA INTO Return Table
query_ar = 'DELETE FROM [KIRA].[dbo].[ivy.sd.fact.ar]'
with engine.connect() as con:    
    con.execute(query_ar)
df_ar.to_sql('ivy.sd.fact.ar', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("AR Completed")
# %%
#########################################################################################################
# DELETE existing data from Open Order table and INSERT NEW DATA INTO Open Order Table
query_oo = 'DELETE FROM [KIRA].[dbo].[ivy.mm.dim.open_order] WHERE act_date BETWEEN {0} AND {1}'.format(oo_md, oo_ld)
with engine.connect() as con:    
    con.execute(query_oo)
df_oo.to_sql('ivy.mm.dim.open_order', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("Open Order Completed")
