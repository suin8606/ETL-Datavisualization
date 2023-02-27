# %%
from sklearn.linear_model import LinearRegression
import pandas as pd
import numpy as np
import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import datetime as dt
import smartsheet
from openpyxl import load_workbook, styles, formatting
import re
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt
import matplotlib.mlab as mlab
import seaborn as sns
import statsmodels.api as sm
from sklearn.preprocessing import MinMaxScaler
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
dff=pd.read_sql("""
SELECT T1.act_date, SUM(qty)/COUNT(DISTINCT T1.shiptoparty) AS qty, T1.upc
FROM [dbo].[ivy.sd.fact.pos] T1 LEFT JOIN [ivy.mm.dim.posupc] T2 ON T1.upc = T2.upc 
INNER JOIN [ivy.mm.dim.shiptoparty_pos] T3 ON T1.shiptoparty = T3.shiptoparty 
WHERE qty > 0  AND T2.LCHDATE>= DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) - 12, 0) AND T2.division NOT in ('V1','X1','Z1') AND T3.active='ACTIVE'
GROUP BY T1.act_date, T1.upc
ORDER BY T1.act_date DESC, qty DESC , T1.upc
""", con=engine)
# %%
dff.sort_values(by=["upc"])
# %%
dff["upc"] = dff["upc"].apply(lambda x: pd.to_numeric(x, errors='coerce')) # filtering for numbers only & replaceing empty value as na
dff.dropna(subset=["upc"], inplace=True) 
dff["upc"]=dff["upc"].astype('string')
dff["upc"]=dff["upc"].str.lstrip('.0')
dff['upc']=dff['upc'].str.replace("\.0", "")
# %%
dff["act_date"]=dff["act_date"].map(dt.datetime.toordinal)
# %%
dff["act_date"].unique()
# %%
from locale import normalize
def pos_slope():
    df_o=pd.DataFrame()
    for listupc in dff.upc.unique():
        y=dff[dff["upc"]==listupc]["qty"]
        x=dff[dff["upc"]==listupc]["act_date"]
        print(y)
        x=np.array(x).reshape(len(x),1)
        y=np.array(y).reshape(len(y),1)
        # scaler=MinMaxScaler(feature_range=(0,1))
        # scaler=scaler.fit(x)
        # x=scaler.transform(x)
        # print(x)
        reg=LinearRegression().fit(x,y)
        d={listupc: reg.coef_[0][0]}              
        reg_output=pd.DataFrame.from_dict(d, orient='index')
        df_o=df_o.append(reg_output)
    return df_o
df_o=pos_slope()
# %%
df_o.head()
# %%
df_o.reset_index(inplace=True)
df_o=df_o.rename(columns={'index':"upc",0:"slope"})
df_o["slope"]=df_o["slope"].apply(lambda x: '%.3f' % x).astype(float)
# %%
df_o.head()
# %%
df_o=df_o[(df_o[["slope"]]!=0).any(axis=1)].reset_index(drop=True)
# %%
df_o.sort_values(by=["slope"], ascending=True)
# %%
df_o.describe()
# %%
df_o.head()
# %%
with engine.connect() as con:
          con.execute("DELETE FROM **")
df_o.to_sql('**dim.posslope', engine, schema = "dbo", if_exists='append', index=False, chunksize=1000)
print("\n Successfully Transported")
