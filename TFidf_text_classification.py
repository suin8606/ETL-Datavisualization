# %%
import pandas as pd
import glob
import os
import pyodbc
from sqlalchemy import column, create_engine
from sqlalchemy.engine import URL

# %%
server = '**' 
database = '**' 
username = '**' 
password = '**' 
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url, fast_executemany=True)
print("Connection Established:")

# %%
df= pd.read_sql(r'SELECT*FROM [ivy.mm.dim.posupc]', con=engine)

# %%
fpath = r'C:\Users\KISS Admin\Documents\Python Scripts'
fnamelist = os.listdir(fpath)

# %%
def raw_etl(path, desc, upc, sname):
    df = pd.read_excel(path, sheet_name=sname)
    df[[upc, desc]] = df[[upc, desc]].astype(str)
    df[upc] = df[upc].str.lstrip('0')
    df = df[df[upc].str.isnumeric()]
    df = df[[upc, desc]]        
    return df

# %%
def fn(fpath, desc, upc):    
    df_final = pd.DataFrame()
    fnamelist = os.listdir(fpath)
    for j in fnamelist:
        snamelist = pd.ExcelFile(fpath + '\\' +j).sheet_names    
        if len(snamelist) > 1:
            for i in snamelist:
                df = raw_etl(fpath + '\\' +j, desc, upc, i)
                df_final = df_final.append(df)
        else:
            df = raw_etl(fpath + '\\' +j, desc, upc, snamelist[0])
            df_final = df_final.append(df)
    df_final.rename(columns={upc: 'upc', desc: 'desc'}, inplace=True)
    return df_final

# %%
# Foxx (C&H beauty)
df_hst=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\FOXX\FOXX DATA JAN 2023.xlsx", engine='openpyxl')
df_hst = df_hst.iloc[df_hst[df_hst['Unnamed: 0']=='UPC'].index[0]:, :]
df_hst.columns=df_hst.iloc[0]
df_hst=df_hst[1:]
df_hst = df_hst[['Description', 'UPC']]
df_hst.rename(columns={'Description':'d', 'UPC':'u'}, inplace=True)

# %%
df_hst.tail()

# %%
# KC - new
def stpassgn(df):
    if df['store'] == 'ARUNDEL':
        return '0011015975'
    elif df['store'] == 'HM':
        return '0011013949'
    elif df["store"]=='FC':
        return '0011016803'    
    elif df['store'] == 'KC':
        return '0011016100'
    elif df['store'] == 'MH':
        return '0011011351'
    elif df['store'] == 'TAKOMA':
        return '0011014190'
    else:
        return ''    
prlst = os.listdir(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\KC\012023")
df_kc = pd.DataFrame()
for i in prlst:
   df=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\KC\012023\\" + i, dtype={'UPC': 'str'}, engine='openpyxl')
   df_kc = df_kc.append(df)
   df_kc=df_kc[['Description', 'UPC']]
   df_kc.rename(columns={'Description':'d', 'UPC':'u'}, inplace=True)

# %%
# KC - new
def stpassgn(df):
    if df['store'] == 'ARUNDEL':
        return '0011015975'
    elif df['store'] == 'HM':
        return '0011013949'
    elif df["store"]=='FC':
        return '0011016803'    
    elif df['store'] == 'KC':
        return '0011016100'
    elif df['store'] == 'MH':
        return '0011011351'
    elif df['store'] == 'TAKOMA':
        return '0011014190'
    else:
        return ''    
prlst = os.listdir(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\KC\012023")
df_kc = pd.DataFrame()
for i in prlst:
   df=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\KC\012023\\" + i, dtype={'UPC': 'str'}, engine='openpyxl')
   df_kc = df_kc.append(df,ignore_index=True)
df_kc=df_kc[['Description', 'UPC']]
df_kc.rename(columns={'Description':'d', 'UPC':'u'}, inplace=True)

# %%
df_kc

# %%
df_mi=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\Maxx\JAN2023.xlsx", engine='openpyxl')
df_mi = df_mi[['Item Number', 'Item Name']]
df_mi.rename(columns={'Item Name':'d', 'Item Number':'u'}, inplace=True)
df_mi=df_mi[["d","u"]]

# %%
df_mi.head()

# %%
# Beautopia - new
path = r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\Beautopia\Beautopia 2023 JAN.xlsx"
df_bt = pd.read_excel(path)    
df_bt.columns = df_bt.iloc[0]
df_bt = df_bt.iloc[1:,:]    
df_bt = df_bt.iloc[:, [0, 1]]
df_bt.columns = ['u', 'd']
df_bt = df_bt[['d', 'u']]

# %%
df_bt.head(10)

# %%
# AG 
df_ag1=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\AG\JAN_2023_DATA.xlsx", engine='openpyxl')
df_ag2=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\AG\NOV_DATA_2022.xlsx", engine='openpyxl')
df_ag3=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\AG\OCT_DATA_2022.xlsx", engine='openpyxl')
frame=[df_ag1,df_ag2,df_ag3]
df_ag=pd.concat(frame)
df_ag = df_ag[['Description', 'UPC']]
df_ag.rename(columns={'Description':'d', 'UPC':'u'}, inplace=True)

# %%
df_ag.tail()

# %%
df_bsw=pd.read_excel(r'Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\BSW\BSW Sales Report 2022-11.xlsx', engine='openpyxl')
df_bsw2=pd.read_excel(r'Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\BSW\BSW Sales Report 2022-12.xlsx', engine='openpyxl')
df_bsw3=pd.read_excel(r'Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\BSW\BSW Sales Report 2023-01.xlsx', engine='openpyxl')
bsw_frame=[df_bsw,df_bsw2,df_bsw3]
bsw_bsw=pd.concat(bsw_frame)
df_bsw = bsw_bsw[['name', 'UPC']]
df_bsw.rename(columns={'name':'d', 'UPC':'u'}, inplace=True)

# %%
df_bsw

# %%
# Charmiss - NJ
cham=os.listdir(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\Charmiss")
df_nj=pd.DataFrame()
for x in cham:
    df=pd.read_excel(r"Z:\Ivykiss Artwork\ZZ_TEMP\SOM\13_POS\3.RAW\Charmiss\\" + x, engine='openpyxl')
    df_nj=df_nj.append(df)
df_nj.columns=range(df_nj.shape[1])
df_nj = df_nj[[3, 4]]
df_nj.rename(columns={4:'d', 3:'u'}, inplace=True)


# %%
df_nj.info()

# %%
frame = [df_hst, df_kc, df_mi, df_bt, df_ag, df_bsw]
fdf = pd.concat(frame)

# %%
fdf.info()

# %%
fdf.tail(10)

# %%
fdf = fdf[~fdf['u'].isna()]
fdf = fdf[pd.to_numeric(fdf['u'], errors='coerce').notnull()]
fdf['u'] = fdf['u'].str.lstrip('0')
fdf['desclen'] = fdf['d'].str.len()
fdf.dropna(inplace=True)
fdf = fdf.sort_values(by='desclen', ascending=False)
fdf.drop_duplicates(subset='u', keep='first', inplace=True)

# %%
fdf

# %%
df_dimupc=pd.read_sql(r'SELECT upc, description, division FROM [ivy.mm.dim.posupc]',con=engine)

# %%
df_dimupc.head()

# %%
# Compare current upc with new upcs
dff = df_dimupc.merge(fdf, left_on='upc', right_on='u', how='outer')
dff['dimdlen'] = dff['description'].str.len()
dff['ndlen'] = dff['d'].str.len()
dff['fdesc'] = dff.apply(lambda x: x['description'] if x['dimdlen'] > x['ndlen'] else x['d'], axis=1)

# %%
dff

# %%
# for updating CURRENT upcs
df_cr_upc = dff[['upc', 'fdesc']]
df_cr_upc['fdesc'] = df_cr_upc['fdesc'].str.strip()
df_cr_upc = df_cr_upc[~df_cr_upc['fdesc'].isna()]
df_cr_upc.to_sql('cr_upc_desc', if_exists='replace', con=engine)

# %%
df_cr_upc

# %%
new_data_query = '''
UPDATE T1
SET 
T1.LCHDATE = T2.LCHDATE
FROM [ivy.mm.dim.posupc] T1 
INNER JOIN (
    SELECT UPC, MIN(ACT_DATE) LCHDATE
    FROM [dbo].[ivy.sd.fact.pos]
    GROUP BY UPC) AS T2 ON T1.upc = T2.upc
WHERE T1.upc = T2.upc;
UPDATE T1
SET T1.[description] = T2.fdesc
FROM [ivy.mm.dim.posupc] AS T1 
INNER JOIN cr_upc_desc AS T2 ON T1.upc = T2.upc

UPDATE T1
SET 
T1.description = T2.description,
T1.kiss = T2.material,
T1.brand = T2.brand,
T1.division = T2.division,
T1.mg = T2.mg,
T1.company='KISS'
FROM [ivy.mm.dim.posupc] T1 INNER JOIN (SELECT*
    FROM [ivy.mm.dim.mtrl]
    WHERE UPC IN (SELECT UPC
    FROM [ivy.mm.dim.mtrl]
    GROUP BY UPC
    HAVING COUNT(*)=1)) T2
    ON T1.upc = T2.upc
WHERE T1.upc = T2.upc;
'''

# %%
#query to update current upcs
with engine.connect() as con:
  con.execute(new_data_query)
# for updating NEW upcs
# Import base packages

# %%
import re
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.metrics import accuracy_score
from sklearn.multiclass import OneVsRestClassifier
from nltk.corpus import stopwords
from nltk import word_tokenize
from sklearn.svm import LinearSVC
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline

# %%
# Read dataset from sql server
pd_dimupc = pd.read_sql("SELECT*FROM [ivy.mm.dim.posupc]", engine)
pd_dimupc = pd_dimupc.sample(frac = 0.8)
# Check data head with columns
pd_dimupc = pd_dimupc[['upc', 'description', 'division']]
dataset = pd_dimupc
dataset = dataset[~dataset['description'].isnull()]
dataset['division_id'] = dataset['division'].factorize()[0]
division_id_df = dataset[['division_id', 'division']].drop_duplicates().sort_values('division_id')
division_to_id = dict(division_id_df.values)
id_to_division = dict(division_id_df[['division_id', 'division']].values)
from sklearn.feature_extraction.text import TfidfVectorizer
tfidf = TfidfVectorizer(sublinear_tf=True, min_df=5, norm='l2', encoding='latin-1', ngram_range=(1, 2), stop_words='english')
features = tfidf.fit_transform(dataset.description).toarray()
labels = dataset.division
features.shape
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.naive_bayes import MultinomialNB
X_train, X_test, y_train, y_test = train_test_split(dataset['description'], dataset['division'], random_state = 0)
count_vect = CountVectorizer()
X_train_counts = count_vect.fit_transform(X_train)
tfidf_transformer = TfidfTransformer()
X_train_tfidf = tfidf_transformer.fit_transform(X_train_counts)
clf = MultinomialNB().fit(X_train_tfidf, y_train)
# %%
df_new = dff[(dff['ndlen']>0) & (dff['division'].isna())]
df_new = df_new[['u', 'fdesc']]
# %%
def machine_div(desc):
    return clf.predict(count_vect.transform([desc]))
df_new['pred_div'] = df_new['fdesc'].apply(machine_div).str.get(0)
df_new.columns = ['upc', 'description', 'division']
# %%
# import datetime 
# cur_date=datetime.date.today()
# %%
df_new = df_new[['upc','description','division']]

# %%
df_new

# %%
df_new.to_sql("**", con=engine, if_exists='append', index=False)


