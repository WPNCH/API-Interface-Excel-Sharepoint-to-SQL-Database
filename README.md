# API-Interface-Excel-Sharepoint-to-SQL-Database

_Only Simple Index File (ExcelSharepointToSQLDB.py) and Config.json (for storing Authenticate Information)_
```ruby
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
import datetime
import pyodbc
import io
import pandas as pd
from openpyxl import load_workbook
import json, os

#TransformData Function
def ExcelTransform(excelfile):
    df = pd.read_excel(excelfile)
    filted_df = df[df['gacchart']!='Grand Total'].fillna(0)
    filted_df = filted_df.drop(columns='Grand Total',axis=1)
    df_unpivot = pd.melt(filted_df, id_vars=['gacchart','GACCHART2'], value_vars=['CIC', 'IMX','IT','MAT'])
    df_unpivot['Month'],df_unpivot['Day'],df_unpivot['Year'] = pd.to_datetime("today").date().month, 1, pd.to_datetime("today").date().year
    return df_unpivot

#Declare Rootpath
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
config_path = '\\'.join([ROOT_DIR, 'DatabaseConfig.json'])

# read config file
with open(config_path) as config_file:
    config = json.load(config_file)
    SPconfig = config['share_point']
    DBconfig = config['database']

#Sharepoint account
url = SPconfig['url']
SPusername = SPconfig['user']
SPpassword = SPconfig['password']
ctx_auth = AuthenticationContext(url)

#ServerInfo
server = DBconfig['server'] 
database = DBconfig['database']
username = DBconfig['username']
password = DBconfig['password']

#Connect to Database
try:
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    print('Database Connected')
except:
    print('Can not connect to Database')

#Sharepoint Authenticate
try:
    ctx_auth.acquire_token_for_user(SPusername, SPpassword)
    ctx = ClientContext(url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("SharePoint Connected")
except:
    print('Login fail')

#Write to Stream
try:
    response = File.open_binary(ctx, url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)
    print('Write to stream completed')
except:
    print('Cannot write to stream')

#TransformData
try:
    finaldf = ExcelTransform(bytes_file_obj)
    print('Transform Completed')
except:print('Transform Fail')

#Insert to SQL
try:
    for index, row in finaldf.iterrows():
        cursor.execute("INSERT INTO [Financial P&L Data] values(?,?,?,?,?,?,?,?,?)", row.gacchart, row.GACCHART2, row.Month, row.Day, row.Year,row.variable,row.value,'I24',row.variable)
    cnxn.commit()
    cursor.close()
    print('Insert Completed')
except:print('Insert fail')
```

