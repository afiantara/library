from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd


url_path = './ojk/(revisi) 3.1 s.d. 3.11 Tabel Keuangan Asuransi Jiwa 2022.xlsx'
pre_title = 'REKAPITULASI'

def load():
    workbook = load_workbook(url_path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet_title = sheet.title.strip()
        #title = sheet_title.split(maxsplit=1)
        tablename=sheet_title.replace(' ','').replace('.','')
        tablename = ''.join([i for i in tablename if not i.isdigit()])
        print(tablename)
        if tablename=='Neraca':
            load_neraca(sheet_title,tablename)
            print(tablename)
        elif tablename=='Investasi':
            print(tablename)
            load_investasi(sheet_title,tablename)
        elif tablename=='HasilInvestasi':
            load_hasil_investasi(sheet_title,tablename)
        elif tablename=='NeracaUL':
            load_neraca_ul(sheet_title,tablename)
        elif tablename=='LabaRugi':
            load_rl(sheet_title,tablename)      
        elif tablename=='LRUL':
            load_rl_ul(sheet_title,tablename)            

def load_rl_ul(sheet_name,table_name):
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    columns = getcolumns(df,10)
    arrcols = []
    for column in columns[:-1]:
        arrcols.append(column)
    arrcols.append('INCREASE (DECREASE) IN ASSET')
    columns=arrcols
    print(date,columns)
    df= getdata(df,13,columns,1)

    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)

def load_rl(sheet_name,table_name):
    columns=['NAME OF COMPANY','NET EARNED PREMIUM','INVESTMENT YIELDS','OTHER INCOME','TOTAL INCOME',
        'NET CLAIM EXPENSES','OPERATIONAL EXPENSES','OTHER EXPENSES','TOTAL EXPENSES','EARNING BEFORE TAX',
        'TAX','EARNING AFTER TAX', 'OTHER COMPREHENSIVE INCOME','TOTAL OTHER COMPREHENSIVE INCOME']
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    print(date,columns)
    df= getdata(df,14,columns,1)
    print(df)
    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)

def load_neraca_ul(sheet_name,table_name):
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    columns = getcolumns(df,10)
    print(date,columns)
    df= getdata(df,13,columns,1)

    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)

def load_hasil_investasi(sheet_name,table_name):
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    columns = getcolumns(df,8)
    df= getdata(df,13,columns,1)

    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)

def load_investasi(sheet_name,table_name):
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    columns = getcolumns(df,8)
    df= getdata(df,13,columns)
    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)
    
def load_neraca(sheet_name,table_name):
    df = pd.read_excel(url_path,sheet_name=sheet_name)
    date = getReportDate(df,3)
    columns = getcolumns(df,8)
    df= getdata(df,13,columns)
    df.insert(0,'REPORT_DATE',date)
    df=df.assign(PRICE_INDEX=1000000)
    dump_to_db('ASURANSI',table_name.upper(),df,date)

def getReportDate(df,start):
    import datetime
    year = int(df.iloc[start].dropna().values[0].split()[-1])
    last_day_of_year = datetime.date.max.replace(year = year)
    return last_day_of_year

def getcolumns(df,start):
    columns = df.iloc[start].dropna().values
    return columns

def getdata(df,start_row,columns,start_col=2):
    df = df[start_row:]
    df = df.iloc[:,start_col:]
    df.columns = columns
    df=df.dropna()
    df=df[df[columns[0]].str.contains( "JUMLAH" )==False]
    return df

def dump_to_db(db_name,table_name,df,report_date):
    from sqlalchemy import create_engine
    db = create_engine('sqlite:///{}.db'.format(db_name))
    # drop 'records table'
    try:
        query="DELETE FROM {} WHERE REPORT_DATE='{}'".format(table_name,report_date)
        print(query)
        db.execute(query)
    except:
        print('ignore it')
    finally:
        df.to_sql(table_name, db, if_exists='append')

if __name__=="__main__":
    load()