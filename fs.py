import pandas as pd
import sqlite3
import numpy as np

months  =['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember']
months_en  =['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def getAccount(dbname,tablename,numcol=0):
    # Read sqlite query results into a pandas DataFrame
    con = sqlite3.connect(dbname)
    if numcol==0:
        query='select distinct Trim(Account) as Account from {} where not Account is null'.format(tablename)
    else:
        query='select distinct Trim(Account) as Account from {} where not Account is null limit {}'.format(tablename,numcol)
    #get dataframe from sqlite
    df=pd.read_sql_query(query, con)
    con.close()
    return df

def getMonthIdx(mth):
    a = np.array(months)
    return np.where(a==mth)[0][0]+1

def getAccountsStr(accountname):
    acc_item =''
    for acct in accountname:
        acc_item+="'" + acct + "',"
    acc_item= acc_item[:len(acc_item)-1]
    return acc_item

def generateRows(df,yr):
    df_new = pd.DataFrame()
    periodes=[]
    values=[]
    for mth in months:
        idxMonth = getMonthIdx(mth)
        periodes.append('{}-{}'.format(yr,str(idxMonth).zfill(2)))
    #print(periodes)

    df['Periode']=periodes    
    #swap column periode to be the first column
    first_col=df.pop('Periode')
    df.insert(0,'Periode',first_col)
    #convert ke datetime untuk columns periode
    df["Periode"] = pd.to_datetime(df["Periode"],format='%Y-%m')
    #ganti angka 0 dengan Nan supaya grafik nya tidak nyungsep
    df.replace(0, np.nan, inplace=True)
    #print(df)
    return df


def getAccountByYear(dbname,tablename,start_period,end_period,accountname):
    # Read sqlite query results into a pandas DataFrame
    con = sqlite3.connect(dbname)
    #create empty dataframe
    df_final=pd.DataFrame()
    
    #build filter
    filter = '{}'.format(getAccountsStr(accountname))
    for yr in range(start_period,end_period+1):
        #query statement to sqlite..
        query = "SELECT report_date,Januari,Februari,Maret,April,Mei,Juni,Juli,Agustus,September,Oktober,November,Desember from {} where strftime('%m',report_date)='12' and trim(Account) in ({}) and strftime('%Y',report_date)='{}' order by report_date asc".format(tablename,filter,yr)
        if yr==2023:
            query="SELECT report_date,Januari,Februari,Maret,April,Mei,Juni,Juli,Agustus,September,Oktober,November,Desember from {} where trim(Account) in ({}) and  strftime('%Y',report_date)='{}' and  strftime('%m',report_date)='09' order by report_date asc".format(tablename,filter,yr)
        #get dataframe from sqlite
        df=pd.read_sql_query(query, con)
        df=df.fillna(0)
        print(df)
        #transpose first
        df =df.transpose()
        print(df)
        
        df = df.iloc[1:]
        df.columns = accountname
        #print(df)
        #start generate rows vertical Periode | Account Name
        df = generateRows(df,yr)
        df_final = pd.concat([df_final, df],ignore_index = True)

    return con,df_final

def plot(df):
    # importing matplotlib library 
    # importing packages 
    import seaborn as sns 
    import matplotlib.pyplot as plt 
    # plotting a line graph 
    print("Line graph: ") 
    sns.lineplot(df)
    plt.show()

def plotly(df):
    import plotly.express as px
    fig = px.line(df,x='Periode',y=df.columns,title='Financial Statement')
    fig.show()

