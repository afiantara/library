import pandas as pd
import sqlite3
import numpy as np

class BalanceSheet:
    global con
    description = "Balance Sheet"
    def __init__(self, dbname,tblname):
        self.dbname = dbname
        self.tblname = tblname
    
    def connectDB(self):
        # Read sqlite query results into a pandas DataFrame
        try:
            con = sqlite3.connect(self.dbname)
            self.con = con
            return True
        except:
            return False

    def closeDB(self):
        self.con.close()

    def getAccount(self):
        query = 'SELECT distinct [Balance Sheet] from {}'.format(self.tblname)
        df=pd.read_sql_query(query, self.con)
        return df

    def getTicker(self):
        query = 'SELECT distinct ticker_code,ticker_desc from {}'.format(self.tblname)
        df=pd.read_sql_query(query, self.con)
        return df

    def setAccounts(self,accounts):
        self.accounts = accounts

    def setTickerName(self,ticker):
        self.ticker = ticker

    def generateYearCols(self):
        colyears=[]
        for yr in range(self.fromyear,self.toyear+1):
            colyears.append('[' + str(yr) + ']')
        return ",".join(str(x) for x in colyears)

    def generateColumns(self):
        tickers=self.ticker
        accounts = self.accounts
        cols =[]
        for tiker in tickers:
            for account in accounts:
                cols.append(tiker + ' - ' + account)
        #print(cols)
        return cols

    def convert_string_to_decimal(self,df):
        for acct in self.accounts:
            df[acct] = df[acct].astype(float)
        return df

    def generateFilters(self,filters):
        items=[]
        for filter in filters:
            items.append("'" + filter + "'")
        return ",".join(str(x) for x in items)

    def addFirstCol(self,df):
        df['Periode']=df.index    
        #swap column periode to be the first column
        first_col=df.pop('Periode')
        df.insert(0,'Periode',first_col)
        #convert ke datetime untuk columns periode
        #df["Periode"] = pd.to_datetime(df["Periode"],format='%Y')
        return df
    
    def appendDF(self,df,dfAppend):
        return pd.concat((df, dfAppend))

    def getBalanceSheet(self, fromyear,toyear):
        self.fromyear=fromyear
        self.toyear=toyear
        fields = self.generateYearCols()
        accounts = self.generateFilters(self.accounts)
        tickers = self.generateFilters(self.ticker)
        
        new_df = pd.DataFrame()

        for ticker in self.ticker:
            ticker_item = "'" + ticker + "'"
            query ='SELECT ticker_code,[Balance Sheet],{} from {}  WHERE Trim([Balance Sheet]) in ({}) and Trim(ticker_code) in ({})'.format(fields,self.tblname,accounts,ticker_item) 
            df=pd.read_sql_query(query, self.con)
            
            df=df.transpose()
            
            df['Ticker']=ticker
            df=df[2:]

            if new_df.shape[0]==0:
                new_df = df
            else:
                new_df=self.appendDF(new_df,df)

        first_col=new_df.pop('Ticker')
        new_df.insert(0,'Ticker',first_col)
        new_df=self.addFirstCol(new_df)
        new_df.columns = ['Periode','Ticker'] + self.accounts
        #new_df["Year"]=new_df["Periode"]
        new_df["Periode"] = pd.to_datetime(new_df["Periode"],format='%Y')
        #ganti angka 0 dengan Nan supaya grafik nya tidak nyungsep
        new_df.replace(0, np.nan, inplace=True)
        
        #change string to float
        new_df=self.convert_string_to_decimal(new_df)
        
        return new_df

    def plot(self,df,account):
        import seaborn as sns 
        import matplotlib.pyplot as plt 
        for ticker in self.ticker:
            plot_df=df.loc[(df.Ticker == ticker)]
            sns.lineplot(data=plot_df[account],label=ticker)
        plt.show()

    def plotly(self,df):
        import plotly.express as px
        fig = px.line(df,x='Year',y=df.columns,title=self.description)
        fig.show()

def test():
    df = pd.DataFrame({'year': [1, 2, 3, 4, 5, 6, 7, 8],
                   'A': [10, 12, 14, 15, 15, 14, 13, 18],
                   'B': [18, 18, 19, 14, 14, 11, 20, 28],
                   'C': [5, 7, 7, 9, 12, 9, 9, 4],
                   'D': [11, 8, 10, 6, 6, 5, 9, 12]})

    import seaborn as sns
    import matplotlib.pyplot as plt
    #plot sales of each store as a line
    print(df)
    sns.lineplot(data=df[['A', 'B', 'C', 'D']])
    plt.show()
if __name__=='__main__':
    dbname = '../datas/Balance_Sheet.db'
    tablename = 'Balance_Sheet'

    bs = BalanceSheet(dbname,tablename)
    if(bs.connectDB()==True):
        #untuk mendapatkan akun dari balance sheet
        df=bs.getAccount()    
        #print(df)
        df = bs.getTicker()

        bs.setTickerName(df['ticker_code'])
        bs.setAccounts(['Total Assets','+ Total Investments','Total Liabilities','Total Equity'])
        df=bs.getBalanceSheet(2015,2022)
        bs.plot(df,'Total Assets')
        bs.closeDB()
