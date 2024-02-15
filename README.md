# library
from individual import *
import pandas as pd
dbname = 'ASURANSI.db'
tablename=['STATALE','STATRL']
if __name__=="__main__":
    df=get_data(dbname,tablename[0])
    #menampilkan distinct dari PA/S
    companies=df['NAME OF COMPANIES'].unique()
    print(companies)
    plotmultiseries(df[df['NAME OF COMPANIES'].str.contains('Wanaartha|Jiwasraya|PT Asuransi BRI Life|PT AXA Mandiri Financial Services')],'REPORT_DATE','ASSETS','NAME OF COMPANIES')
