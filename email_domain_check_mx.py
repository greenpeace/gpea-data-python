#Python 3.6.4
#190723 created
#220310 translated

#import modules
import pymssql
import pandas as pd
import numpy as np
import xlsxwriter
import datetime
import pymailcheck

#get today's date
now = datetime.datetime.now()
today = str(now)[:10]
thismonth = str(now)[:7]

#pymailcheck
#put domains you want to keep in custom_domains
custom_domains = ["naver.com","hanmail.net","gmail.com.hk","gmail.com","daum.net", "ymail.com", "hotmail.com.tw", "hotmail.com", "yahoo.com", "yahoo.com.tw", "hotmail.com.hk", "yahoo.com.hk"]

                  #"nate.com","hotmail.com","korea.kr","empal.com","icloud.com","sen.go.kr","yahoo.com","dreamwiz.com","empas.com","korea.com","lycos.co.kr","samsung.com","snu.ac.kr","yahoo.co.kr","me.com","paran.com","chol.com","korea.ac.kr","seoul.go.kr","sk.com","msn.com","outlook.com","kaist.ac.kr","cj.net","kt.com"]

def check(row):
    a = pymailcheck.suggest(row['Email'], domains=custom_domains, top_level_domains=[], second_level_domains=[])
    if a:
        return a['full']
    else:
        return ''

#import excel file
#modify infile location
infile = 'C:\\Users\\Dojin Kim\\Downloads\\HK TW KR Invalid Email (Typo)-2021-11-16-11-49-53.xlsx'
df = pd.read_excel(infile, header = 0)


#apply function
df['checked'] = df.apply(check, axis=1)

df = df[df.checked != '']

#export excel file
outfile = 'D:\\email_target\\emailcheck_'+today+'.xlsx'
writer = pd.ExcelWriter(outfile)
df.to_excel(writer, 'Sheet1', index=False)
writer.save()
writer.close()
