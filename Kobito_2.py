import pandas as pd
import pyodbc
import pandas.io.sql
import datetime as date
import smtplib
from email.message import EmailMessage


#Connection script to SQL server
server = 'tcp:xxx.xxx.xxx.xxx'
database = 'xxxx'
username = 'xxxx'
password = 'xxxx'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';BASE='+database+';UID='+username+';PWD='+ password)

#Get data
df = pd.read_sql(
                #Chose colum and Select DB which is excuted some conditions to excerpt data. 
                '''SELECT TOP 150000 DATE,DDI,Number,code\
                FROM Fraud.dbo.Wrong where CallingNumber not like '+81%' '''
                #Connect SQL
                ,cnxn
                )

#Process data by pands
## Excerpt 6 digits from code
df = df.query('code.str.match("^([0-9]{6})$")', engine='python')

## Check Duplicated CLI from CallingNUmber
df = (df[df.duplicated(subset=['CallingNumber'],keep=False,)])

## Change format DATE
df['WRONG'] = df['WROHG'].dt.strftime('%Y-%m-%d %H:%M:%S')

## Sort by latest date
df = df.sort_values('WRONG', ascending=False)

## Exceprt the latest of 20 data.
df = df.head(20)

print(df[['WRONG','code','DDI','CallingNumber']])
print(df.dtypes)

#メール送信用のテキストファイルのパス指定
path = 'C:/Users/x.xxxxx/Desktop/6digits.txt'

#日付関連の処理
dt = date.date.today()
dts = dt.strftime('[%Y-%m-%d]')

#メールに記載する文字列の処理
s = 'チームのみなさん、お疲れ様です。\
\nなんとなく作ったのでテストで使わせてください\
\nデータの6桁のみ表示、最新20行だけSQLから吸い取っています。\
\n\nスクリプト起動時間は09時30分\n\n'

h1 = '=====================================================\n'
h2 = 'Date/Time        Pass   DDI            CallingNumber\n'
h3 = '=====================================================\n'

#メールの表題に記載する文字列の処理
sub = '6 digits check in JP'
subs = dts+sub

#メール本文に記載していく処理
with open(path,mode='w',encoding="utf-8") as fp:
    fp.write(s)
    fp.write(h1)
    fp.write(h2)
    fp.write(h3)
    df[['WRONG','code','DDI','CallingNumber']].to_csv(fp,sep="\t",header=False,index=False)

    
with open(path,mode='r',encoding="utf-8") as fp: #pathをreadモードで読み込み
    message = EmailMessage()
    message.set_content(fp.read())

#メール処理関連
smtp = smtplib.SMTP('xxxxxxxxx.xxxxxxxx.outlook.com',587)
user = 'xxxxxxx@xxxxxx.com'
password = "xxxxxxxxx"
message['From'] = 'xxxxxxx@xxxxxx.com'
message['To'] = 'xxxxxxx@xxxxxx.com'
message['Subject'] = subs


#SMTP処理
smtp.ehlo()
smtp.starttls()
smtp.ehlo()
smtp.login(user,password)
smtp.send_message(message)
smtp.quit()