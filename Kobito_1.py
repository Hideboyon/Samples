# import and define modules.
# モジュールのインポート hoge hoge hoge
# hoge2 hoge2 hoge2
import datetime as date
from win32com.client import Dispatch
import pandas as pd
import smtplib
from email.message import EmailMessage

# Define an attached file path. 
# 添付フォイルを保存するパス。
save_path = r'C:\Users\xxxxxxx\Desktop'

# Existing mail of the subuject which has an attached file.
# ファイルが添付されているメールタイトル
sub_ = 'Daily Data Report'

# Name of the attahced file.
# 添付ファイルの名前
att_ = 'data.xlsx'

# Define a variable for today's email.
# 当日のメール参照用変数
lag = 0

# Access to mailbox of outlook.
# 自分のメールボックスへアクセスする
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# Define cheking folder. In this case which name is "Data_Folder". 
# 見に行くフォルダーを指定。"Data_Folder"
inbox = outlook.GetDefaultfolder(6).Folders.Item("Data_Folder")
all_inbox = inbox.Items

# Get a date of the latest email.
# メールの最新日時を取得
Val_date = (date.date.today() - date.timedelta(lag)).strftime("%d/%m/%y")
print(Val_date)

# Search email of the target.
# 条件文で該当のメールを探す処理
for msg in all_inbox:
    if sub_ in msg.Subject:
       for att in msg.Attachments:
            if att_ in att.FileName:
                # Add file name then for the PATH of the save file.
                # ファイル名を追加して最初に指定した保存用パス、それに任意のファイル名前を追加
                att.SaveAsFile(save_path + '\\Today_data.xlsx') 

# Define the PATH for contents of sending email.
# メール送信用のテキストファイルのパス指定
path = 'C:/Users/xxxxxxx/Desktop/report.txt'

# Processing data frame.
# dataフレームの抽出・加工処理
df = pd.read_excel('C:/Users/xxxxxxx/Desktop/data.xlsx')
df.columns= [str(s).replace(' ','_') for s in df.columns] #query メソッドでindex（列名）読み込めるように列の文字列にある空白は＿に変換
df['Group'].fillna('NA', inplace=True) #query　メソッドで読み込めるようにNaNがある場合は、NA文字に置き換える処理にした
df = df.query('Group.str.contains("JP")', engine = 'python')#Groupの JP という文字列を含むものを抽出
df = (df[['Score','Group','Number','Type','Car']])#表示する列を抽出
df = df.sort_values('Type',ascending=False)#該当の列（ヘッダー）を昇順・降順にするかの処理
df2 = df[df['Score']<=3.0]#スコアが、3.0以下のデータを抽出
df3 = df[df['Score'].isnull()]#スコアに欠損値（空データ）がある場合、該当する行を抽出


# Processing for a related date.
# 日付関連の処理
dt = date.date.today()
dts = dt.strftime('[%Y-%m-%d]')

# Create mail contents.
# メールに記載する文字列の処理
s = 'みなさん、お疲れ様です。\
\n毎朝10時10分に自動でファイルを取得\
\nデータを整形し配信しております。\
\n\nえー本日も平和なり、本日も平和なり・・・というのはうそで、\
\nしっかりとスコア確認して対応しましょう\n'

s2 = '\n::::::::::: 確認必要事項:::::::::::\n\n' 

s3 = '\n::::::::::: データサマリ :::::::::::\n'   

s4 = '\n\n完全自動化されました。ただしぼくのPCが起動中のみ（汗） \n'

h = '\nScore	Group	Number    Type 	Car\n'

# Create mail subject content.
# メールの表題に記載する文字列の処理.
sub = 'daily check'
subs = dts+sub

#メール本文に記載していく処理
with open(path,mode='w',encoding="utf-8") as fp:
    fp.write(s)
    fp.write(s2)
    fp.write(h)
    df2.to_csv(fp,sep="\t",header=False,index=False)#Scor特定の数値以下の場合のデータ出力
    df3.to_csv(fp,sep="\t",header=False,index=False)#Scoreがnaの場合のデータ出力
    fp.write(s3)
    fp.write(h)
    df.to_csv(fp,sep="\t",header=False,index=False) # pandasを使ってデータシートを出力
   # fp.write(member)
    fp.write(s4)
   # fp.write(member2)

with open(path,mode='r',encoding="utf-8") as fp: #pathをreadモードで読み込み
    message = EmailMessage()
    message.set_content(fp.read())


#メール処理関連
smtp = smtplib.SMTP('smtp.office365.com',587)
user = 'xxxxxxx@xxxxxx.com'
password = "xxxxxxxxxxxxxx"
message['From'] = 'xxxxxxx@xxxxxx.com'
message['To'] = 'xxxxxxxxxxx@xxxxxx.com'
message['Subject'] = subs


#SMTP処理
smtp.ehlo()
smtp.starttls()
smtp.ehlo()
smtp.login(user,password)
smtp.send_message(message)
smtp.quit()


