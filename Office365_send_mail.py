import smtplib
from email.message import EmailMessage

#SMTP接続パラメータ入力
#Enter SMTP Paramater 
smtp = smtplib.SMTP('smtp.office365.com',587)
user = 'x.xxxxxx@xxxx.com'
password = "xxxxxx"

# メール 送信元・宛先・タイトル作成
# Define From / To / Subject with messag function
message = EmailMessage()
message['From'] = 'x.xxxxxx@xxxxx.com'
message['To'] = 'x.xxxxxx@xxxxx.com'
message['Subject'] = 'xxxxx'

#SMTP処理
#Processing SMTP
smtp.ehlo()
smtp.starttls()
smtp.ehlo()
smtp.login(user,password)
smtp.send_message(message)
smtp.quit()