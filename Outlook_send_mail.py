import smtplib
from win32com.client import Dispatch
from email.message import EmailMessage


#メールのパラメータ入力
#Enter SMTP Paramater 
smtp = smtplib.SMTP('smtp.office365.com',587)
user = 'x.xxxxxx@xxxx.com'
password = "xxxxxx"
message['From'] = 'x.xxxxxx@xxxxx.com'
message['To'] = 'x.xxxxxx@xxxxx.com'
message['Subject'] = xxxxx


#SMTP処理
#SMTP process
smtp.ehlo()
smtp.starttls()
smtp.ehlo()
smtp.login(user,password)
smtp.send_message(message)
smtp.quit()