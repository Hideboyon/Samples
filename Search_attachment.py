from win32com.client import Dispatch
import datetime as date

# 添付フォイルを保存するパス。
save_path = r'C:\Users\xxxxxxx\Desktop'

# ファイルが添付されているメールタイトル
sub_ = 'Daily Data Report'

# 添付ファイルの名前
att_ = 'data.xlsx'

# 当日のメール参照用変数
lag = 0

# 自分のメールボックスへアクセスする
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# 見に行くフォルダーを指定。"Data_Folder"
inbox = outlook.GetDefaultfolder(6).Folders.Item("Data_Folder")
all_inbox = inbox.Items

# メールの最新日時を取得
Val_date = (date.date.today() - date.timedelta(lag)).strftime("%d/%m/%y")
print(Val_date)

# 条件文で該当のメールを探す処理
for msg in all_inbox:
    if sub_ in msg.Subject:
       for att in msg.Attachments:
            if att_ in att.FileName:
                # ファイル名を追加して最初に指定した保存用パス、それに任意のファイル名前を追加
                att.SaveAsFile(save_path + '\\Today_data.xlsx') 