import re
import win32com.client
import to_excel
import get_prise_string
from datetime import datetime
import os
import subprocess

def main():
    FOLDER = "coin"
    SENDER_ADDR = "hirotomoku@gmail.com"
    ENV_LAST_RECEIVED_DATE = os.getenv('Last_Received_Date')

    print("環境変数'A'の値:", ENV_LAST_RECEIVED_DATE)


    #文字列から特定の文字列を検索し、その後の文字列を抽出する関数

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 受信トレイフォルダ
    received_time = None

    #環境変数ENV_LAST_RECEIVED_DATEをdatetime型に変換
    last_date = datetime.strptime(ENV_LAST_RECEIVED_DATE, "%Y-%m-%d %H:%M:%S")

    # "FOLDER"フォルダを探す
    test_folder = None
    for folder in inbox.Folders:
        if folder.Name == FOLDER:
            test_folder = folder
            break

    if test_folder:
        mails = test_folder.Items
        for mail in mails:
            print("oooooooooooooooooo")
            #受信時間取得
            received_time = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            received_time_date = datetime.strptime(received_time, "%Y-%m-%d %H:%M:%S")
            #受信日が環境変数の日付より古かったらスルー
            if last_date < received_time_date:
                print(received_time)
                #メール本文から複数block( 1)2)...単位)を取得
                priseString = get_prise_string.PriseString()
                blocks = priseString.get_mail_stiring_blocks(mail.Body)
                #ブロックごとに必要なオブジェクトを取り出す
                for block in blocks:
                    object_list = priseString.get_usage_words_from_block(block)
                    if object_list:
                        print(f'{object_list}')
                        opeExcel = to_excel.OpeExcel()
                        opeExcel.to_excel_from_mail(object_list)
                        del opeExcel
    else:
        print("Folder 'Test' not found.")
    #最新の受信日を環境変数にセット
    command = f'setx Last_Received_Date "{received_time}"'
    subprocess.run(command, shell=True)
    pass

if __name__ == '__main__':
    main()