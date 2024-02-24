import re
import win32com.client
import to_excel
import get_prise_string

def main():
    FOLDER = "coin"
    SENDER_ADDR = "hirotomoku@gmail.com"


    #文字列から特定の文字列を検索し、その後の文字列を抽出する関数

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 受信トレイフォルダ

    # "FOLDER"フォルダを探す
    test_folder = None
    for folder in inbox.Folders:
        if folder.Name == FOLDER:
            test_folder = folder
            break

    if test_folder:
        mails = test_folder.Items
        for mail in mails:
            #送信者取得
            sender = mail.SenderEmailAddress
            #特定の送信者じゃなければスルー
            if sender != SENDER_ADDR:
                continue
            #件名取得
            subject = mail.Subject
            #受信時間取得
            received_time = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
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
    pass

if __name__ == '__main__':
    main()