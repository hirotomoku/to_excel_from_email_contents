import re

class PriseString:

    # メールからn)ごとの文字列を取得する関数
    def get_mail_stiring_blocks(self,mail):
        blocks = []
        # 数字)改行で文字列を分割
        blocks = re.split('\d+\)\r\n', mail)
        # 分割後最初の要素は不要なので削除
        blocks.pop(0)
        return blocks
    

    # n)ごとのブロックからさらに必要な表題など必要な情報を抽出する関数
    def get_usage_words_from_block(self,body):
        object_list = []
        # 各情報を抽出する正規表現パターン
        pattern_number = r'管理NO\s*:(.*?)(?:\r\n|\r|\n|$)'
        pattern_date = r'発信日付\s*:(.*?)(?:\r\n|\r|\n|$)'
        pattern_title = r'表題\s*:(.*?)(?:\r\n|\r|\n|$)'
        pattern_out_number = r'発信番号\s*:(.*?)(?:\r\n|\r|\n|$)'
        pattern_name = r'発信者名\s*:(.*?)(?:\r\n|\r|\n|$)'
        pattern_file = r'原文ファイル\s*:\s*(?:\r\n|\r|\n)(.*?)(?:\r\n|\r|\n|$)'
       

        # 正規表現にマッチする部分を検索
        match1 = re.search(pattern_number, body)
        match2 = re.search(pattern_date, body)
        match3 = re.search(pattern_title, body)
        match4 = re.search(pattern_out_number, body)
        match5 = re.search(pattern_name, body)
        match6 = re.search(pattern_file, body)

        #パターンにマッチしたらlistに格納
        if match1:
            num = match1.group(1)
        else:
            num = ""
        object_list.append(num)

        if match2:
            date = match2.group(1)
        else:
            date = ""
        object_list.append(date)

        if match3:
            title = match3.group(1)
        else:
            title = ""
        object_list.append(title)

        if match4:
            out_number = match4.group(1)
        else:
            out_number = ""
        object_list.append(out_number)

        if match5:
            name = match5.group(1)
        else:
            name = ""
        object_list.append(name)

        if match6:
            file = match6.group(1)
        else:
            file = ""
        object_list.append(file)

        return object_list
        
    
