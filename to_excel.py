import openpyxl
ROW = 1
MAX_ROW = 100000

class OpeExcel:
    def to_excel_from_mail(self,objects):
        # Excelファイルを開く
        MOST_RECENT_NUM = 1
        workbook = openpyxl.load_workbook(r"C:\Users\hirot\Desktop\python\python_base\test0215.xlsx")

        # シートを選択する
        sheet = workbook.active
        #特定のシートを選択
        #sheet = workbook["PRISE"]

        for row in range(2, MAX_ROW):  # max_rowはワークシートの最大行数
        # B列のセルを取得
            cell_value = sheet[f'A{row}'].value
        # セルが空ならその行番号を出力
            if not cell_value:  # セルが空（None、''、0、False）の場合
                print(f"最初の空白行は {row} 行目です。")
                ROW = row
                break
            num = int(sheet[f'J{row}'].value)
            if MOST_RECENT_NUM < num:
                MOST_RECENT_NUM = num
                
        print(MOST_RECENT_NUM)
        if int(objects[0]) > MOST_RECENT_NUM:
            sheet[f'A{ROW}'] = "Praise-net"
            sheet[f'B{ROW}'] =objects[1]
            sheet[f'C{ROW}'] =objects[2]
            sheet[f'D{ROW}'] =objects[3]
            sheet[f'E{ROW}'] =objects[4]
            sheet[f'J{ROW}'] =objects[0]
            sheet[f'K{ROW}'] =objects[5]

        # Excelファイルを保存する
        workbook.save(r'C:\Users\hirot\Desktop\python\python_base\test0215.xlsx')

