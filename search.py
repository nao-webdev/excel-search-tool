import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter.simpledialog import askstring
from openpyxl.utils import get_column_letter
import logging
from datetime import datetime

# ログファイル名を設定する(起動日時)
now = datetime.now().strftime('%Y%m%d_%H%M%S')
log_filename = os.path.join('logs',f'result-{now}.txt')

# ロガーを作成する
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# フォーマットを作成する
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# コンソール出力用ハンドラ
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# ファイル出力用ハンドラ
file_handler = logging.FileHandler(log_filename)
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

# ハンドラをロガーに追加
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# メインウィンドウを非表示で初期化する
root = tk.Tk()
root.withdraw()

# フォルダ選択ダイアログを表示する
folder_path = filedialog.askdirectory(title="検索したいフォルダを選んでください")

# キーワード入力ダイアログを表示する
keyword = askstring("キーワード入力", "検索したい文字列を入力してください")

# フォルダかキーワードの入力が空の場合、ツールを終了する。
if not(folder_path and keyword):
    print('検索対象のフォルダとキーワードは必ず入力してください。ツールを終了します。')
    sys.exit()

# 指定したフォルダ内のすべてのファイル名をリストに入れる
file_list = os.listdir(folder_path)

# リストにexcelファイルだけを残す
file_list =[file_name for file_name in file_list if file_name.endswith(".xlsx")]

# ヒット数カウント用変数
count = 0

# excelファイル一つずつ文字検索をかける
for file in file_list:

    # ファイルパス生成
    file_path = os.path.join(folder_path, file)

    # 辞書型に読み込む
    sheets_dict = pd.read_excel(file_path, sheet_name=None, header=None, na_filter = False, engine='openpyxl')

    # シート一つずつ文字検索をかける
    for key in sheets_dict.keys():
        df_sheet = sheets_dict[key]

        # カラムごとに検索をかける
        num_columns = len(df_sheet.columns)

        # (カラム数分for文)
        for number in range(num_columns):

            # contains('')のところに検索したい文字列をいれる
            column_data = df_sheet[df_sheet[number].str.contains(keyword,na = False)]

            # 当てはまっていれば(indexが0でなければ)コンソール出力
            if(not len(column_data.index) == 0):

                for row_index in column_data.index:

                    # 行番号、列番号、セルの値を変数に格納
                    index = row_index + 1
                    column = get_column_letter(number + 1)
                    cell_value = df_sheet.at[row_index, number]

                    # 結果を出力
                    logger.info('【ファイル名】' + file)
                    logger.info('【シート名　】' + key)
                    logger.info('【セルの場所】' + str(column) + str(index))
                    logger.info('【セルの値　】' + cell_value)
                    logger.info('--------------------------------------------------')
                    count += 1

# 最後の出力行に全ヒット数を表示する                
logger.info('--------------------------------------------------')        
logger.info('【全ヒット数】' + str(count) + '件')
logger.info('--------------------------------------------------') 