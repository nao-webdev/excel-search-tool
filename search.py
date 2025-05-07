import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter.simpledialog import askstring
from openpyxl.utils import get_column_letter

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
file_list =[file_name for file_name in file_list if "xlsx" in file_name]

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

                # 行番号、列番号、セルの値を変数に格納
                index = column_data.index[0] + 1
                column = get_column_letter(number + 1)
                cell_value = df_sheet.at[column_data.index[0], number]

                # 結果を出力
                print('【ファイル名】' + file)
                print('【シート名　】' + key)
                print('【セルの場所】' + str(column) + str(index))
                print('【セルの値　】' + cell_value)
                print('--------------------------------------------------')