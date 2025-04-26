import pandas as pd
import os

# 指定したフォルダ内のすべてのファイル名をリストに入れる
folder_path = '/検索対象フォルダ'
file_list = os.listdir(folder_path)

# リストにexcelファイルだけを残す
file_list =[file_name for file_name in file_list if "xlsx" in file_name]

# excelファイル一つずつ文字検索をかける
for file in file_list:
    # ファイルパス生成
    file_path = folder_path +"/" + file
    # dictに変換
    dict = pd.read_excel(file_path, sheet_name=None, header=None, na_filter = False, engine='openpyxl')

    # シート一つずつ文字検索をかける
    for key in dict.keys():
        df_sheet = dict[key]

        # カラムごとに検索をかける
        num_columns = len(df_sheet.columns)

        # (カラム数分for文)
        for number in range(num_columns):
            # contains('')のところに検索したい文字列をいれる
            column_data = df_sheet[df_sheet[number].str.contains('検索したい文字')]
            # 当てはまっていれば(indexが0でなければ)行出力
            if(not len(column_data.index) == 0):
                print('---------------' + file +'の'+ key + '----------------')
                print(column_data)

            
        
        

        



    