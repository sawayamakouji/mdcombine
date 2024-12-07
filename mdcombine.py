import pandas as pd
import glob
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

# エクスプローラーを開いてフォルダを選択する
def select_folder():
    root = Tk()
    root.withdraw()  # Tkinterウィンドウを非表示にする
    folder_path = askdirectory(title="Excelファイルが格納されているフォルダを選択してください")
    if not folder_path:
        print("フォルダが選択されませんでした。")
        exit()  # プログラム終了
    return folder_path

# フォルダ選択を実行
folder_path = select_folder()

# フォルダ内のすべてのExcelファイル（.xlsx）をリストに追加
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# 必要な項目（列）を抽出する関数
def load_filtered_data(file_path):
    # "新データ入力"シートからデータを読み込む
    df = pd.read_excel(file_path, sheet_name='新データ入力', header=None)

    # 2列目（index 1）にNaN値がある行をすべて削除
    df_cleaned = df.dropna(subset=[1])

    # 最初の行を新しいヘッダーとして設定
    df_cleaned.columns = df_cleaned.iloc[0]
    df_cleaned = df_cleaned[1:]  # 新しいヘッダー行をデータから削除

    # インデックスをリセット
    df_cleaned.reset_index(drop=True, inplace=True)

    # 抽出したい列名のリスト
    columns_to_extract = ['伝票日付', 'チラシ日', '売価開始日', '売価終了日', '重要度', '部門', 'JAN', 
                          'メーカー名', '商品名', '規格', '発注単位', '原価', '本体売価', 
                          '総額売価', '税率', '値入率', '帳合先選択', '帳合先', '帳合先枝番', '備考', 'ＭＶ稲田','ＭＶ池田','ＭＶ中札内','MV音更',]
    
    # 必要な列をフィルタリング（存在しない列がある場合は無視する）
    filtered_df = df_cleaned.loc[:, df_cleaned.columns.isin(columns_to_extract)].copy()

    # ファイル名の列を追加
    filtered_df.loc[:, 'ファイル名'] = os.path.basename(file_path)

    return filtered_df

# 全てのファイルからデータを抽出し、リストに保存
all_data = []
for file_path in excel_files:
    try:
        data = load_filtered_data(file_path)
        all_data.append(data)
    except Exception as e:
        print(f"エラーが発生しました: {file_path} - {e}")

# 各ファイルのデータを結合
combined_data = pd.concat(all_data, ignore_index=True)

# 列の順序を調整して「ファイル名」を左端に配置
cols = ['ファイル名'] + [col for col in combined_data.columns if col != 'ファイル名']
combined_data = combined_data[cols]

# CSVファイルに出力
output_path = os.path.join(folder_path, 'combined_data_output.csv')  # 入力フォルダ内に出力
combined_data.to_csv(output_path, index=False, encoding='utf-8-sig')

# 出力ファイルのパスを表示
print(f"結合したデータは以下のパスに保存されました: {output_path}")
