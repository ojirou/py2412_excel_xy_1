import pandas as pd
import subprocess

def process_excel(input_file, output_file):
    # Excelファイルを読み込む
    data = pd.read_excel(input_file)
    
    # データを加工（縦に並べるために1列にまとめる）
    new_row = []
    for i in range(len(data)):
        row = list(data.iloc[i])  # 各行をリスト化
        new_row.extend(row)      # 1つのリストに結合

    # 縦に並べる（1列のデータフレームに変換）
    processed_data = pd.DataFrame(new_row, columns=["Values"])

    # Excelとして保存（インデックスなし）
    processed_data.to_excel(output_file, index=False)

# 入力ファイルと出力ファイルのパス
input_file = r"sample_xy_1\\sample.xlsx"
output_file = r"sample_xy_1\\sample_output.xlsx"

# 処理を実行
process_excel(input_file, output_file)

# 出力ファイルを開く
subprocess.Popen(["start", "", output_file], shell=True)