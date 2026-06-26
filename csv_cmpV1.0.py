import os
import pandas as pd
import datetime
import chardet  # 追加: エンコーディング判定用ライブラリ

def detect_encoding(file_path):
    """ファイルの文字コードを自動判定する関数"""
    with open(file_path, 'rb') as f:
        # パフォーマンス向上のため、最初の10万バイトだけ読み込んで判定
        raw_data = f.read(100000)
    result = chardet.detect(raw_data)
    # 判定結果の文字コードを返す（例: 'utf-8', 'shift_jis', 'EUC-JP'など）
    return result['encoding']

# 指定したフォルダのパス
folder_path = './'

# フォルダ内のCSVファイルを取得
csv_files = [file for file in os.listdir(folder_path) if file.endswith(".csv")]

# ファイルを更新日時でソート
csv_files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)

# newest_file = csv_files[0]

# 更新日時が最も新しい2つのCSVファイルを選択
if len(csv_files) < 2:
    print("フォルダ内に2つ以上のCSVファイルが必要です。")
else:
    newest_file_path = os.path.join(folder_path, csv_files[0])
    second_newest_file_path = os.path.join(folder_path, csv_files[1])

    # それぞれのファイルのエンコーディングを判定
    encoding_newest = detect_encoding(newest_file_path)
    encoding_second = detect_encoding(second_newest_file_path)

    # 判定したエンコーディングを指定してCSVファイルを読み込み
    df_newest = pd.read_csv(newest_file_path, encoding=encoding_newest)
    df_second_newest = pd.read_csv(second_newest_file_path, encoding=encoding_second)

    # 差分を抽出
    diff = pd.concat([df_newest, df_second_newest]).drop_duplicates(keep=False)

    # 差分CSVファイルを指定したパスに保存
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")  # 現在の日時を取得
    
    # OSの違いによるパスエラーを防ぐため、os.path.joinに変更
    diff_path = os.path.join(folder_path, f"差分ファイル_{timestamp}.csv")
    
    if not diff.empty:
        # 出力時は元の指定通りShift-JISで保存（変換できない文字エラーを防ぐため errors='replace' を追加）
        diff.to_csv(diff_path, index=False, encoding='shift-jis', errors='replace')
        print(f"差分が見つかりました。差分ファイルを {diff_path} に保存しました。")
    else:
        print("差分は見つかりませんでした。")