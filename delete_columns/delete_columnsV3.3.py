import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import os
import glob
import logging
from datetime import datetime

def excel_col_to_number(col_letter):
    """Excel列のアルファベットを列番号に変換（A=1, B=2, ..., Z=26, AA=27, ...）"""
    col = 0
    for char in col_letter:
        col = col * 26 + (ord(char) - ord('A') + 1)
    return col

def get_delete_ranges(filename):
    """
    ファイル名からキーワードを検出し、対応する削除列範囲を返す
    
    Args:
        filename: ファイル名
    
    Returns:
        削除する列の範囲リスト [('A', 'BV'), ('EY', 'GB'), ...]
    """
    delete_config = {
        'ASSY 試作用': [('A', 'BV'), ('EX', 'GA')],
        'ASSY': [('A', 'BV'), ('EY', 'GB')],
        'CNC1': [('A', 'BV'), ('EH', 'EN'), ('EP', 'ES'), ('EU', 'FX'), ('JU', 'KS'), ('TA', 'TW'), ('TY', 'TY')],
        'CNC2': [('A', 'BV'), ('EH', 'EN'), ('EP', 'ES'), ('EU', 'FX'), ('JU', 'KS'), ('TA', 'TW'), ('TY', 'TY')],
        'S1研磨': [('A', 'BV'), ('EN', 'ET'), ('EV', 'GA'), ('OR', 'PM')],
        'S1切断': [('A', 'BV'), ('EH', 'EO'), ('EQ', 'FU'), ('JR', 'KP'), ('SX', 'TT')],
        'S2研磨': [('A', 'BV'), ('EN', 'ET'), ('EV', 'GA'), ('OR', 'PM')],
        'S2切断': [('A', 'BV'), ('EH', 'EO'), ('EQ', 'FU'), ('JR', 'KP'), ('SX', 'TT')],
        'マーク': [('A', 'BV'), ('DA', 'DD')],
        'ラミ': [('A', 'BV'), ('EM', 'FQ'), ('GC', 'GC')],
        '印刷 タブレット用': [('A', 'BV'), ('LR', 'MQ')],
        '印刷': [('A', 'BV'), ('ET', 'EX')],
        '強化': [('A', 'BV'), ('FB', 'FF')],
        '研磨': [('A', 'BV'), ('EM', 'ER'), ('ET', 'ET'), ('EV', 'GA'), ('OR', 'PM')],
        '孔あけ': [('A', 'BV'), ('EH', 'FX'), ('JK', 'JK')],
        '切断＋研磨': [('A', 'BV'), ('EN', 'GD'), ('KA', 'KZ'), ('OW', 'PV'), ('YD', 'ZA')],
        '切断': [('A', 'BV'), ('EH', 'EN'), ('EP', 'ES'), ('EU', 'FX'), ('JU', 'KS'), ('TA', 'TW'), ('TY', 'TY')],
        '特殊作業': [('A', 'BV'), ('EP', 'ET')],
        '品証': [('A', 'BV'), ('EH', 'EL')],
    }
    
    # ファイル名からキーワードを検出（長いキーワードから順にチェック）
    for keyword in sorted(delete_config.keys(), key=len, reverse=True):
        if keyword in filename:
            return delete_config[keyword], keyword
    
    return None, None

def log_message(message, logger=None):
    """ロガーが指定されていればロガーに出力し、なければprintする"""
    if logger:
        logger.info(message)
    else:
        print(message)

def delete_columns_from_csv(input_file, output_file, delete_ranges, keyword='', logger=None):
    """
    CSVファイルから指定した列を削除する
    
    Args:
        input_file: 入力CSVファイルパス
        output_file: 出力CSVファイルパス
        delete_ranges: 削除する列の範囲リスト [('A', 'BV'), ('EY', 'GB'), ...]
        keyword: ファイルを識別するキーワード
        logger: ロギングオブジェクト
    """
    for enc in ['cp932','shift_jis','utf-8','utf-8-sig','latin1']:
        try:
            df = pd.read_csv(input_file, encoding=enc)
            log_message(f"loaded with {enc}", logger)
            break
        except Exception:
            continue
    else:
        raise RuntimeError('読み込みに全て失敗しました')
    try:
        log_message(f"  元のデータ形状: {df.shape}", logger)
        log_message(f"  元の列数: {len(df.columns)}", logger)
        
        # 削除する列番号を収集
        cols_to_delete = []
        
        for start_col, end_col in delete_ranges:
            start_num = excel_col_to_number(start_col)
            end_num = excel_col_to_number(end_col)
            
            log_message(f"  削除範囲: {start_col}({start_num}) ～ {end_col}({end_num})", logger)
            
            # 列インデックスは0ベースなので、-1する
            for col_num in range(start_num - 1, end_num):
                if col_num < len(df.columns):
                    cols_to_delete.append(col_num)
        
        # 重複を削除してソート
        cols_to_delete = sorted(set(cols_to_delete))
        log_message(f"  削除する列数: {len(cols_to_delete)}", logger)
        
        # 削除する列のインデックスで列を削除
        df_result = df.drop(df.columns[cols_to_delete], axis=1)
        
        log_message(f"  削除後のデータ形状: {df_result.shape}", logger)
        log_message(f"  削除後の列数: {len(df_result.columns)}", logger)
        
        # S2研磨の特殊処理：69列目と70列目を入れ替える
        if keyword == 'S2研磨':
            try:
                if len(df_result.columns) >= 70:
                    log_message(f"  特殊処理: Col 69（設備コード）とCol 70（受注番号）を入れ替えます", logger)
                    # 69列目と70列目のインデックス（0ベース）
                    col69_idx = 68
                    col70_idx = 69
                    
                    # 列を入れ替え
                    cols = list(df_result.columns)
                    cols[col69_idx], cols[col70_idx] = cols[col70_idx], cols[col69_idx]
                    df_result = df_result[cols]
                    log_message(f"  ✓ 列の入れ替えが完了しました", logger)
            except Exception as e:
                log_message(f"  ⚠ 列の入れ替え処理中にエラーが発生しました: {e}", logger)
        
        # 結果をCSVに保存（UTF-8エンコーディング）
        df_result.to_csv(output_file, index=False, encoding='utf-8-sig')
        log_message(f"  ✓ ファイルを保存しました: {output_file}", logger)
        return True
        
    except Exception as e:
        log_message(f"  ✗ エラーが発生しました: {e}", logger)
        return False

def delete_columns_from_excel(input_file, output_file, delete_ranges, keyword='', logger=None):
    """
    Excelファイルから指定した列を削除する
    
    Args:
        input_file: 入力Excelファイルパス
        output_file: 出力Excelファイルパス
        delete_ranges: 削除する列の範囲リスト [('A', 'BV'), ('EY', 'GB'), ...]
        keyword: ファイルを識別するキーワード
        logger: ロギングオブジェクト
    """
    try:
        # Excelファイルを読み込む
        wb = openpyxl.load_workbook(input_file)
        ws = wb.active
        
        log_message(f"  シート名: {ws.title}", logger)
        log_message(f"  元の最大列数: {ws.max_column}", logger)
        
        # 削除する列番号を逆順で収集（後ろから削除する）
        cols_to_delete = []
        
        for start_col, end_col in delete_ranges:
            start_num = excel_col_to_number(start_col)
            end_num = excel_col_to_number(end_col)
            
            log_message(f"  削除範囲: {start_col}({start_num}) ～ {end_col}({end_num})", logger)
            
            for col_num in range(start_num, end_num + 1):
                cols_to_delete.append(col_num)
        
        # 重複を削除して降順でソート（後ろから削除するため）
        cols_to_delete = sorted(set(cols_to_delete), reverse=True)
        log_message(f"  削除する列数: {len(cols_to_delete)}", logger)
        
        # 列を削除（後ろから削除することで列番号のずれを防ぐ）
        for col_num in cols_to_delete:
            ws.delete_cols(col_num)
        
        # S2研磨の特殊処理：69列目（BQ）と70列目（BR）を入れ替える
        if keyword == 'S2研磨':
            try:
                bq_col_num = 69  # BQ列
                br_col_num = 70  # BR列
                
                # 削除後に69列目と70列目が存在する場合のみ処理
                if bq_col_num <= ws.max_column and br_col_num <= ws.max_column:
                    log_message(f"  特殊処理: Col 69（BQ）とCol 70（BR）を入れ替えます", logger)
                    
                    # BQ列のデータを一時保存
                    bq_data = []
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=bq_col_num, max_col=bq_col_num):
                        bq_data.append([cell.value for cell in row])
                    
                    # BR列のデータを一時保存
                    br_data = []
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=br_col_num, max_col=br_col_num):
                        br_data.append([cell.value for cell in row])
                    
                    # BQ列の位置にBR列のデータを配置
                    for row_idx, row_data in enumerate(br_data, 1):
                        ws.cell(row=row_idx, column=bq_col_num).value = row_data[0]
                    
                    # BR列の位置にBQ列のデータを配置
                    for row_idx, row_data in enumerate(bq_data, 1):
                        ws.cell(row=row_idx, column=br_col_num).value = row_data[0]
                    
                    log_message(f"  ✓ 列の入れ替えが完了しました", logger)
            except Exception as e:
                log_message(f"  ⚠ 列の入れ替え処理中にエラーが発生しました: {e}", logger)
        
        log_message(f"  削除後の最大列数: {ws.max_column}", logger)
        
        # 結果を保存
        wb.save(output_file)
        log_message(f"  ✓ ファイルを保存しました: {output_file}", logger)
        return True
        
    except Exception as e:
        log_message(f"  ✗ エラーが発生しました: {e}", logger)
        return False

def process_all_files(directory='.'):
    """
    ディレクトリ内のすべてのCSV/Excelファイルを処理する
    
    Args:
        directory: 処理対象 of ディレクトリ
    """
    output_dir = os.path.join(directory, '処理ログ')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # ロギング設定
    logger = logging.getLogger('delete_columns')
    logger.setLevel(logging.INFO)
    
    # ハンドラーがすでに存在する場合はクリア（複数回実行時の重複ハンドラー防止）
    if logger.hasHandlers():
        logger.handlers.clear()
        
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(output_dir, f'処理ログ_{timestamp}.txt')
    
    # ファイルハンドラーの作成 (UTF-8)
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(message)s'))
    logger.addHandler(file_handler)
    
    # コンソールハンドラーの作成
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter('%(message)s'))
    logger.addHandler(console_handler)

    logger.info("=" * 80)
    logger.info("列削除自動処理プログラム")
    logger.info("=" * 80)
    logger.info(f"出力フォルダ: {output_dir}")
    logger.info(f"ログファイル: {log_file}\n")
    
    # CSVファイルを処理
    csv_files = glob.glob(os.path.join(directory, '*.csv'))
    excel_files = glob.glob(os.path.join(directory, '*.xlsx')) + glob.glob(os.path.join(directory, '*.xls'))
    
    all_files = csv_files + excel_files
    
    if not all_files:
        logger.info("対象ファイルが見つかりません")
        return
    
    success_count = 0
    fail_count = 0
    
    for input_file in sorted(all_files):
        filename = os.path.basename(input_file)
        delete_ranges, keyword = get_delete_ranges(filename)
        
        if delete_ranges is None:
            logger.info(f"\n⊘ スキップ: {filename}")
            logger.info("  対応するキーワードが見つかりません")
            continue
        
        # 出力ファイル名を生成（削除済フォルダへ、ファイル名は変更しない）
        output_file = os.path.join('.',filename)
        
        logger.info(f"\n▶ 処理中: {filename}")
        logger.info(f"  キーワード: {keyword}")
        
        # ファイル形式に応じて処理
        if input_file.endswith('.csv'):
            if delete_columns_from_csv(input_file, output_file, delete_ranges, keyword, logger):
                success_count += 1
            else:
                fail_count += 1
        else:
            if delete_columns_from_excel(input_file, output_file, delete_ranges, keyword, logger):
                success_count += 1
            else:
                fail_count += 1
    
    logger.info("\n" + "=" * 80)
    logger.info(f"処理完了: 成功 {success_count}件, 失敗 {fail_count}件")
    logger.info("=" * 80)

# メイン処理
if __name__ == "__main__":
    # カレントディレクトリを処理
    process_all_files()
