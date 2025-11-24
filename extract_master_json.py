import csv
import json
import os
import datetime
import config

def read_csv_with_encoding(filename):
    """utf-8-sig または cp932 でCSV読み込みを試行する"""
    for enc in config.ENCODINGS:
        try:
            with open(filename, 'r', encoding=enc, newline='') as f:
                reader = csv.reader(f)
                data = list(reader)
            return data
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"エラー: {enc} での読み込みに失敗しました: {e}")
            raise
    raise ValueError(f"サポートされているエンコーディングで {filename} をデコードできませんでした。")

def extract_resume_data():
    # 出力ディレクトリの作成
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    # CSV読み込み
    try:
        rows = read_csv_with_encoding(config.INPUT_FILE)
    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{config.INPUT_FILE}' が見つかりません。")
        return

    # 2. 読み取りロジック
    start_index = config.START_INDEX
    
    # フッターの位置（"その他"）を先に検索する
    footer_marker_index = len(rows)
    for i in range(start_index, len(rows)):
        if len(rows[i]) > 0 and rows[i][0].strip() == "その他":
            footer_marker_index = i
            break

    work_history = []
    current_index = start_index
    
    # 職務経歴のループ
    while current_index < len(rows):
        # フッター開始行に到達したら終了
        if current_index >= footer_marker_index:
            break
            
        # 行が存在するか確認
        if current_index >= len(rows):
            break
            
        row = rows[current_index]
        
        # A列の値を確認
        col_a_val = row[0].strip() if len(row) > 0 else ""
        
        # A列が空の場合はスキップ（空行対応）
        if not col_a_val:
            current_index += 1
            continue
            
        # 5行分のブロックを取得できるか確認
        # ※最後のブロックがフッター行（"その他"）にかかる場合でも、5行固定ルールに従い取得する
        if current_index + config.BLOCK_SIZE > len(rows):
            break
            
        block_rows = rows[current_index : current_index + config.BLOCK_SIZE]
        
        # 3. カラムマッピング
        # 値を安全に取得するヘルパー関数
        def get_val(row, col_idx):
            if col_idx < len(row):
                val = row[col_idx]
                return val if val is not None else ""
            return ""

        # 5行にわたってカラムから配列を取得するヘルパー関数
        def get_col_array(rows_subset, col_idx):
            return [get_val(r, col_idx) for r in rows_subset]

        entry = {
            "no": get_val(block_rows[0], 0),
            "period": {
                "start": get_val(block_rows[0], 1),
                "end": get_val(block_rows[2], 1)
            },
            "business_content": {
                "title_col_e": get_col_array(block_rows, 4),
                "role_col_f": get_col_array(block_rows, 5),
                "detail_col_g": get_col_array(block_rows, 6)
            },
            "technology": {
                "environment_col_u": get_col_array(block_rows, 20),
                "language_col_z": get_col_array(block_rows, 25),
                "process_col_ae": get_col_array(block_rows, 30)
            }
        }
        
        work_history.append(entry)
        
        # 次のブロックへ（5行進める）
        current_index += config.BLOCK_SIZE

    # 4. フッター抽出
    footer_data = {"other_col_b": []}
    
    # フッターマーカーの次の行から5行を取得
    footer_start_data_index = footer_marker_index + 1
    
    if footer_start_data_index < len(rows):
        # フッターデータとして5行取得（ファイル末尾までが5行未満の場合はあるだけ取得）
        end_idx = min(footer_start_data_index + config.BLOCK_SIZE, len(rows))
        footer_rows = rows[footer_start_data_index : end_idx]
        
        # 5行に満たない場合、空文字で埋める必要がある場合はここで調整
        # 要件は「5行分を配列化」なので、足りない場合は空文字を追加する
        extracted_footer = [get_val(r, 1) for r in footer_rows]
        while len(extracted_footer) < config.BLOCK_SIZE:
            extracted_footer.append("")
            
        footer_data["other_col_b"] = extracted_footer

    # 5. JSON構築
    output_data = {
        "meta": {
            "source": config.INPUT_FILE,
            "extracted_at": datetime.date.today().isoformat()
        },
        "work_history": work_history,
        "footer": footer_data
    }

    # JSON書き出し
    with open(config.OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    
    print(f"成功: {config.OUTPUT_FILE} を生成しました。")

if __name__ == "__main__":
    extract_resume_data()
