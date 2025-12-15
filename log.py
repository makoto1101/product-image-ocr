import datetime
import gspread
from google.oauth2 import service_account
import streamlit as st

def log_ocr_execution(creds_info, spreadsheet_id, user_info, image_count, input_tokens, output_tokens):
    """
    OCR実行ログをスプレッドシートの「logs」シートに記録する関数
    入力/出力トークンを分けて記録し、概算コストも計算する
    ※日時は日本時間(JST)で記録する
    """
    try:
        # --- 認証とスプレッドシートへの接続 ---
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=scopes
        )
        gc = gspread.authorize(creds)

        # スプレッドシートを開く
        try:
            sh = gc.open_by_key(spreadsheet_id)
        except Exception as e:
            print(f"ログ記録エラー: スプレッドシートが見つかりません。 {e}")
            return

        # 「logs」シートの取得、なければ作成
        SHEET_NAME = 'logs'
        try:
            worksheet = sh.worksheet(SHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=SHEET_NAME, rows=100, cols=8)

        # --- ヘッダーの確認と追加 ---
        if not worksheet.get_all_values():
            # ヘッダーを拡張
            worksheet.append_row([
                "日時", 
                "利用者", 
                "画像枚数", 
                "入力トークン", 
                "出力トークン", 
                "合計トークン", 
                "概算コスト(円)"
            ])
            
        # --- 記録するデータの準備 ---
        
        # 日本時間 (JST: UTC+9) を生成
        t_delta = datetime.timedelta(hours=9)
        JST = datetime.timezone(t_delta, 'JST')
        now = datetime.datetime.now(JST)
        now_str = now.strftime('%Y/%m/%d %H:%M:%S')
        
        user_display = str(user_info)
        total_tokens = input_tokens + output_tokens

        # --- 概算コスト計算 (1ドル=150円換算) ---
        # Input: $2.50 / 1M tokens -> 375円
        # Output: $10.00 / 1M tokens -> 1500円
        cost_input = (input_tokens / 1_000_000) * 2.50 * 150
        cost_output = (output_tokens / 1_000_000) * 10.00 * 150
        total_cost = round(cost_input + cost_output, 2) # 小数点第2位まで

        # 行データ
        row_data = [
            now_str,        # A: 日時 (JST)
            user_display,   # B: 利用者
            image_count,    # C: 画像枚数
            input_tokens,   # D: 入力
            output_tokens,  # E: 出力
            total_tokens,   # F: 合計
            total_cost      # G: 概算コスト
        ]

        # --- 挿入実行 (2行目) ---
        worksheet.insert_row(row_data, index=2)
        
    except Exception as e:
        print(f"ログ記録中にエラーが発生しました: {e}")