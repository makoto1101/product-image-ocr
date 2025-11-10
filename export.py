import streamlit as st
import pandas as pd
import io
import copy
import re 
import tempfile 
import os       

# --- Google / Excel 関連 ---
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
import gspread_dataframe as gd
import pandas.io.formats.excel # ExcelWriter を import するために必要

# --- サービスアカウントのインポート ---
from google.oauth2 import service_account


# === Excel出力 (export.py) ===

def create_excel_output(df_excel, portal_files):
    """
    DataFrameとポータルファイルリストから、Excelファイルのバイナリデータを生成する。
    [方法A] =IMAGE() 関数を使用
    """
    
    # ポータル名のリストを取得（Excelの列順のため）
    all_portal_names = sorted(list(portal_files.keys())) if portal_files else []

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('OCR結果')

        # --- セルの書式設定 ---
        font_props = {'font_name': '游ゴシック'}
        border_props = {'border': 1, 'border_color': '#808080'}
        base_props = {**font_props, **border_props, 'valign': 'top', 'text_wrap': True}
        highlight_bg = {'bg_color': '#FFE5E5'} # 問題のある行の背景色
        font_color_error = {'font_color': 'red'} # "要確認"用
        font_color_neutral = {'font_color': 'gray'}
        font_color_ok = {'font_color': 'blue'} # 青文字の定義

        header_format = workbook.add_format({**base_props, 'bold': True, 'bg_color': '#E0E0E0', 'valign': 'vcenter'})
        default_format = workbook.add_format(base_props)
        url_format = workbook.add_format({**base_props, 'color': 'blue', 'underline': 1})
        
        status_ok_format = workbook.add_format({**base_props, **font_color_ok}) 
        status_error_format = workbook.add_format({**base_props, **font_color_error})
        status_neutral_format = workbook.add_format({**base_props, **font_color_neutral})
        default_highlight_format = workbook.add_format({**base_props, **highlight_bg})
        url_highlight_format = workbook.add_format({**base_props, 'color': 'blue', 'underline': 1, **highlight_bg})
        status_ok_highlight_format = workbook.add_format({**base_props, **font_color_ok, **highlight_bg}) 
        status_error_highlight_format = workbook.add_format({**base_props, **font_color_error, **highlight_bg})
        status_neutral_highlight_format = workbook.add_format({**base_props, **font_color_neutral, **highlight_bg})
        
        status_normal_ok_format = workbook.add_format({**base_props, **font_color_ok}) 
        status_normal_error_format = workbook.add_format({**base_props, **font_color_error})
        status_highlight_ok_format = workbook.add_format({**base_props, **font_color_ok, **highlight_bg}) 
        status_highlight_error_format = workbook.add_format({**base_props, **font_color_error, **highlight_bg})

        # --- 列幅設定 ---
        worksheet.set_column_pixels('A:A', 50) # No
        worksheet.set_column_pixels('B:B', 150) # 画像名
        worksheet.set_column_pixels('C:C', 100) # ステータス
        
        # --- ▼▼▼ [変更] デフォルトの行高さを設定 (112.5pt = 150px) ▼▼▼ ---
        worksheet.set_default_row(112.5)
        # --- ▲▲▲ [変更] ここまで ▲▲▲ ---
        # ヘッダー行の高さは別途設定
        worksheet.set_row(0, 30) 

        col_idx = 3 # D列から
        for _ in all_portal_names:
            worksheet.set_column_pixels(col_idx, col_idx, 200); col_idx += 1 # 画像
            worksheet.set_column_pixels(col_idx, col_idx, 300); col_idx += 1 # OCR
            worksheet.set_column_pixels(col_idx, col_idx, 150); col_idx += 1 # 内容量
        
        worksheet.set_column_pixels(col_idx, col_idx, 150) # テキスト比較
        worksheet.set_column_pixels(col_idx + 1, col_idx + 1, 200) # 誤字脱字
        worksheet.set_column_pixels(col_idx + 2, col_idx + 2, 150) # NENG内容量
        worksheet.set_column_pixels(col_idx + 3, col_idx + 3, 150) # 内容量比較
        worksheet.set_column_pixels(col_idx + 4, col_idx + 4, 150) # エラー検出


        # ヘッダー書き込み
        for col_num, value in enumerate(df_excel.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # データ書き込み (行ごとに書式を設定)
        for row_num, row_data in df_excel.iterrows():
            is_highlight_row = (row_data.get('ステータス', '') == '要確認')

            for col_num, col_name in enumerate(df_excel.columns):
                cell_value = row_data[col_name]
                if pd.isna(cell_value) or cell_value == '':
                    empty_format = default_highlight_format if is_highlight_row else default_format
                    worksheet.write(row_num + 1, col_num, '', empty_format)
                    continue

                cell_format = None
                if col_name == "ステータス":
                    if cell_value == "異常なし":
                        cell_format = status_highlight_ok_format if is_highlight_row else status_normal_ok_format
                    else: # "要確認"
                        cell_format = status_highlight_error_format if is_highlight_row else status_normal_error_format
                
                elif col_name in ["テキスト比較", "誤字脱字", "内容量比較", "エラー検出"]:
                    if cell_value == "OK！":
                        cell_format = status_ok_highlight_format if is_highlight_row else status_ok_format
                    elif cell_value in ["差分あり", "要確認"] or \
                         (col_name == "誤字脱字" and "OK！" not in str(cell_value)) or \
                         (col_name == "エラー検出" and str(cell_value) != ""): 
                        cell_format = status_error_highlight_format if is_highlight_row else status_error_format
                    elif cell_value in ["比較対象なし", "内容量記載なし"]:
                        cell_format = status_neutral_highlight_format if is_highlight_row else status_neutral_format
                    else: 
                        cell_format = status_error_highlight_format if is_highlight_row else status_error_format 
                
                elif '（画像）' in col_name:
                    cell_format = url_highlight_format if is_highlight_row else url_format
                    
                    file_id_match = re.search(r'/d/([a-zA-Z0-9_-]+)', str(cell_value))
                    
                    if file_id_match:
                        file_id = file_id_match.group(1)
                        # --- ▼▼▼ [変更] =IMAGE(URL) 形式の文字列を生成 ▼▼▼ ---
                        image_formula = f'=IMAGE("https://drive.google.com/uc?id={file_id}")'
                        # --- ▲▲▲ [変更] ここまで ▲▲▲ ---
                        worksheet.write_formula(row_num + 1, col_num, image_formula, cell_format)
                    else:
                        worksheet.write(row_num + 1, col_num, '', cell_format)
                    
                    continue
                
                else: 
                    cell_format = default_highlight_format if is_highlight_row else default_format

                worksheet.write(row_num + 1, col_num, cell_value, cell_format)

    return output.getvalue()


# === スプレッドシート出力 (export.py) ===

@st.cache_resource
def get_google_services(creds_info): 
    """サービスアカウント認証情報(辞書)からDrive, Sheets(v4), gspreadのサービスを取得"""
    if creds_info is None:
        return None, None, None

    try:
        scopes = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=scopes
        )
        
        user_drive_service = build('drive', 'v3', credentials=creds)
        user_sheets_service_v4 = build('sheets', 'v4', credentials=creds)
        gc = gspread.service_account_from_dict(creds_info)
        
        return user_drive_service, user_sheets_service_v4, gc
    except Exception as e:
        st.error(f"Googleサービス(export.py)への接続に失敗しました: {e}")
        st.warning("gspreadが 'google-auth' ライブラリと競合している可能性があります。")
        return None, None, None

# 色の定義 (Google Sheets API用)
def hex_to_rgb(hex_code):
    hex_code = hex_code.lstrip('#')
    return {
        "red": int(hex_code[0:2], 16) / 255.0,
        "green": int(hex_code[2:4], 16) / 255.0,
        "blue": int(hex_code[4:6], 16) / 255.0
    }

# --- 書式定義 ---
COLOR_RED_GS = hex_to_rgb("#FF0000")
COLOR_BLUE_GS = hex_to_rgb("#0000FF") 
COLOR_GRAY_GS = hex_to_rgb("#808080")
COLOR_HIGHLIGHT_BG_GS = hex_to_rgb("#FFE5E5")

BORDER_STYLE_GS = {"style": "SOLID", "width": 1, "color": hex_to_rgb("#808080")}
BORDERS_GS = {"top": BORDER_STYLE_GS, "bottom": BORDER_STYLE_GS, "left": BORDER_STYLE_GS, "right": BORDER_STYLE_GS}

BASE_CELL_FORMAT_GS = {
    "textFormat": {"fontFamily": "Yu Gothic"}, 
    "verticalAlignment": "TOP",
    "wrapStrategy": "WRAP",
    "borders": BORDERS_GS
}

HEADER_FORMAT_GS = {
    "backgroundColor": hex_to_rgb("#E0E0E0"),
    "textFormat": {"bold": True},
    "verticalAlignment": "MIDDLE"
}

IMAGE_CELL_FORMAT_GS = {
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "MIDDLE"
}

def get_cell_format_request(sheet_id, row_idx, col_idx, cell_format):
    """BatchUpdate用のリクエストボディを作成"""
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_idx,
                "endRowIndex": row_idx + 1,
                "startColumnIndex": col_idx,
                "endColumnIndex": col_idx + 1
            },
            "cell": {"userEnteredFormat": cell_format},
            "fields": "userEnteredFormat"
        }
    }


def format_worksheet_gspread(sheets_service, spreadsheet_id, sheet_id, df, portal_files):
    """
    Sheets API v4のBatchUpdateを使用して書式設定を行う。
    """
    
    requests = [] 
    all_portal_names = sorted(list(portal_files.keys())) if portal_files else []
    
    # --- 1. 列幅設定 ---
    col_width_requests = []
    
    col_properties = [
        {"pixelSize": 50},  # A (No)
        {"pixelSize": 150}, # B (画像名)
        {"pixelSize": 100}, # C (ステータス)
    ]
    
    col_idx = 3
    image_col_indices = [] 
    
    for _ in all_portal_names:
        col_properties.append({"pixelSize": 200}) # 画像 (幅)
        image_col_indices.append(col_idx) 
        col_idx += 1
        
        col_properties.append({"pixelSize": 300}) # OCR (広め)
        col_idx += 1
        
        col_properties.append({"pixelSize": 150}) # 内容量
        col_idx += 1

    col_properties.extend([
        {"pixelSize": 150}, # テキスト比較
        {"pixelSize": 200}, # 誤字脱字
        {"pixelSize": 150}, # NENG内容量
        {"pixelSize": 150}, # 内容量比較
        {"pixelSize": 150}, # エラー検出
    ])

    for i, props in enumerate(col_properties):
        col_width_requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": props,
                "fields": "pixelSize"
            }
        })
    
    # --- 2. 行の高さ設定 (ヘッダーのみ) ---
    col_width_requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": 0,
                "endIndex": 1
            },
            "properties": {"pixelSize": 40}, # ヘッダーの高さ
            "fields": "pixelSize"
        }
    })
    
    if len(df) > 0: 
        col_width_requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": len(df) + 1 # データ行の最後まで
                },
                "properties": {"pixelSize": 150}, # デフォルトの高さ
                "fields": "pixelSize"
            }
        })
    
    requests.extend(col_width_requests)

    # --- 3. ヘッダー書式 (A1からヘッダーの最終列まで) ---
    final_header_format = copy.deepcopy(BASE_CELL_FORMAT_GS)
    final_header_format.update(HEADER_FORMAT_GS)
    
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(df.columns)},
            "cell": {"userEnteredFormat": final_header_format},
            "fields": "userEnteredFormat"
        }
    })
    
    # --- 4. データ行のセルごとの書式設定 (Excelロジックと同様) ---
    
    # 書式定義 (gspread_formatting.CellFormat ではない、辞書形式)
    fmt_default = BASE_CELL_FORMAT_GS
    fmt_highlight_bg = {"backgroundColor": COLOR_HIGHLIGHT_BG_GS}
    fmt_text_red = {"textFormat": {"foregroundColor": COLOR_RED_GS}}
    fmt_text_blue = {"textFormat": {"foregroundColor": COLOR_BLUE_GS}}
    fmt_text_gray = {"textFormat": {"foregroundColor": COLOR_GRAY_GS}}
    
    fmt_image_cell = copy.deepcopy(BASE_CELL_FORMAT_GS)
    fmt_image_cell.update(IMAGE_CELL_FORMAT_GS)

    cell_format_requests = []

    for row_num, row_data in df.iterrows():
        row_idx_gspread = row_num + 1 # 0始まりのヘッダー行(0) + 1
        is_highlight_row = (row_data.get('ステータス', '') == '要確認')
        
        for col_num, col_name in enumerate(df.columns):
            cell_value = row_data[col_name]
            col_idx_gspread = col_num

            # デフォルト書式（基本 + 必要なら背景ハイライト）
            current_cell_format = copy.deepcopy(fmt_default)
            if is_highlight_row:
                current_cell_format.update(fmt_highlight_bg)

            # --- Excelと同じ色付けロジック ---
            if col_name == "ステータス":
                if cell_value == "異常なし":
                    current_cell_format.update(fmt_text_blue)
                else: # "要確認"
                    current_cell_format.update(fmt_text_red)
            
            elif col_name in ["テキスト比較", "誤字脱字", "内容量比較", "エラー検出"]:
                if cell_value == "OK！":
                    current_cell_format.update(fmt_text_blue)
                elif cell_value in ["差分あり", "要確認"] or \
                       (col_name == "誤字脱字" and "OK！" not in str(cell_value)) or \
                       (col_name == "エラー検出" and str(cell_value) != ""): 
                    current_cell_format.update(fmt_text_red)
                elif cell_value in ["比較対象なし", "内容量記載なし"]:
                    current_cell_format.update(fmt_text_gray)
                elif cell_value != "": 
                    current_cell_format.update(fmt_text_red)
            
            elif '（画像）' in col_name:
                # 画像列は中央揃えを適用
                current_cell_format.update(IMAGE_CELL_FORMAT_GS)
                # 注: =IMAGE() 関数自体に色は付かない
            
            cell_format_requests.append(
                get_cell_format_request(sheet_id, row_idx_gspread, col_idx_gspread, current_cell_format)
            )

    # --- 5. バッチアップデート実行 (チャンク化) ---
    
    # チャンクサイズ (一度に送信するリクエスト数)
    CHUNK_SIZE = 100 
    
    # 最初に列幅・行高・ヘッダー書式を適用
    if requests:
        body = {'requests': requests}
        try:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
        except HttpError as e:
            st.error(f"スプレッドシートの基本書式設定に失敗しました: {e}")
            raise # 致命的なエラーとしてスロー 
            
    # 次に、セルごとの書式設定をチャンクに分けて送信
    if cell_format_requests:
        total_chunks = (len(cell_format_requests) + CHUNK_SIZE - 1) // CHUNK_SIZE
        
        for i in range(0, len(cell_format_requests), CHUNK_SIZE):
            chunk = cell_format_requests[i:i + CHUNK_SIZE]
            body = {'requests': chunk}
            
            try:
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=body
                ).execute()
            except HttpError as e:
                st.error(f"スプレッドシートの書式設定に失敗しました (Chunk {i//CHUNK_SIZE + 1}): {e}")
                st.warning("接続エラーが発生したため、一部の色付けが不完全な可能性があります。")
                break 
            except Exception as e:
                st.error(f"スプレッドシートの書式設定中に予期せぬエラー (Chunk {i//CHUNK_SIZE + 1}): {e}")
                break


def save_to_spreadsheet(df_excel, spreadsheet_id, sheet_name, creds_info, portal_files, image_bytes_data):
    """
    既存のスプレッドシートIDに、指定したシート名で新しいシートを作成し、
    データを書き込む (サービスアカウント使用)
    """
    
    # サービスアカウントの「辞書」から各サービスをビルド
    user_drive_service, user_sheets_service_v4, gc = get_google_services(creds_info) 
    
    if not user_drive_service or not gc or not user_sheets_service_v4:
        st.error("Googleサービスへの接続に失敗しました。")
        return

    try:
        with st.spinner(f"スプレッドシートを開き、「{sheet_name}」シートを準備中..."):
            # 1. スプレッドシートを開く
            try:
                sh = gc.open_by_key(spreadsheet_id)
            except gspread.exceptions.SpreadsheetNotFound:
                st.error("スプレッドシートが見つかりません。URLが正しいか、サービスアカウントに編集権限が付与されているか確認してください。")
                return
            except Exception as e:
                st.error(f"スプレッドシートを開けませんでした: {e}")
                return

            # 2. ワークシート（タブ）の準備
            worksheet_title = sheet_name
            
            try:
                # 同名のシートが既に存在するか確認
                worksheet = sh.worksheet(worksheet_title)
                # 存在したらクリア
                worksheet.clear() 
                # サイズ変更 (行数+1はヘッダー分)
                worksheet.resize(rows=len(df_excel) + 1, cols=len(df_excel.columns))
            except gspread.exceptions.WorksheetNotFound:
                # 存在しなければ作成
                worksheet = sh.add_worksheet(title=worksheet_title, rows=len(df_excel) + 1, cols=len(df_excel.columns))
            except Exception as e:
                st.error(f"シート「{worksheet_title}」の準備中にエラーが発生しました: {e}")
                return

        with st.spinner("スプレッドシートにデータを書き込み中..."):
            # --- データ書き込み準備 ---
            df_excel_gspread = df_excel.fillna('').copy()
            
            # --- [修正] =IMAGE() 関数を使用するロジック (変更なし) ---
            for col_name in df_excel_gspread.columns:
                if '（画像）' in col_name:
                    # 元のURL (https://drive.google.com/file/d/FILE_ID/view) から file_id を抽出
                    file_id_series = df_excel_gspread[col_name].apply(
                        lambda url: re.search(r'/d/([a-zA-Z0-9_-]+)', str(url))
                    )
                    
                    # --- ▼▼▼ [変更] =IMAGE(URL, 4, 高さ, 幅) 形式の文字列を生成 ▼▼▼ ---
                    df_excel_gspread[col_name] = file_id_series.apply(
                        lambda match: f'=IMAGE("https://drive.google.com/uc?id={match.group(1)}")' if match else ""
                    )
                    # --- ▲▲▲ [変更] ここまで ▲▲▲ ---
            
            headers = df_excel_gspread.columns.values.tolist()
            data_values = df_excel_gspread.values.tolist()
            values_to_update = [headers] + data_values
            
            worksheet.update(
                values_to_update,
                value_input_option='USER_ENTERED' # これで =IMAGE() が関数として解釈される
            )
        
        with st.spinner("スプレッドシートの書式設定中..."):
            # 書式設定 (df_excel (元の値) を渡して判定させる)
            format_worksheet_gspread(user_sheets_service_v4, spreadsheet_id, worksheet.id, df_excel, portal_files)

        # 実行後のURLを生成 (シートIDを指定)
        sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={worksheet.id}"

        st.toast(f"シート「{sheet_name}」に保存しました！", icon="✅")
        #st.success(f"スプレッドシートに保存しました: [開く]({sheet_url})", icon="📄")

    except Exception as e:
        st.error(f"スプレッドシートへの書き込みまたは書式設定中にエラーが発生しました: {e}")
        # 失敗した場合でも、作成途中のシートへのリンクを表示
        sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
        st.warning(f"データは保存されましたが、書式が適用されていない可能性があります。 [スプレッドシートリンク]({sheet_url})")