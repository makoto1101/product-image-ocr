import streamlit as st
import pandas as pd
import io
import re
import json
import base64
import datetime
from openai import AsyncOpenAI
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import streamlit.components.v1 as components
import asyncio
from functools import partial
import math
import requests
import os

# --- ローカルモジュールのインポート ---
from neng_api import get_neng_content
# [削除] create_excel_output のインポートを削除
from export import save_to_spreadsheet
# [追加] 操作マニュアルのインポート
from manual import show_instructions


# --- Streamlit ページ設定 ---
st.set_page_config(
    page_title="商品画像OCR",
    page_icon="static/ocr_img.png",
    layout="wide"
)

# 画像ファイルを読み込み、Base64にエンコードする関数
def get_image_as_base64(path):
    # スクリプトの絶対パスを基準に画像パスを指定
    script_dir = os.path.dirname(os.path.abspath(__file__))
    img_path = os.path.join(script_dir, path)
    
    if not os.path.exists(img_path):
        return None # 画像がない場合はNoneを返す

    with open(img_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

# --- 画像の準備 ---
# page_iconはパスのままでOK
page_icon_path = "static/ocr_img.png"

# h1タグ埋め込み用のBase64文字列を生成
img_base64 = get_image_as_base64(page_icon_path)
if img_base64:
    # f-stringで埋め込めるように、data URI形式にする
    img_src = f"data:image/png;base64,{img_base64}"
else:
    # 画像が見つからなかった場合の代替テキスト
    img_src = "" 
    st.error(f"画像ファイルが見つかりません: {page_icon_path}")


st.set_page_config(
    page_title="商品画像OCR",
    page_icon=page_icon_path, # こちらは元のパス指定のままでOK
    layout="wide"
)

# === ログイン画面の表示 ===
if not st.user.get("is_logged_in", False): 
    
    # img_srcが空でない（画像が正しく読み込めた）場合のみ画像タグを表示
    image_tag = ""
    if img_src:
        image_tag = f'<img src="{img_src}" style="vertical-align: middle; height: 80px;">'

    st.markdown(f"""
        <h1 style='font-size: 40px; text-align: center; margin-bottom: 30px;'>
            {image_tag}
            商品画像OCR
        </h1>
    """, unsafe_allow_html=True)

    # カラムを使って中央に配置
    _, form_col, _ = st.columns([3, 2, 3])
    with form_col:
        st.warning('Googleアカウントでログインしてください。')

        # 1. Streamlitのボタンを表示
        if st.button("Googleアカウントでログイン", icon=":material/login:", width='stretch'):
            # 2. ボタンが押されたら、st.login() を呼び出し（認証プロセスを開始）
            st.login() # Google認証画面へリダイレクト

    # 未ログイン時はここでスクリプトの実行を停止
    st.stop()

# === ログイン成功後のメインアプリ ===
else: # Google認証済みの場合のみ以下を実行

    # --- CSSファイルを読み込む関数 ---
    def load_css(file_name):
        try:
            with open(file_name, "r", encoding="utf-8") as f:
                st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
        except FileNotFoundError:
            st.error(f"CSSファイル '{file_name}' が見つかりません。app.pyと同じフォルダに配置してください。")

    load_css("style.css")

    # --- サイドバーを一番下にスクロールする関数 ---
    def scroll_sidebar_to_bottom():
        js = """
        <script>
            setTimeout(function() {
                const sidebar = window.parent.document.querySelector('[data-testid="stSidebar"] > div:first-child');
                if (sidebar) {
                    sidebar.scrollTop = sidebar.scrollHeight;
                }
            }, 100);
        </script>
        """
        components.html(js, height=0)

    # --- ページ全体を一番下にスクロールする関数 ---
    def scroll_page_to_bottom():
        """ページ全体を一番下にスクロールするJSを実行"""
        js = """
        <script>
            setTimeout(function() {
                # ページ全体のドキュメント要素とbody要素の両方に対してスクロールを実行
                window.parent.document.body.scrollTop = window.parent.document.body.scrollHeight;
                window.parent.document.documentElement.scrollTop = window.parent.document.documentElement.scrollHeight;
            }, 100);
        </script>
        """
        components.html(js, height=0)

    # --- 状態管理 (Session State) の初期化 ---
    # ログイン関連のsession_stateを削除 ('logged_in', 'username', 'just_logged_out')
    if 'drive_folder_id' not in st.session_state:
        st.session_state.drive_folder_id = ""
    if 'portal_files' not in st.session_state:
        st.session_state.portal_files = None
    if 'business_codes' not in st.session_state:
        st.session_state.business_codes = []
    if 'product_codes' not in st.session_state:
        st.session_state.product_codes = []
    if 'ocr_result_df' not in st.session_state:
        st.session_state.ocr_result_df = None
    if 'ocr_plain_df' not in st.session_state: # 検索用の平文DF
        st.session_state.ocr_plain_df = None
    # [削除] ocr_excel_output を削除
    if 'ocr_excel_df' not in st.session_state: # スプレッドシート保存用の元DF
        st.session_state.ocr_excel_df = None
    if 'ocr_image_bytes' not in st.session_state: # [追加] 画像バイナリデータ
        st.session_state.ocr_image_bytes = None
    if 'show_ocr_confirmation' not in st.session_state:
        st.session_state.show_ocr_confirmation = False
    if 'record_count_to_process' not in st.session_state:
        st.session_state.record_count_to_process = 0
    if 'image_total_count_to_process' not in st.session_state:
        st.session_state.image_total_count_to_process = 0
    if 'municipality_map' not in st.session_state:
        st.session_state.municipality_map = {}
    if 'db_update_count' not in st.session_state:
        st.session_state.db_update_count = 0
    if 'current_page' not in st.session_state:
        st.session_state.current_page = 1

    # 変更検知用のキー
    if 'old_municipality' not in st.session_state:
        st.session_state.old_municipality = None
    if 'old_business_code' not in st.session_state:
        st.session_state.old_business_code = None
    if 'old_product_code' not in st.session_state:
        st.session_state.old_product_code = None
    if 'show_clear_confirmation' not in st.session_state:
        st.session_state.show_clear_confirmation = False
    if 'pending_change' not in st.session_state:
        st.session_state.pending_change = None
    # ドライブ読み込み確認用のキー
    if 'show_drive_clear_confirmation' not in st.session_state:
        st.session_state.show_drive_clear_confirmation = False
    if 'execute_drive_load_now' not in st.session_state:
        st.session_state.execute_drive_load_now = False

    # --- スプレッドシート保存先入力用のキー ---
    if 'show_gspread_url_input' not in st.session_state:
        st.session_state.show_gspread_url_input = False
    if 'gspread_sheet_url_input' not in st.session_state: 
        st.session_state.gspread_sheet_url_input = ""
    if 'gspread_save_success_url' not in st.session_state:
        st.session_state.gspread_save_success_url = None
    if 'gspread_save_error_message' not in st.session_state:
        st.session_state.gspread_save_error_message = None


    # --- スプレッドシートから自治体リストとコードを取得する関数 ---
    @st.cache_data(ttl=3600) # 1時間キャッシュ
    def get_municipality_map(_sheets_service):
        """スプレッドシートから自治体名とコードのマップを取得する"""
        SPREADSHEET_ID = '1n8qDS8OvuFJwDy2J6wduDHI32GxDmbx1QIrqHPFjdGo'
        RANGE_NAME = '自治体DB!A2:B'
        municipality_map = {}
        try:
            sheet = _sheets_service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
            values = result.get('values', [])

            if not values:
                # st.error はサイドバーではなくメイン画面に出てしまうため、サイドバーのエラーに変更
                st.sidebar.error("スプレッドシートから自治体データを取得できませんでした。")
                return {}

            # { "自治体名": "コード", ... } の辞書を作成
            for row in values:
                if len(row) >= 2 and row[0] and row[1]: # 名前とコードが両方存在する場合のみ
                    municipality_map[row[0]] = row[1]

            # 名前順にソートした辞書を返す
            sorted_map = dict(sorted(municipality_map.items()))
            return sorted_map

        except HttpError as err:
            st.sidebar.error(f"自治体リストのスプレッドシートへのアクセスに失敗しました: {err}")
            return {}
        except Exception as e:
            st.sidebar.error(f"自治体リスト取得中にエラーが発生しました: {e}")
            return {}

    # --- 変更検知・確認ダイアログ用コールバック関数 ---

    def _check_for_change_and_warn(key_name, old_key_name):
        """
        selectboxのon_changeコールバック。
        変更を検知し、結果が存在する場合は確認フラグを立て、ウィジェットの値を元に戻す。
        """

        # OCR確認メッセージが表示されていれば消す
        st.session_state.show_ocr_confirmation = False

        new_value = st.session_state.get(key_name)
        old_value = st.session_state.get(old_key_name)

        if new_value == old_value or st.session_state.show_clear_confirmation:
            return

        if st.session_state.ocr_result_df is not None:
            st.session_state.pending_change = {"key": key_name, "old_key": old_key_name, "new_value": new_value}
            st.session_state.show_clear_confirmation = True
            st.session_state[key_name] = old_value # 警告中は値を戻す
        else:
            st.session_state[old_key_name] = new_value # 警告なしなら old の値を更新

    def _confirm_clear_results():
        """
        確認ダイアログで「OK」が押された時の処理
        """
        st.session_state.ocr_result_df = None
        st.session_state.ocr_plain_df = None
        # [削除] ocr_excel_output のクリアを削除
        st.session_state.ocr_excel_df = None 
        st.session_state.ocr_image_bytes = None # [追加]
        st.session_state.current_page = 1

        if st.session_state.pending_change:
            key_name = st.session_state.pending_change["key"]
            old_key_name = st.session_state.pending_change["old_key"]
            new_value = st.session_state.pending_change["new_value"]

            # 保留していた変更を適用
            st.session_state[key_name] = new_value
            st.session_state[old_key_name] = new_value

            # 事業者コードが変更された場合、品番を「すべて」に戻す
            if key_name == "business_select_key":
                st.session_state.product_select_key = "すべて"
                st.session_state.old_product_code = "すべて"

        st.session_state.show_clear_confirmation = False
        st.session_state.pending_change = None

    def _cancel_clear_results():
        """
        確認ダイアログで「キャンセル」が押された時の処理
        """
        st.session_state.show_clear_confirmation = False
        st.session_state.pending_change = None # 保留していた変更は破棄

    def _trigger_drive_load_check():
        """
        「ドライブから読み込み」ボタンのon_clickコールバック
        """

        # OCR確認メッセージが表示されていれば消す
        st.session_state.show_ocr_confirmation = False

        if st.session_state.ocr_result_df is not None:
            st.session_state.show_drive_clear_confirmation = True
        else:
            st.session_state.execute_drive_load_now = True

    def _execute_drive_load():
        """
        実際のドライブ読み込み処理（確認ダイアログOK時、または結果なしのボタンクリック時）
        短縮URLにも対応。
        """
        drive_url_input = st.session_state.get("drive_url_input_key", "")

        st.session_state.show_ocr_confirmation = False
        st.session_state.current_page = 1

        # ドライブ読み込み時は結果と選択状態をリセット
        st.session_state.ocr_result_df = None
        st.session_state.ocr_plain_df = None
        # [削除] ocr_excel_output のクリアを削除
        st.session_state.ocr_excel_df = None 
        st.session_state.ocr_image_bytes = None # [追加]
        st.session_state.old_municipality = None
        st.session_state.old_business_code = None
        st.session_state.old_product_code = None
        st.session_state.portal_files = None
        st.session_state.business_codes = []
        # selectboxのキーもリセット
        st.session_state.pop("municipality_select_key", None)
        st.session_state.pop("business_select_key", None)
        st.session_state.pop("product_select_key", None)

        if not drive_url_input:
            st.warning("読み込むフォルダのURLを入力してください。")
            st.session_state.show_drive_clear_confirmation = False # 確認ダイアログが開いていれば閉じる
            return # URLがないので処理終了

        # --- URL解決とID抽出処理 ---
        final_drive_url = None
        drive_folder_id = None

        # まず入力されたURLが直接Google Driveフォルダの形式かチェック
        if 'drive.google.com' in drive_url_input and '/folders/' in drive_url_input:
            final_drive_url = drive_url_input
            drive_folder_id = get_folder_id_from_url(final_drive_url)
        else:
            # Google Driveの直接URLでない場合、短縮URLの可能性を考慮して解決を試みる
            st.info("URLを解決中...") # ユーザーに進捗を表示
            resolved_url = resolve_url(drive_url_input)
            st.info("") # メッセージを消す

            if resolved_url:
                # 解決後のURLがGoogle Driveフォルダの形式かチェック
                if 'drive.google.com' in resolved_url and '/folders/' in resolved_url:
                    final_drive_url = resolved_url
                    drive_folder_id = get_folder_id_from_url(final_drive_url)
                    st.success(f"短縮URLを解決しました: {final_drive_url}") # 解決後のURLを表示
                else:
                    # 解決したがGoogle DriveフォルダURLではなかった
                    st.error("入力されたURL、またはリダイレクト先が有効なGoogleドライブのフォルダURLではありません。")
            # else:
                # resolve_url内でエラーメッセージ表示済み

        # フォルダIDが取得できた場合のみ、ファイルリスト取得処理へ進む
        if drive_folder_id:
            st.session_state.drive_folder_id = drive_folder_id
            with st.spinner("Google Driveをスキャン中..."):
                portal_files, business_codes = list_drive_files_and_business_codes(drive_folder_id)
                st.session_state.portal_files = portal_files
                st.session_state.business_codes = business_codes
                if business_codes:
                    st.toast("Googleドライブの読み込みに成功しました。", icon="✅")
                else: # portal_files はあるが business_codes が空の場合
                    if portal_files is None: # list_drive_files_and_business_codes が None を返した場合
                        pass # エラーメッセージは関数内で表示済み
                    else: # ファイルはあるが対象コードがない場合
                        st.warning("画像ファイルは見つかりましたが、事業者コード・品番を抽出できませんでした。ファイル名を確認してください。")
        elif final_drive_url: # Google Drive URLではあったがID抽出に失敗した場合 (形式が違うなど)
            st.error("GoogleドライブのフォルダURL形式が無効です。`/folders/` を含むURLを入力してください。")

        # 確認ダイアログが開いていれば閉じる
        st.session_state.show_drive_clear_confirmation = False

    def _cancel_drive_load():
        """
        ドライブ読み込みキャンセルの処理
        """
        st.session_state.show_drive_clear_confirmation = False


    # --- APIクライアントと認証情報の取得 ---
    @st.cache_resource
    def get_google_credentials():
        try:
            google_credentials_info = json.loads(st.secrets["google"]["credentials_json"])
            creds = service_account.Credentials.from_service_account_info(
                google_credentials_info,
                scopes=[
                    'https://www.googleapis.com/auth/drive',
                    'https://www.googleapis.com/auth/spreadsheets'
                ]
            )
            
            # --- 辞書とオブジェクトの両方を返す ---
            return creds, google_credentials_info
        except (KeyError, json.JSONDecodeError, FileNotFoundError) as e:
            st.error(f"secrets.tomlの読み込みまたは設定にエラーがあります。エラー: {e}")
            st.stop()
            return None, None # 戻り値を2つに

    def get_drive_service(_credentials):
        return build('drive', 'v3', credentials=_credentials)

    def get_sheets_service(_credentials):
        return build('sheets', 'v4', credentials=_credentials)

    try:
        # --- 戻り値を2つ受け取る ---
        google_creds, google_creds_info = get_google_credentials()
        drive_service = get_drive_service(google_creds)
        sheets_service = get_sheets_service(google_creds)
        async_openai_client = AsyncOpenAI(api_key=st.secrets["openai"]["api_key"])

        # --- アプリ起動時に自治体マップを読み込む ---
        if 'municipality_map' not in st.session_state or not st.session_state.municipality_map:
            # メイン画面ではなくサイドバーにスピナーを表示する
            with st.sidebar:
                with st.spinner("自治体リストを読み込み中..."):
                    st.session_state.municipality_map = get_municipality_map(sheets_service)

    except Exception as e:
        st.error(f"APIキーまたは認証情報の読み込み中にエラーが発生しました: {e}")
        st.stop()

    def resolve_url(url):
        """
        URLにアクセスし、リダイレクトがあれば最終的なURLを返す。
        短縮URLの解決などに使用。
        """
        try:
            # HEADリクエストでリダイレクトを追跡 (コンテンツはダウンロードしない)
            # タイムアウトを設定
            response = requests.head(url, allow_redirects=True, timeout=10)
            response.raise_for_status() # HTTPエラーがあれば例外を発生させる
            return response.url # 最終的なURLを返す
        except requests.exceptions.Timeout:
            st.error(f"URLへのアクセスがタイムアウトしました: {url}")
            return None
        except requests.exceptions.ConnectionError:
            st.error(f"URLに接続できませんでした。ネットワーク接続を確認してください: {url}")
            return None
        except requests.exceptions.HTTPError as e:
            st.error(f"URLへのアクセス中にHTTPエラーが発生しました ({e.response.status_code}): {url}")
            return None
        except requests.exceptions.RequestException as e:
            st.error(f"URLの解決中にエラーが発生しました: {e}")
            return None

    # --- ヘルパー関数群 ---
    def get_folder_id_from_url(url):
        match = re.search(r'folders/([a-zA-Z0-9_-]+)', url)
        return match.group(1) if match else None
    
    # --- [追加] スプレッドシートURLからIDを抽出する関数 ---
    def get_spreadsheet_id_from_url(url):
        # /d/ の後から、次の / までを抽出
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        return match.group(1) if match else None

    # @st.cache_data 
    def list_drive_files_and_business_codes(drive_folder_id):
        portal_files = {}
        business_codes = set()
        try:
            # まず指定されたフォルダ自体を取得（存在確認と名前取得のため）
            # --- 共有ドライブ対応 ---
            folder_info = drive_service.files().get(
                fileId=drive_folder_id, 
                fields="id, name",
                supportsAllDrives=True # 共有ドライブ対応
            ).execute()

            # サブフォルダを検索
            subfolders_query = f"'{drive_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
            subfolders_response = drive_service.files().list(
                q=subfolders_query, 
                fields="files(id, name)",
                supportsAllDrives=True, # 共有ドライブ対応
                includeItemsFromAllDrives=True # 共有ドライブ対応
            ).execute()
            subfolders = subfolders_response.get('files', [])

            # サブフォルダがあればそれらを処理、なければ指定されたフォルダ自体を処理対象に
            folders_to_process = subfolders if subfolders else [folder_info]

            total_image_count = 0
            for folder in folders_to_process:
                portal_name = folder['name']
                portal_files[portal_name] = []
                # フォルダ内の画像ファイルを検索
                files_query = f"'{folder['id']}' in parents and (mimeType='image/jpeg' or mimeType='image/png')"
                files_response = drive_service.files().list(
                    q=files_query, 
                    fields="files(id, name, mimeType)",
                    supportsAllDrives=True, # 共有ドライブ対応
                    includeItemsFromAllDrives=True # 共有ドライブ対応
                ).execute()
                files_in_folder = files_response.get('files', [])
                total_image_count += len(files_in_folder)

                for file in files_in_folder:
                    portal_files[portal_name].append({'id': file['id'], 'name': file['name'], 'mimeType': file['mimeType']})
                    # ファイル名から事業者コードを抽出
                    prod_code = get_product_code_from_filename(file['name'])
                    bus_code = get_business_code_from_product_code(prod_code)
                    if bus_code:
                        business_codes.add(bus_code)

            # 画像ファイルが1つも見つからなかった場合
            if total_image_count == 0:
                st.warning("指定されたGoogleドライブフォルダ（またはそのサブフォルダ）内に、処理対象の画像ファイル（.jpg, .png）が見つかりませんでした。")
                return None, [] 

            # 画像ファイルはあったが事業者コードが抽出できなかった場合
            if not business_codes and total_image_count > 0:
                st.warning("画像ファイルは見つかりましたが、ファイル名から事業者コードを抽出できませんでした。ファイル名の形式を確認してください。")

            return portal_files, sorted(list(business_codes))

        except HttpError as e:
            if e.resp.status == 404:
                st.error("指定されたフォルダが見つからないか、アクセス権限がありません。URLを確認してください。")
            elif e.resp.status == 403:
                st.error("Google Drive APIへのアクセス権限がありません。サービスアカウントの設定や共有設定を確認してください。")
            else:
                st.error(f"Google Driveからのファイル一覧取得中にエラーが発生しました: {e}")
            return None, []
        except Exception as e:
            st.error(f"ファイル一覧の処理中に予期せぬエラーが発生しました: {e}")
            return None, []


    def get_product_code_from_filename(filename):
        # 拡張子を除去
        name_without_ext = filename.rsplit('.', 1)[0]
        # 最初のハイフンまでを取得（ハイフンがない場合は全体）
        return name_without_ext.split('-')[0]

    def get_business_code_from_product_code(product_code):
        if not product_code: return None
        # 正規表現パターン (大文字小文字を区別しない)
        # 1. 数字2桁 + 英字4桁 (例: 01ABCD)
        # 2. 英字4桁 (例: ABCD)
        # 3. 英字3桁 (例: ABC)
        patterns = [r'^[0-9]{2}[a-zA-Z]{4}', r'^[a-zA-Z]{4}', r'^[a-zA-Z]{3}']
        for p in patterns:
            match = re.match(p, product_code) # re.IGNORECASE は不要かも
            if match:
                return match.group(0).upper() # マッチした部分を大文字で返す
        return None # どのパターンにもマッチしない場合

    def get_product_codes_for_business_code(portal_files, selected_business_code):
        if not portal_files or not selected_business_code: return []
        product_codes_set = set()
        for portal_name, files in portal_files.items():
            for file in files:
                full_product_code = get_product_code_from_filename(file['name'])
                business_code = get_business_code_from_product_code(full_product_code)
                # 選択された事業者コードと一致する場合のみ追加
                if business_code == selected_business_code:
                    product_codes_set.add(full_product_code)
        return sorted(list(product_codes_set))


    def count_images_to_process(portal_files, selected_business_code, selected_product_code):
        if not portal_files: return 0, 0
        image_records = set() # 画像名 (ユニークなファイル名) をカウント
        image_total_count = 0 # ポータルごとの画像の合計枚数
        for portal_name, files in portal_files.items():
            for file in files:
                full_product_code = get_product_code_from_filename(file['name'])
                business_code = get_business_code_from_product_code(full_product_code)
                # 事業者コードが一致し、かつ (品番が「すべて」 OR 品番が一致) する場合
                if business_code == selected_business_code and \
                   (selected_product_code == "すべて" or full_product_code == selected_product_code):
                    image_records.add(file['name'])
                    image_total_count += 1
        return len(image_records), image_total_count

    # --- 非同期処理の定義 (OpenAI API関連) ---
    async def call_openai_vision_api_async(client, prompt, image_base64, mime_type, model="gpt-4o", max_tokens=1000):
        try:
            messages = [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}}]}]
            response = await client.chat.completions.create(model=model, messages=messages, temperature=0.0, max_tokens=max_tokens)
            return response.choices[0].message.content
        except Exception as e: return f"OpenAI APIエラー: {e}"

    async def call_openai_text_api_async(client, prompt, model="gpt-4o"):
        try:
            messages = [{"role": "user", "content": prompt}]
            response = await client.chat.completions.create(model=model, messages=messages, temperature=0.0, response_format={"type": "json_object"})
            return response.choices[0].message.content
        except Exception as e: return f'{{"status": "api_error", "message": "OpenAI APIエラー: {e}"}}'

    async def call_openai_simple_text_api_async(client, prompt, model="gpt-4o"):
        try:
            messages = [{"role": "user", "content": prompt}]
            response = await client.chat.completions.create(model=model, messages=messages, temperature=0.0)
            return response.choices[0].message.content
        except Exception as e: return f"OpenAI APIエラー: {e}"

    async def check_typos_async(client, texts):
        filtered_texts = [t for t in texts if t and "テキストは検出されませんでした。" not in t and "APIエラー" not in t and "予期せぬエラー" not in t]
        if not filtered_texts: return "OK！"
        prompt = f"""あなたは日本語の「誤字」と「脱字」を厳密に発見する校正AIです。以下の広告文テキストをチェックしてください。
ルール:
1. 表現、スタイル、句読点、文法に関する指摘は絶対にしないでください。
2. 数字や広告デザインによる意図的な語順は問題ないと判断してください。
3. **テキストは単語がスペースなしで連結されています。単語の区切りがないことで不自然に見える文字列は、誤字脱字として指摘しないでください。**
4. 明確な誤字のみを指摘してください。
判断:
- 上記ルールに反する問題が一切なければ {{"status": "ok"}} を返してください。
- 誤字脱字がある場合のみ、{{"status": "error", "message": "「(問題のある単語のみ)」を確認"}} という形式で返してください。
- 必ずJSON形式で応答してください。
---テキスト---
{'\n---\n'.join(filtered_texts)}"""
        response_str = await call_openai_text_api_async(client, prompt)
        try:
            result_json = json.loads(response_str)
            if result_json.get("status") == "ok": return "OK！"
            elif result_json.get("status") == "error": return result_json.get("message", "エラー")
            else: return "不明" # APIが予期しない形式で返した場合
        except (json.JSONDecodeError, AttributeError): return "解析不能" # JSON解析失敗など

    async def extract_content_volume_async(client, portal_name, text):
        # OCR結果がない、またはエラー文字列の場合は空を返す
        if not text or "テキストは検出されませんでした。" in text or "APIエラー" in text:
            return portal_name, ""
        prompt = f"""あなたは、テキストから商品の「数量」や「重量」に関する部分のみを正確に抜き出す専門家です。
以下のルールに従って、与えられたテキストから内容量を示す数値と単位の部分だけを抽出してください。

### ルール
1. **抽出対象:** 「〇個」「〇ml×〇個」「〇kg」「〇~〇本」のような、**数量、重さ、個数を示す部分のみ**を抽出します。
2. **商品名は除外:** 「いちごソルベ」「安納芋」といった商品名は絶対に含めないでください。
3. **完全な維持:** 抽出するテキストは、元のテキストに含まれる文字、数字、記号、改行(\\n)を完全に維持してください。一文字も変更、追加、削除してはいけません。
4. **除外対象:** 「お届け内容」「セット内容」といった見出しや、商品の特徴・産地などの説明的な文章は抽出しないでください。
5. **出力形式:** 抽出したテキストだけを返してください。余計な説明や前置きは一切含めないでください。
6. **該当なしの場合:** 内容量に関する記述が見つからない場合は、必ず**空の文字列**を返してください。

### 例
- 元テキスト: "お届け内容\\nいちごソルベ\\n90ml×6個"
- 抽出結果: "90ml×6個"
---
元テキスト:
"{text}"
"""
        response = await call_openai_simple_text_api_async(client, prompt)
        # レスポンスの前後の空白を除去し、エスケープされた改行を実際の改行に置換
        processed_response = response.strip().replace('\\n', '\n')
        return portal_name, processed_response

    async def compare_content_volume_async(client, base_content, portal_contents):
        """AIを使用して、基準となる内容量と複数の比較対象内容量が一致するか判定する"""
        # 空でない有効な内容量テキストのみを抽出
        valid_portal_contents = [c for c in portal_contents if c and c.strip()]

        # 比較対象となるポータルの内容量が一つもなければ「記載なし」
        if not valid_portal_contents:
            return "内容量記載なし"

        # NENGの内容量（基準）が空なら「要確認」
        if not base_content:
            return "要確認"

        prompt = f"""あなたは、商品の内容量テキストを解釈し、意味が同じかどうかを判定する専門家AIです。
あなたのタスクは、以下の思考プロセスに従って、「基準テキスト」と「比較テキストリスト」の内容が実質的に同じか判断することです。
### 思考プロセス
1.  **要素の抽出**: まず、各テキストから内容量を構成する「数値」と「単位」のペアを全て抽出します。（例：「90ml×6個」からは「90ml」と「6個」の2つの要素を抽出）
2.  **構造化**: 抽出した各要素を、標準的なJSONオブジェクトに変換します。数値が範囲（例：4～6本）の場合は `"range": "4-6"` のように表現します。
3.  **正規化**: 変換したJSONオブジェクトの配列を、unit（単位）のアルファベット順で並べ替えます。これにより、「2kg 4~6本」と「4~6本 2kg」が同じJSON表現になります。
4.  **比較**: 「基準テキスト」から生成した正規化JSONと、「比較テキストリスト」の各テキストから生成した正規化JSONが、全て完全に一致するかどうかを比較します。
5.  **最終判断**: 全ての比較テキストの正規化JSONが、基準テキストの正規化JSONと一致した場合のみ `{{"result": "ok"}}` とします。一つでも不一致があれば `{{"result": "ng"}}` とします。
### 絶対的なルール
* **計算の禁止**: テキストに表記されている数値をそのまま使ってください。**絶対に計算してはいけません。** 「1.2kg」と「600g×2」は、表記が異なるため「不一致」です。
* **記号の統一**: 範囲を示す記号（`~`, `～`, `-`）は、全て半角ハイフン `-` に統一して `range` を作ります。
### 例
-   **入力**:
    * 基準テキスト: "2kg(4～6本)"
    * 比較テキストリスト: ["2kg\\n4-6本"]
-   **あなたの応答**: `{{"result": "ok"}}`
---
-   **入力**:
    * 基準テキスト: "合計1.2kg"
    * 比較テキストリスト: ["600g×2パック"]
-   **あなたの応答**: `{{"result": "ng"}}`
---
それでは、以下のテキストを比較してください。
### 基準テキスト
{base_content}
### 比較テキストリスト
{json.dumps(valid_portal_contents, ensure_ascii=False)}
"""
        response_str = await call_openai_text_api_async(client, prompt)
        try:
            result_json = json.loads(response_str)
            return "OK！" if result_json.get("result") == "ok" else "要確認"
        except (json.JSONDecodeError, AttributeError):
            return "要確認" # JSON解析失敗や result キーがない場合は「要確認」扱い

    def download_drive_image_sync(file_id, credentials):
        """同期的にGoogle Driveから画像データをダウンロードする"""
        drive = build('drive', 'v3', credentials=credentials)
        # --- 共有ドライブ対応 ---
        request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
        return request.execute() # 画像のバイナリデータを返す

    async def extract_text_from_drive_image_async(portal_name, file_id, mime_type, credentials, client):
        """非同期でDrive画像を取得し、OpenAI Vision APIでOCRを実行"""
        image_bytes = None # 初期化
        try:
            loop = asyncio.get_running_loop()
            # 同期的なダウンロード処理を非同期イベントループで実行
            image_bytes = await loop.run_in_executor(None, partial(download_drive_image_sync, file_id, credentials))
            # Base64エンコード
            image_base64 = base64.b64encode(image_bytes).decode('utf-8')
        except HttpError as e:
            return portal_name, f"Google Drive画像取得失敗 (HttpError {e.resp.status})", None
        except Exception as e:
            return portal_name, f"Google Drive画像取得失敗: {e}", None # その他のエラー

        # --- プロンプト ---
        prompt = """あなたは、商品広告画像のテキストを人間が読む通りに正確に書き起こす専門家です。
あなたのタスク:
渡された画像の中から全てのテキストを読み取り、人間が目で追う自然な順序（上から下、左から右）に並べ替えてください。そして、単語や意味のまとまり（フレーズ）ごとに改行を入れて出力してください。
書き起こしと改行のルール:
1. 読む順序は厳密に「上から下へ、左から右へ」です。画像のレイアウトを最優先してください。
2. デザイン上の理由で分離している文字（例：「ギ」と「ュ」と「っと」）は、意味が通じるように自然な1つの単語（例：「ギュっと」）として結合してください。
3. 結合した後の単語やフレーズは、それぞれ独立した行になるように改行(\\n)を挿入してください。
文字の正規化ルール:
1. 記号の統一: 「×」(掛ける記号)や「X」(大文字のエックス)は、すべて「x」(半角小文字のエックス)に統一してください。
2. 全角・半角の統一: 全角の英数字は、すべて対応する半角の文字に統一してください。(例: 「Ａ」→「A」、「３」→「3」)
絶対的なルール:
1. 画像の視覚的なレイアウトが絶対的な正解です。
2. 画像に含まれる全ての文字・数字・記号は、一切省略せず、推測して文字を追加することも絶対にしないでください。
3. 元のテキストに含まれる文字の種類（ひらがな、カタカナなど「っ」と「ッ」）や大きさ（例: 「っ」と「つ」）は、絶対に変更しないでください。
4. **応答形式:** 抽出したテキスト**のみ**を返してください。「画像内のテキストは以下の通りです」といった前置きや説明、マークダウン(` ``` `)は一切含めないでください。
5. **テキストなしの場合:** 画像にテキストが一切含まれていないと判断した場合のみ、**空の文字列**を返してください。
"""

        # OpenAI Vision API呼び出し
        response_text = await call_openai_vision_api_async(client, prompt, image_base64, mime_type)

        # --- 後処理 (簡素化) ---
        if "OpenAI APIエラー" in response_text:
             return portal_name, response_text, image_bytes # エラーメッセージをそのまま返す

        # プロンプトが無視されてマークダウンが返ってきた場合に備える
        cleaned_text = re.sub(r"```(json|text|plaintext)?\n?", "", response_text).replace("```", "")
        
        # 各行の前後の空白を除去し、空行を除外して改行で結合
        result_lines = [line.strip() for line in cleaned_text.strip().split('\n')]
        final_text = '\n'.join([line for line in result_lines if line])

        # --- AIが空文字列の代わりに '""' という文字列を返した場合の対策 ---
        if final_text == '""':
            final_text = "" # 本当の空文字列に変換する

        return portal_name, final_text, image_bytes

    # --- メインの非同期処理ワーカー ---
    async def process_single_record_async(image_name, data, selected_product_code, credentials, client, semaphore, neng_content_map):
        async with semaphore: # 同時実行数を制限
            # OCRタスクとNENG APIタスクをリストに格納
            ocr_tasks = [extract_text_from_drive_image_async(p_name, p_data['id'], p_data['mimeType'], credentials, client)
                         for p_name, p_data in data['portals'].items()]

            # NENG API呼び出しを削除し、マップから値を取得

            # 品番が「すべて」の場合はファイル名から取得、そうでなければ選択された品番を使用
            product_code_for_neng = get_product_code_from_filename(image_name) if selected_product_code == "すべて" else selected_product_code

            # マップからNENG内容量を取得（見つからない場合は空文字）
            raw_neng_content = neng_content_map.get(product_code_for_neng, "")
    
            # 品番が「すべて」の場合はファイル名から取得、そうでなければ選択された品番を使用
            product_code_for_neng = get_product_code_from_filename(image_name) if selected_product_code == "すべて" else selected_product_code
            
            # マップからNENG内容量を取得（見つからない場合は空文字）
            raw_neng_content = neng_content_map.get(product_code_for_neng, "") 

            ocr_task_results = await asyncio.gather(*ocr_tasks) # OCRタスクのみ実行

            # OCR結果と画像データを辞書に整理
            ocr_results, image_bytes_data = {}, {}
            for p_name, extracted_text, img_bytes in ocr_task_results:
                ocr_results[p_name] = extracted_text
                if img_bytes: image_bytes_data[p_name] = img_bytes # 画像データも保持

            # 内容量抽出、NENG内容量抽出、誤字脱字チェックのタスクを作成
            volume_tasks = [extract_content_volume_async(client, p_name, text)
                            for p_name, text in ocr_results.items()]
            neng_volume_task = extract_content_volume_async(client, "NENG_DUMMY", raw_neng_content) # NENG用
            typo_task = check_typos_async(client, ocr_results.values()) # 全OCR結果を渡す

            # 上記タスクを並行実行
            secondary_results = await asyncio.gather(
                asyncio.gather(*volume_tasks), # 内容量抽出タスク群
                neng_volume_task,
                typo_task
            )
            volume_task_results, neng_volume_tuple, typo_result = secondary_results

            # 結果を整理
            volume_results = {p_name: vol for p_name, vol in volume_task_results}
            _, processed_neng_content = neng_volume_tuple # NENG内容量を取得

            # NENG内容量とポータル内容量を比較するタスクを実行
            cleaned_neng_content = processed_neng_content.strip().strip('"') if processed_neng_content else ""
            portal_volumes = [v.strip().strip('"') if v else "" for v in volume_results.values()]
            comparison_result = await compare_content_volume_async(client, cleaned_neng_content, portal_volumes)

            # ポータル間のOCRテキスト比較
            ocr_texts = [text for text in ocr_results.values() if text and "テキストは検出されませんでした。" not in text and "APIエラー" not in text]
            text_comparison_result = ""
            if len(ocr_texts) <= 1:
                text_comparison_result = "比較対象なし"
            else:
                # 空白(連続含む)を単一スペースに正規化し、前後の空白を除去した上で比較
                normalized_texts = {re.sub(r'\s+', ' ', text).strip() for text in ocr_texts}
                if len(normalized_texts) == 1: # ユニークなテキストが1種類ならOK
                    text_comparison_result = "OK！"
                else:
                    text_comparison_result = "差分あり"

            return image_name, ocr_results, volume_results, image_bytes_data, typo_result, processed_neng_content, comparison_result, text_comparison_result


    async def main_async_runner(image_groups, selected_product_code, credentials, client, progress_bar, total_records, neng_content_map):
        semaphore = asyncio.Semaphore(25)
        tasks = [process_single_record_async(
                    name,
                    data,
                    selected_product_code,
                    credentials,
                    client,
                    semaphore,
                    neng_content_map 
                )
                for name, data in image_groups.items()]
        results = []
        # as_completed で完了したものから順次処理
        for i, future in enumerate(asyncio.as_completed(tasks)):
            try:
                result = await future
                results.append(result)
            except Exception as e:
                st.error(f"非同期処理中にエラーが発生しました: {e}")
            finally:
                # プログレスバーを更新
                progress_bar.progress((i + 1) / total_records, text=f"2. OCR実行中... ({i + 1}/{total_records})")
        return results

    # --- メインの実行関数 ---
    def run_ocr_process(portal_files, municipality_code, selected_business_code, selected_product_code, credentials, client, progress_bar):
        # 画像ファイル名ごとにポータル情報をグループ化
        image_groups = {}
        # NENG APIで取得するユニークな品番を収集するセット
        unique_product_codes_to_fetch = set()

        for portal_name, files in portal_files.items():
            for file in files:
                full_product_code = get_product_code_from_filename(file['name'])
                business_code = get_business_code_from_product_code(full_product_code)
                # 選択された条件に合う画像のみを対象とする
                if business_code == selected_business_code and \
                   (selected_product_code == "すべて" or full_product_code == selected_product_code):

                    if file['name'] not in image_groups:
                        image_groups[file['name']] = {'portals': {}}
                    image_groups[file['name']]['portals'][portal_name] = {'id': file['id'], 'mimeType': file['mimeType']}

                    if selected_product_code == "すべて":
                        unique_product_codes_to_fetch.add(full_product_code)
                    else:
                        unique_product_codes_to_fetch.add(selected_product_code)

        if not image_groups:
            st.warning("処理対象の画像が見つかりませんでした。")
            # [修正] 戻り値を4つに変更
            return None, None, None, None

        print(unique_product_codes_to_fetch)

        # NENG APIの事前一括呼び出し
        neng_content_map = {}
        if unique_product_codes_to_fetch:
            progress_bar.progress(0.0, text="2. NENGデータ取得中...") 
            try:
                # 非同期でNENG APIを並列実行
                neng_tasks = [get_neng_content(prod_code, municipality_code) for prod_code in unique_product_codes_to_fetch]

                async def fetch_neng_data():
                    return await asyncio.gather(*neng_tasks)

                neng_results = asyncio.run(fetch_neng_data())

                # 結果を辞書（品番: 内容量）にマッピング
                neng_content_map = dict(zip(unique_product_codes_to_fetch, neng_results))

                print(neng_content_map)

                if any("エラー" in res for res in neng_results if isinstance(res, str)):
                    st.toast("一部のNENG APIの取得でエラーが発生しました。", icon="⚠️")

            except Exception as e:
                st.error(f"NENG APIの一括取得中にエラーが発生しました: {e}")
        total_records = len(image_groups)

        all_results = asyncio.run(main_async_runner(
            image_groups,
            selected_product_code,
            credentials,
            client,
            progress_bar,
            total_records,
            neng_content_map
        ))

        # 結果をDataFrame用に整形
        results_list_display, results_list_excel = [], []

        # --- 画像バイナリデータを格納する辞書 ---
        all_image_bytes_data = {} # {image_name: {portal_name: bytes}}

        # ポータル名のリストを取得（Excelの列順のため）
        all_portal_names = sorted(list(portal_files.keys()))

        # DataFrameの列順を定義
        ordered_columns = ["画像名", "ステータス"]
        for p in all_portal_names:
            ordered_columns.extend([f"{p}（画像）", f"{p}（OCR）", f"{p}（内容量）"])
        ordered_columns.extend(["テキスト比較", "誤字脱字", "NENG内容量", "内容量比較", "エラー検出"])

        for image_name, ocr_results, volume_results, image_bytes, typo_result, neng_content, comparison_result, text_comparison_result in all_results:
            
            # --- 画像バイナリデータを辞書に格納 ---
            all_image_bytes_data[image_name] = image_bytes # image_bytes は {portal_name: bytes}
            
            row_data_display, row_data_excel = {"画像名": image_name}, {"画像名": image_name}

            image_acquisition_failed = False
            ocr_failed_for_existing_image = False

            for portal_name in all_portal_names:
                ocr_result_text = ocr_results.get(portal_name, "")
                if "画像取得失敗" in ocr_result_text:
                    image_acquisition_failed = True
                    break 

            if not image_acquisition_failed:
                for portal_name in all_portal_names:
                    img_exists = image_bytes.get(portal_name) is not None
                    ocr_result_text = ocr_results.get(portal_name, "")
                    if img_exists and (not ocr_result_text or "APIエラー" in ocr_result_text or "AI OCRエラー" in ocr_result_text):
                        ocr_failed_for_existing_image = True
                        break

            final_typo_result = typo_result

            error_detection_message = ""
            if image_acquisition_failed:
                error_detection_message = "画像読み込み失敗あり"
            elif ocr_failed_for_existing_image:
                error_detection_message = "テキスト検出失敗あり"

            is_error = (
                "差分あり" in text_comparison_result or
                "OK！" not in final_typo_result or 
                "要確認" in comparison_result or
                error_detection_message != "" 
            )
            status = "要確認" if is_error else "異常なし"

            status_color = "red" if status == "要確認" else "blue"
            row_data_display["ステータス"] = f'<span style="color: {status_color};">{status}</span>'
            row_data_excel["ステータス"] = status

            for portal_name in all_portal_names:
                img_bytes_data = image_bytes.get(portal_name)
                extracted_text = ocr_results.get(portal_name)
                extracted_volume = volume_results.get(portal_name)
                file_id = image_groups.get(image_name, {}).get('portals', {}).get(portal_name, {}).get('id')

                img_col_name = f"{portal_name}（画像）"
                ocr_col_name = f"{portal_name}（OCR）"
                vol_col_name = f"{portal_name}（内容量）"

                if file_id:
                    #correct_image_url = f"[https://drive.google.com/file/d/](https://drive.google.com/file/d/){file_id}/view" # ← 【注意！！】画像のリンクがおかしくなるので使用しない
                    correct_image_url = f"https://drive.google.com/file/d/{file_id}/view"
                    cleaned_volume = extracted_volume.strip().strip('"') if extracted_volume else ""
                    
                    row_data_display[img_col_name] = f'<a href="{correct_image_url}" target="_blank"><img src="data:image/png;base64,{base64.b64encode(img_bytes_data).decode()}" style="max-height: 100px; display: block; margin: auto;"></a>' if img_bytes_data else ""
                    row_data_display[ocr_col_name] = str(extracted_text).replace('\n', '<br>') if extracted_text else ""
                    row_data_display[vol_col_name] = cleaned_volume.replace('\n', '<br>')
                    
                    row_data_excel[img_col_name] = correct_image_url
                    row_data_excel[ocr_col_name] = extracted_text if extracted_text else ""
                    row_data_excel[vol_col_name] = cleaned_volume
                else:
                    row_data_display[img_col_name], row_data_display[ocr_col_name], row_data_display[vol_col_name] = None, None, None
                    row_data_excel[img_col_name], row_data_excel[ocr_col_name], row_data_excel[vol_col_name] = None, None, None

            cleaned_neng_content = neng_content.strip().strip('"') if neng_content else ""
            row_data_display["NENG内容量"] = cleaned_neng_content.replace('\n', '<br>')
            row_data_excel["NENG内容量"] = cleaned_neng_content

            if text_comparison_result == "OK！":
                row_data_display["テキスト比較"] = f'<span style="color: blue;">{text_comparison_result}</span>'
            elif text_comparison_result == "差分あり":
                row_data_display["テキスト比較"] = f'<span style="color: red;">{text_comparison_result}</span>'
            else:
                row_data_display["テキスト比較"] = f'<span style="color: gray;">{text_comparison_result}</span>'
            row_data_excel["テキスト比較"] = text_comparison_result

            if final_typo_result == "OK！":
                row_data_display["誤字脱字"] = f'<span style="color: blue;">{final_typo_result}</span>'
            else:
                display_typo_text = final_typo_result.replace('\n', '<br>')
                row_data_display["誤字脱字"] = f'<span style="color: red;">{display_typo_text}</span>'
            row_data_excel["誤字脱字"] = final_typo_result

            if comparison_result == "OK！":
                row_data_display["内容量比較"] = f'<span style="color: blue;">{comparison_result}</span>'
            elif comparison_result == "要確認":
                row_data_display["内容量比較"] = f'<span style="color: red;">{comparison_result}</span>'
            elif comparison_result == "内容量記載なし":
                row_data_display["内容量比較"] = f'<span style="color: gray;">{comparison_result}</span>'
            else:
                row_data_display["内容量比較"] = comparison_result
            row_data_excel["内容量比較"] = comparison_result

            if error_detection_message:
                row_data_display["エラー検出"] = f'<span style="color: red;">{error_detection_message}</span>'
            else:
                row_data_display["エラー検出"] = ""
            row_data_excel["エラー検出"] = error_detection_message

            results_list_display.append(row_data_display)
            results_list_excel.append(row_data_excel)

        df_display = pd.DataFrame(results_list_display).sort_values(by="画像名").reindex(columns=ordered_columns)
        df_excel = pd.DataFrame(results_list_excel).sort_values(by="画像名").reindex(columns=ordered_columns)

        df_display = df_display.reset_index(drop=True)
        df_display.insert(0, "No", df_display.index + 1)

        df_excel = df_excel.reset_index(drop=True)
        df_excel.insert(0, "No", df_excel.index + 1)

        df_display, df_excel = df_display.fillna(''), df_excel.fillna('')

        df_plain_text_for_search = df_display.map(
            lambda x: re.sub('<[^<]+?>', '', str(x)) if isinstance(x, str) else x
        )

        # --- all_image_bytes_data と df_excel も返す (スプレッドシート保存用) ---
        return df_display, df_plain_text_for_search, df_excel, all_image_bytes_data

    # --- Streamlit UI ---
    col1, col2 = st.columns([4, 1.5]) 
    with col1:
        image_tag = ""
        if 'img_src' in locals() and img_src: 
            image_tag = f'<img src="{img_src}" style="vertical-align: middle; height: 70px;">'
        
        st.markdown(f"""
            <h2 style='font-size: 35px; font-weight: 600; margin-top: 5px; margin-bottom: 0px;'>
                {image_tag}
                商品画像OCR
            </h2>
        """, unsafe_allow_html=True)

        # 操作マニュアルボタンの配置
        if st.button("📖 操作マニュアル", type="tertiary"):
            show_instructions()

    with col2:
        st.markdown(f"<div style='text-align: center; margin-bottom: 1px;'>ようこそ！ <strong>{st.user.name}</strong> さん</div>", unsafe_allow_html=True) 
        if st.button("ログアウト", icon=":material/logout:", width='stretch', key="logout_button"): 
            st.logout() 

    # --- サイドバー ---
    with st.sidebar:
        with st.container(border=True):
            st.header("1. Google Drive 設定")
            st.info("""以下のサービスアカウントを「**閲覧者**」以上の権限で、対象のGoogleドライブフォルダに共有してください。フォルダは**マイドライブ**に配置されている必要があります。

共有ドライブからマイドライブへのコピーは[**こちら**](https://script.google.com/a/macros/steamship.co.jp/s/AKfycbxfHwqxcl-tAnUOsx8OT9lHa2c4DV13VivUMt6CFC6dtDRweQOI53RT5mAP-5VVB5KF/exec)をご利用ください。""")
            st.code("ocr-app@ai-project-427106.iam.gserviceaccount.com", language=None) 
            drive_url_input = st.text_input(
                "読み込み元のGoogleドライブフォルダURL",
                key="drive_url_input_key",
                help="""
【フォルダ構成・画像ファイル名ルール】\n
URL指定用フォルダ ＞ ポータル名等が付いたフォルダ（複数OK） ＞ 画像ファイル（複数OK）\n
※画像ファイル名：品番.jpg
　例） aedg001.jpg、 aedg001-1.jpg
"""
            )

            st.button(
                "ドライブから読み込み",
                width='stretch',
                on_click=_trigger_drive_load_check
            )

            if st.session_state.show_drive_clear_confirmation:
                scroll_sidebar_to_bottom()
                st.warning("ドライブを再読み込みすると、現在の実行結果はクリアされます。よろしいですか？")

                c1, c2 = st.columns(2)
                with c1:
                    st.button("OK", width='stretch', on_click=_execute_drive_load, key="drive_load_ok")
                with c2:
                    st.button("キャンセル", width='stretch', on_click=_cancel_drive_load, key="drive_load_cancel")

            if st.session_state.pop('execute_drive_load_now', False):
                _execute_drive_load()

        with st.container(border=True):
            st.header("2. 実行設定")

            with st.expander("自治体リストの参照元"):
                st.markdown("[こちらのスプレッドシートのデータを参照しています。](https://docs.google.com/spreadsheets/d/1n8qDS8OvuFJwDy2J6wduDHI32GxDmbx1QIrqHPFjdGo/)", unsafe_allow_html=True)

            municipality_map = st.session_state.get("municipality_map", {})

            if isinstance(municipality_map, list):
                st.sidebar.warning("自治体リストの形式が古いです。データを再読み込みします。")
                with st.sidebar, st.spinner("自治体リストを更新中..."):
                    if 'sheets_service' not in locals():
                        google_creds_check, _ = get_google_credentials()
                        sheets_service = get_sheets_service(google_creds_check)

                    st.session_state.municipality_map = get_municipality_map(sheets_service)
                    st.cache_data.clear() 
                st.rerun()

            municipality_options = list(municipality_map.keys()) 

            current_municipality_selection = st.session_state.get("municipality_select_key")
            default_municipality = None
            if st.session_state.old_municipality in municipality_options:
                default_municipality = st.session_state.old_municipality
            if current_municipality_selection is None or current_municipality_selection not in municipality_options:
                st.session_state.municipality_select_key = default_municipality

            selected_municipality_name = st.selectbox(
                "自治体を選択",
                options=municipality_options, 
                disabled=not municipality_options, 
                help="NENG APIで内容量を取得する為に自治体の設定が必要となります。",
                key="municipality_select_key",
                placeholder="選択してください", 
                on_change=_check_for_change_and_warn,
                args=("municipality_select_key", "old_municipality")
            )
            if st.session_state.old_municipality is None and st.session_state.municipality_select_key is not None:
                st.session_state.old_municipality = st.session_state.municipality_select_key

            business_code_options = st.session_state.business_codes
            current_business_selection = st.session_state.get("business_select_key")
            default_business_code = None
            if st.session_state.old_business_code in business_code_options:
                default_business_code = st.session_state.old_business_code
            elif business_code_options: 
                default_business_code = business_code_options[0]
            if current_business_selection is None or current_business_selection not in business_code_options:
                if default_business_code is not None: 
                    st.session_state.business_select_key = default_business_code

            col1, col2 = st.columns(2)
            with col1:
                selected_business_code = st.selectbox(
                    "事業者コードを選択",
                    options=business_code_options,
                    disabled=not business_code_options,
                    key="business_select_key",
                    on_change=_check_for_change_and_warn,
                    args=("business_select_key", "old_business_code")
                )
                if st.session_state.old_business_code is None and business_code_options:
                    st.session_state.old_business_code = st.session_state.business_select_key

            if selected_business_code: 
                st.session_state.product_codes = get_product_codes_for_business_code(st.session_state.portal_files, selected_business_code)
            else:
                st.session_state.product_codes = []

            product_code_options = ["すべて"] + st.session_state.product_codes
            current_product_selection = st.session_state.get("product_select_key")
            default_product_code = "すべて"
            if st.session_state.old_product_code in product_code_options:
                default_product_code = st.session_state.old_product_code
            if current_product_selection is None or current_product_selection not in product_code_options:
                st.session_state.product_select_key = default_product_code

            with col2:
                selected_product_code = st.selectbox(
                    "品番を選択",
                    options=product_code_options,
                    disabled=not selected_business_code, 
                    key="product_select_key",
                    on_change=_check_for_change_and_warn,
                    args=("product_select_key", "old_product_code")
                )
                if st.session_state.old_product_code is None:
                    st.session_state.old_product_code = st.session_state.product_select_key

            if st.session_state.show_clear_confirmation:
                scroll_sidebar_to_bottom()
                st.warning("設定を変更すると、現在の実行結果はクリアされます。よろしいですか？")

                c1_confirm, c2_confirm = st.columns(2)
                with c1_confirm:
                    st.button("OK", width='stretch', on_click=_confirm_clear_results, key="clear_results_ok")
                with c2_confirm:
                    st.button("キャンセル", width='stretch', on_click=_cancel_clear_results, key="clear_results_cancel")


        with st.container(border=True):
            st.header("3. OCR実行")
            run_disabled = not (selected_business_code and selected_municipality_name is not None) \
                            or st.session_state.show_clear_confirmation \
                            or st.session_state.show_drive_clear_confirmation \
                            or not st.session_state.portal_files 

            if st.button("OCR実行", type="primary", width='stretch', disabled=run_disabled):
                st.session_state.old_municipality = selected_municipality_name
                st.session_state.old_business_code = selected_business_code
                st.session_state.old_product_code = selected_product_code

                st.session_state.current_page = 1 
                record_count, total_images = count_images_to_process(st.session_state.portal_files, selected_business_code, selected_product_code)
                st.session_state.record_count_to_process = record_count
                st.session_state.image_total_count_to_process = total_images
                if record_count > 0:
                    st.session_state.show_ocr_confirmation = True
                else:
                    st.warning("処理対象の画像が見つかりませんでした。事業者コード・品番を確認してください。")
                st.rerun()

            if st.session_state.get("show_ocr_confirmation"):
                scroll_sidebar_to_bottom()
                record_count, total_images = st.session_state.record_count_to_process, st.session_state.image_total_count_to_process
                IMAGE_LIMIT = 500
                if total_images > IMAGE_LIMIT:
                    st.error(f"画像枚数の合計が多すぎます ({IMAGE_LIMIT}枚まで)。現在: {total_images}枚")
                    if st.button("閉じる", width='stretch'):
                        st.session_state.show_ocr_confirmation = False
                        st.rerun()
                else:
                    st.info(f"**{record_count}件**（画像 全{total_images}枚）の処理を開始します。")
                    c1_ocr, c2_ocr = st.columns(2)
                    if c1_ocr.button("OK", width='stretch', key="ocr_exec_ok"):
                        st.session_state.show_ocr_confirmation = False
                        
                        # --- 実行前に前回の結果をクリアする ---
                        st.session_state.ocr_result_df = None
                        st.session_state.ocr_plain_df = None
                        # [削除] ocr_excel_output のクリアを削除
                        st.session_state.ocr_excel_df = None 
                        st.session_state.ocr_image_bytes = None
                        st.session_state.current_page = 1

                        municipality_code = None
                        try:
                            municipality_map = st.session_state.get("municipality_map", {})
                            municipality_code = municipality_map.get(selected_municipality_name)

                            if not municipality_code:
                                st.error("選択された自治体のコードが見つかりません。")
                        except Exception as e:
                            st.error(f"自治体コード取得中にエラー: {e}")


                        if municipality_code:
                            with st.spinner("OCR処理を実行中です..."):
                                progress_bar = st.progress(0, text="準備中...")
                                try:
                                    # --- 戻り値に image_bytes_data と df_excel を追加 ---
                                    df, df_plain, df_excel, image_bytes_data = run_ocr_process(
                                        st.session_state.portal_files,
                                        municipality_code,
                                        selected_business_code,
                                        selected_product_code,
                                        google_creds,
                                        async_openai_client,
                                        progress_bar
                                    )
                                    if df is not None: 
                                        
                                        # [削除] Excelファイル生成処理を削除
                                        
                                        st.session_state.ocr_result_df = df
                                        st.session_state.ocr_plain_df = df_plain
                                        st.session_state.ocr_excel_df = df_excel 
                                        # --- 画像バイナリデータをセッションに保存 ---
                                        st.session_state.ocr_image_bytes = image_bytes_data
                                        st.session_state.show_success_message = True
                                except Exception as e:
                                    st.error(f"OCR処理の実行中にエラーが発生しました: {e}")
                                    if 'progress_bar' in locals():
                                        progress_bar.empty()

                        st.rerun() 

                    if c2_ocr.button("キャンセル", width='stretch', key="ocr_exec_cancel"):
                        st.session_state.show_ocr_confirmation = False
                        st.rerun()

    if st.session_state.pop("show_success_message", False):
        st.toast("処理が完了しました。", icon="🎉")
    
    # --- 結果表示エリア ---
    if 'ocr_result_df' in st.session_state and st.session_state.ocr_result_df is not None:
        df_display_source = st.session_state.ocr_result_df 
        total_count = len(df_display_source)

        # --- [移動・変更] スプレッドシート保存エリア (開閉式) ---
        
        # 保存ボタン表示条件
        show_gspread_button = 'ocr_excel_df' in st.session_state and \
                              st.session_state.ocr_excel_df is not None and \
                              not st.session_state.ocr_excel_df.empty

        # --- _execute_gspread_save コールバック関数 (ここに移動) ---
        def _execute_gspread_save():
            # 実行前に以前のメッセージをリセット
            st.session_state.gspread_save_success_url = None
            st.session_state.gspread_save_error_message = None

            url = st.session_state.gspread_sheet_url_input 
            if not url:
                st.session_state.gspread_save_error_message = "スプレッドシートURLを入力してください。" 
                return

            spreadsheet_id = get_spreadsheet_id_from_url(url) 
            if not spreadsheet_id:
                st.session_state.gspread_save_error_message = "有効なGoogleスプレッドシートのURLではありません。`/d/.../` を含むURLを入力してください。" 
                return
            
            try:
                image_bytes_data = st.session_state.get("ocr_image_bytes", {})

                # ユーザーに処理中であることを視覚的に伝える
                with st.spinner("スプレッドシートに保存中..."):
                    save_to_spreadsheet(
                        st.session_state.ocr_excel_df, 
                        spreadsheet_id, 
                        sheet_name,  
                        google_creds_info, 
                        st.session_state.portal_files,
                        image_bytes_data
                    )
                
                #  GID（シートID）を取得してURLを生成
                with st.spinner("シートURLを取得中..."):
                    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
                    sheets = sheet_metadata.get('sheets', [])
                    gid = None
                    for s in sheets:
                        if s.get('properties', {}).get('title') == sheet_name:
                            gid = s.get('properties', {}).get('sheetId')
                            break
                
                # 成功時
                st.session_state.gspread_sheet_url_input = "" # 入力欄をクリア
                
                base_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/"
                if gid is not None:
                    st.session_state.gspread_save_success_url = f"{base_url}edit#gid={gid}" # GID付きURL
                else:
                    st.session_state.gspread_save_success_url = base_url # GIDが見つからなかった場合
                
                st.toast(f"シート「{sheet_name}」に保存しました！", icon="✅") 

            except HttpError as e: # HttpErrorをキャッチ
                st.session_state.gspread_save_error_message = f"スプレッドシート処理中にエラーが発生しました: {e}"
            except NameError as e:
                st.session_state.gspread_save_error_message = "サービスアカウントの認証情報(google_creds_info)の読み込みに失敗しました。" 
            except Exception as e:
                st.session_state.gspread_save_error_message = str(e)

        # --- 保存用変数定義（シート名など） ---
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        municipality_name = st.session_state.old_municipality if st.session_state.old_municipality else "unknown"
        municipality_name_safe = re.sub(r'[\\/*?:"<>|]', '_', municipality_name) 
        product_part = st.session_state.old_product_code if st.session_state.old_product_code != "すべて" else "all"
        business_part = st.session_state.old_business_code if st.session_state.old_business_code else "unknown"
        sheet_name = f"{municipality_name_safe}_{business_part}_{product_part}_{today_str}"

        # --- UI表示 (Expander) ---
        if show_gspread_button:
            # 入力フォームとボタンは Expander の中
            with st.expander("スプレッドシートに保存", expanded=False):
                st.info(f"[**こちらのスプレッドシート**](https://docs.google.com/spreadsheets/d/1Hi4TYK16lsezrp2Hnv6ICQQzLPcb_xhkneOEzXMk9Rc)を**マイドライブ**にコピーして、サイドメニューのサービスアカウントを「**編集者**」権限で共有してください。コピーしたスプレッドシートのURLを以下に入力して「保存」ボタンを押下してください。")
                
                st.text_input(
                    "GoogleスプレッドシートURL", 
                    key="gspread_sheet_url_input", 
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                )
                
                st.button("保存", key="gspread_create_button", type="primary", width='stretch', on_click=_execute_gspread_save)

            # --- メッセージ表示エリア (Expanderの外・すぐ下) ---
            if st.session_state.gspread_save_error_message:
                st.error(st.session_state.gspread_save_error_message, icon="🚨")

            if st.session_state.gspread_save_success_url:
                success_url = st.session_state.gspread_save_success_url
                st.success(f"スプレッドシートに保存しました: [開く]({success_url})", icon="📄")

        # ---------------------------------------------------------

        if not df_display_source.empty:
            product_codes_in_result = sorted(list(
                df_display_source['画像名'].apply(get_product_code_from_filename).unique()
            ))
            product_filter_options = ["すべて"] + product_codes_in_result
        else:
            product_filter_options = ["すべて"]

        # --- フィルターUI ---
        filter_col, view_col = st.columns(2)
        with view_col:
            # --- [変更] UI部分 ---
            with st.container(border=True):
                st.markdown("##### 表示列の設定")
                # カラムを3つにする
                cc1, cc2, cc3 = st.columns(3)
                
                with cc1:
                    # [追加] テキスト比較のチェックボックス
                    show_text_compare = st.checkbox("テキスト比較", value=True, help="テキスト比較列を表示")
                with cc2:
                    # [変更] ヘルプテキストから「テキスト比較」を削除
                    show_ocr_cols = st.checkbox("誤字脱字", value=True, help="OCR/誤字脱字列を表示")
                with cc3:
                    show_content_cols = st.checkbox("内容量", value=True, help="内容量/NENG/内容量比較列を表示")

                with st.expander("表示ポータルの絞り込み"):
                    portal_names = list(st.session_state.portal_files.keys()) if st.session_state.portal_files else []
                    selected_portals = st.multiselect(
                        "表示するポータルを選択してください",
                        options=portal_names,
                        default=portal_names, 
                        label_visibility="collapsed",
                        disabled=not portal_names 
                    )
        with filter_col:
            with st.container(border=True):
                st.markdown("##### フィルター")
                fc1, fc2, fc3 = st.columns([2, 2, 1])
                with fc1:
                    search_term = st.text_input("全文検索", placeholder="表全体からキーワードで検索...")
                with fc2:
                    selected_product_filter = st.selectbox(
                        "品番",
                        options=product_filter_options,
                        key="product_filter_selectbox", 
                        disabled=(len(product_filter_options) <= 1)
                    )
                with fc3:
                    status_filter = st.selectbox("ステータス", ["すべて", "要確認", "異常なし"])
                st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

        # --- テーブル表示 ---
        # --- [変更] テーブル表示用ハイライト関数 ---
        # 引数に text_compare_visible を追加
        def highlight_row(row, ocr_visible, content_visible, text_compare_visible):
            style = ''
            has_visible_error = False

            if "エラー検出" in row and "失敗あり" in str(row["エラー検出"]):
                has_visible_error = True

            # [追加] テキスト比較が表示されている場合のみチェック
            if text_compare_visible and not has_visible_error:
                if "テキスト比較" in row and "差分あり" in str(row["テキスト比較"]):
                    has_visible_error = True

            # [変更] ocr_visible のブロックからテキスト比較の判定を削除
            if ocr_visible and not has_visible_error:
                # ここから「テキスト比較」の判定を削除しました
                if "誤字脱字" in row and '<span style="color: blue;">OK！</span>' not in str(row["誤字脱字"]):
                    has_visible_error = True

            if content_visible and not has_visible_error:
                if "内容量比較" in row and "要確認" in str(row["内容量比較"]):
                    has_visible_error = True

            if ocr_visible and not has_visible_error:
                for col_name in row.index:
                    if col_name.endswith('（OCR）'):
                        cell_content = str(row[col_name])
                        if "APIエラー" in cell_content or "AI OCRエラー" in cell_content or "画像取得失敗" in cell_content:
                            has_visible_error = True
                            break 

            if has_visible_error:
                style = 'background-color: #ffe5e5' 

            return [style] * len(row)

        # --- データフィルタリング ---
        df_to_process = df_display_source.copy() 

        # [削除] 下部にあったファイル名生成などは上部に移動済み
        # [削除] 下部にあったshow_gspread_buttonブロックは上部に移動済み

        col_header_left, col_header_right = st.columns([3, 2])
        
        # [削除] コールバック関数 _execute_gspread_save は上部に移動済み

        if not selected_portals:
            with col_header_left:
                st.markdown(f"<h2 style='font-size: 20px; font-weight: 600; margin-bottom: 0px;'>実行結果 0 / {total_count}件</h2>", unsafe_allow_html=True)
            with col_header_right:
                pass
            st.info("「表示ポータルの絞り込み」で表示するポータルを1つ以上選択してください。")
        else:
            # 1. 品番フィルター
            if selected_product_filter != "すべて":
                mask_product = df_to_process['画像名'].apply(get_product_code_from_filename) == selected_product_filter
                df_to_process = df_to_process[mask_product]

            # 2. ステータスフィルター
            if status_filter != "すべて":
                df_filtered_before_status = df_to_process.copy()
                all_columns = df_display_source.columns
                visible_columns = []
                for col in all_columns:
                    if col in ["No", "画像名", "ステータス", "エラー検出"]:
                        visible_columns.append(col)
                        continue
                    
                    # [追加] テキスト比較の制御
                    if col == "テキスト比較" and show_text_compare:
                        visible_columns.append(col)
                        continue
                    
                    # [変更] 誤字脱字のみ show_ocr_cols で制御
                    if col == "誤字脱字" and show_ocr_cols:
                        visible_columns.append(col)
                        continue
                    
                    if col in ["NENG内容量", "内容量比較"] and show_content_cols:
                        visible_columns.append(col)
                        continue
                    is_portal_col = False
                    for portal in selected_portals:
                        if portal in col:
                            is_portal_col = True
                            break
                    if is_portal_col:
                        if '（画像）' in col:
                            visible_columns.append(col)
                        elif '（OCR）' in col and show_ocr_cols:
                            visible_columns.append(col)
                        elif '（内容量）' in col and show_content_cols:
                            visible_columns.append(col)

                df_visible_cols_only = df_filtered_before_status[visible_columns]
                is_row_highlighted = df_visible_cols_only.apply(
                    # [変更] 引数に text_compare_visible=show_text_compare を追加
                    lambda row: 'background-color: #ffe5e5' in highlight_row(row, show_ocr_cols, show_content_cols, show_text_compare)[0],
                    axis=1
                )

                if status_filter == "要確認":
                    df_to_process = df_filtered_before_status[is_row_highlighted]
                elif status_filter == "異常なし":
                    df_to_process = df_filtered_before_status[~is_row_highlighted] 

            # 3. 全文検索フィルター
            if search_term:
                df_plain_text_filtered = st.session_state.ocr_plain_df.loc[df_to_process.index]
                mask_search = df_plain_text_filtered.apply(
                    lambda row: row.astype(str).str.contains(search_term, case=False, na=False).any(),
                    axis=1
                )
                df_to_process = df_to_process[mask_search]

            # --- フィルタリング結果表示 ---
            filtered_count = len(df_to_process)
            is_filtered = (status_filter != "すべて") or (search_term != "") or (selected_product_filter != "すべて")

            col_header_left, col_header_right = st.columns([3, 2], vertical_alignment="bottom")

            with col_header_left:
                # フィルター適用後のデータフレームから「要確認」が含まれる行数をカウント
                need_check_count = df_to_process['ステータス'].astype(str).str.contains("要確認").sum()
                
                # 件数が1件以上ある場合のみ表示
                if need_check_count > 0:
                    # 全体を少し小さく(0.8em)し、カッコは黒字、中の文字だけ赤色にする
                    check_status_html = f"<span style='font-size: 0.8em;'>（<span style='color: red;'>要確認 {need_check_count}件</span>）</span>"
                else:
                    check_status_html = ""

                if is_filtered and filtered_count != total_count:
                    st.markdown(f"<h2 style='font-size: 20px; font-weight: 600; margin-bottom: 0px;'>実行結果 {filtered_count} / {total_count}件 {check_status_html}</h2>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<h2 style='font-size: 20px; font-weight: 600; margin-bottom: 0px;'>実行結果 {total_count}件 {check_status_html}</h2>", unsafe_allow_html=True)

            is_zoom_mode = False # 初期化
            with col_header_right:
                # カラム比率を変更して右端に寄せる
                _, toggle_col = st.columns([1, 0.32]) 
                with toggle_col:
                    # 「拡大表示」をONにするスイッチ
                    is_zoom_mode = st.toggle("拡大表示", value=False, key="view_mode_toggle")

            all_columns = df_display_source.columns 
            final_columns_to_show = []
            for col in all_columns:
                if col in ["No", "画像名", "ステータス", "エラー検出"]:
                    final_columns_to_show.append(col)
                    continue
                
                # [追加] テキスト比較の独立制御
                if col == "テキスト比較" and show_text_compare:
                    final_columns_to_show.append(col)
                    continue
                
                # [変更] 誤字脱字のみ show_ocr_cols で制御
                if col == "誤字脱字" and show_ocr_cols:
                    final_columns_to_show.append(col)
                    continue
                    
                if col in ["NENG内容量", "内容量比較"] and show_content_cols:
                    final_columns_to_show.append(col)
                    continue

                is_portal_col = False
                for portal in selected_portals:
                    if portal in col:
                        is_portal_col = True
                        break
                if is_portal_col:
                    if '（画像）' in col:
                        final_columns_to_show.append(col)
                    elif '（OCR）' in col and show_ocr_cols:
                        final_columns_to_show.append(col)
                    elif '（内容量）' in col and show_content_cols:
                        final_columns_to_show.append(col)

            df_filtered_display = df_to_process[final_columns_to_show] 

            ITEMS_PER_PAGE = 20
            total_pages = math.ceil(filtered_count / ITEMS_PER_PAGE) if filtered_count > 0 else 1

            if st.session_state.current_page > total_pages:
                st.session_state.current_page = 1

            start_idx = (st.session_state.current_page - 1) * ITEMS_PER_PAGE
            end_idx = start_idx + ITEMS_PER_PAGE
            df_paginated = df_filtered_display.iloc[start_idx:end_idx]

            if not df_paginated.empty:
                from functools import partial 
                # [変更] text_compare_visible を追加
                highlight_func = partial(highlight_row, ocr_visible=show_ocr_cols, content_visible=show_content_cols, text_compare_visible=show_text_compare)
                styler = df_paginated.style.apply(highlight_func, axis=1)

                styler.hide(axis="index") 

                styler.hide(axis="columns", subset=["ステータス"])

                html_table = styler.to_html(escape=False, table_attributes='class="custom_df"')

                # 基本クラス
                container_classes = ["table-container"]
                
                # 1. 画像モード判定（チェックボックスの状態）
                # [変更] show_text_compare も条件に追加（すべてOFFの場合などレイアウト崩れ防止のため）
                if not (show_ocr_cols and show_content_cols and show_text_compare):
                    container_classes.append("image-mode-only")
                
                # 2. 全体表示モード判定（トグルスイッチの状態）
                # トグルがOFF (False) の場合、「全体表示 (fit-mode)」を適用します
                if not is_zoom_mode:
                    container_classes.append("fit-mode")
                
                # クラスを結合
                final_class = " ".join(container_classes)

                st.markdown(f'<div class="{final_class}">{html_table}</div>', unsafe_allow_html=True)
            else:
                st.info("フィルター条件に一致する結果がありません。")

            if total_pages > 1:
                st.write("") 
                p1, p2, p3, p4, p5 = st.columns([1, 1, 2, 1, 1]) 

                is_disabled_prev = (st.session_state.current_page <= 1)
                is_disabled_next = (st.session_state.current_page >= total_pages)

                if p1.button("＜＜", width='stretch', disabled=is_disabled_prev, key="page_first"):
                    st.session_state.current_page = 1
                    st.rerun()
                if p2.button("＜", width='stretch', disabled=is_disabled_prev, key="page_prev"):
                    st.session_state.current_page -= 1
                    st.rerun()
                p3.markdown(f"<div style='text-align: center; margin-top: 5px;'>{st.session_state.current_page} / {total_pages} ページ</div>", unsafe_allow_html=True)
                if p4.button("＞", width='stretch', disabled=is_disabled_next, key="page_next"):
                    st.session_state.current_page += 1
                    st.rerun()
                if p5.button("＞＞", width='stretch', disabled=is_disabled_next, key="page_last"):
                    st.session_state.current_page = total_pages
                    st.rerun()

        # [削除] ここにあったスプレッドシート保存処理は上部に移動済み

    # --- OCR結果がまだない場合の表示 ---
    else:
        col_header_left_c, col_header_right_c = st.columns([3, 2]) 

        with col_header_left_c:
            st.markdown("<h2 style='font-size: 20px; font-weight: 600; margin-bottom: 0px;'>実行結果 0件</h2>", unsafe_allow_html=True)

        with col_header_right_c:
            pass 

        st.info("サイドバーで設定を行い、「OCR実行」ボタンを押してください。")