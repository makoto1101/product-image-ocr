import aiohttp
import asyncio
import streamlit as st

async def get_neng_content(product_code: str, municipality_code: str) -> str:
    """
    品番と自治体コードを基にNENG APIから「内容量・規格等」を非同期で取得する。

    Args:
        product_code (str): 返礼品の品番 (SKU)。
        municipality_code (str): 自治体コード。

    Returns:
        str: 取得した「内容量・規格等」のテキスト。エラーや該当なしの場合は空文字を返す。
    """
    if not product_code or not municipality_code:
        return ""

    sku = product_code.upper()

    try:
        user = st.secrets["NENG"]["NENG_USER"]
        password = st.secrets["NENG"]["NENG_PASSWORD"]
    except KeyError as e:
        st.error(f"NENG APIの認証情報がsecrets.tomlに設定されていません: {e}")
        return "NENG認証情報エラー"
    
    url = f"https://n2.steamship.co.jp/{municipality_code}/wp-admin/admin-ajax.php?action=n2_items_api&mode=json&code={sku}"
    auth = aiohttp.BasicAuth(login=user, password=password)

    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, auth=auth, timeout=10) as response:
                if response.status == 200:
                    try:
                        json_response = await response.json(content_type=None) # content-typeを無視
                        
                        if json_response and "items" in json_response:
                            items = json_response["items"]
                            # itemsがリストでも単一の辞書でも対応
                            item = items[0] if isinstance(items, list) and items else (items if isinstance(items, dict) else None)
                            
                            if item and isinstance(item, dict):
                                content = item.get("内容量・規格等", "")
                                return content if content is not None else ""
                        
                        # itemsがない、または空の場合
                        return ""
                        
                    except aiohttp.ContentTypeError:
                        # JSONでない場合(HTMLエラーページなど)
                        return ""
                else:
                    # HTTPステータスコードが200以外の場合
                    return ""
                    
    except asyncio.TimeoutError:
        return "タイムアウトエラー"
    except aiohttp.ClientError:
        # 接続エラーなど
        return "API接続エラー"
    except Exception:
        return "予期せぬエラー"
