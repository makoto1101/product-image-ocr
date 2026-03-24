import aiohttp
import asyncio
import json
import os
import streamlit as st

RETRYABLE_STATUS_CODES = {408, 425, 429, 500, 502, 503, 504}
MAX_RETRIES = 3
RETRY_BASE_DELAY_SEC = 0.4

async def get_n3_content(product_code: str, municipality_code: str) -> str:
    """
    品番(SKU)と自治体コードを基にN3 Public APIから「内容量・規格等」を非同期で取得する。

    期待する secrets.toml 例:
    [N3]
    N3_BASE_URL = "https://n3.example.com"
    N3_PUBLIC_API_TOKEN = "your-token"

    Args:
        product_code (str): 返礼品の品番 (SKU)
        municipality_code (str): 自治体コード

    Returns:
        str: 「内容量・規格等」の文字列。
             該当なしは空文字。
             接続系エラー時はエラーメッセージ文字列を返す。
    """
    if not product_code or not municipality_code:
        return ""

    sku = product_code.strip().upper()
    city_code = municipality_code.strip()
    if not sku or not city_code:
        return ""

    base_url = ""
    token = ""
    try:
        base_url = st.secrets["N3"]["N3_BASE_URL"].rstrip("/")
        token = st.secrets["N3"]["N3_PUBLIC_API_TOKEN"]
    except Exception:
        # Streamlit outside context or secrets未設定時は環境変数から読む
        base_url = os.getenv("N3_BASE_URL", "").rstrip("/")
        token = os.getenv("N3_PUBLIC_API_TOKEN", "")

    if not base_url or not token:
        return "N3認証情報エラー"

    url = f"{base_url}/api/public/{city_code}/items"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    params = {
        "itemCode": sku,  # N3は基本的に完全一致
        "limit": "1",
        "page": "1",
    }

    timeout = aiohttp.ClientTimeout(total=10)
    try:
        async with aiohttp.ClientSession(timeout=timeout) as session:
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    async with session.get(url, headers=headers, params=params) as response:
                        if response.status == 401:
                            return "認証エラー"
                        if response.status == 404:
                            return ""
                        if response.status in RETRYABLE_STATUS_CODES:
                            if attempt < MAX_RETRIES:
                                await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                                continue
                            return "API一時エラー"
                        if response.status != 200:
                            return ""

                        # Content-Typeが不正でも読めるようにtext->jsonで処理
                        raw_text = await response.text()
                        try:
                            payload = json.loads(raw_text)
                        except json.JSONDecodeError:
                            if attempt < MAX_RETRIES:
                                await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                                continue
                            return ""

                        items = payload.get("items")
                        if isinstance(items, list) and items:
                            item = items[0]
                        elif isinstance(items, dict):
                            item = items
                        elif isinstance(payload.get("item"), dict):
                            item = payload.get("item")
                        else:
                            # 検索インデックスの反映遅れ等を考慮して空ヒットは短く再試行
                            if attempt < MAX_RETRIES:
                                await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                                continue
                            return ""

                        if not isinstance(item, dict):
                            if attempt < MAX_RETRIES:
                                await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                                continue
                            return ""

                        # N3のラベル済みキーを優先し、未ラベルキーにもフォールバック
                        candidate_keys = [
                            "内容量・規格等",
                            "内容量",
                            "specifications",
                            "content",
                            "capacity",
                            "description",
                            "detail",
                            "param",
                            "volumeText",
                        ]
                        for key in candidate_keys:
                            content = item.get(key)
                            if isinstance(content, str) and content.strip():
                                return content.strip()

                        # itemはあるが内容量が空の場合も短く再試行して取りこぼしを減らす
                        if attempt < MAX_RETRIES:
                            await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                            continue
                        return ""

                except asyncio.TimeoutError:
                    if attempt < MAX_RETRIES:
                        await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                        continue
                    return "タイムアウトエラー"
                except aiohttp.ClientError:
                    if attempt < MAX_RETRIES:
                        await asyncio.sleep(RETRY_BASE_DELAY_SEC * attempt)
                        continue
                    return "API接続エラー"

        return ""
    except Exception:
        return "予期せぬエラー"
    

if __name__ == "__main__":
    test_product = os.getenv("N3_TEST_PRODUCT_CODE", "ABH052")
    test_city = os.getenv("N3_TEST_CITY_CODE", "402303")
    result = asyncio.run(get_n3_content(test_product, test_city))
    print(f"product={test_product}, city={test_city}")
    print(f"content={result!r}")