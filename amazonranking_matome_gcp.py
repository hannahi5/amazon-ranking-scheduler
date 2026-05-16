# -*- coding: utf-8 -*-
"""
Amazon ランキング取得（requests + BeautifulSoup方式）
- Playwrightを使わずrequestsで軽量にHTML取得
- BeautifulSoupでDOM解析してランキングを抽出
- 結果はGoogle スプレッドシートに書き込み
"""
import os
import re
import time
import random
import datetime
import requests
from bs4 import BeautifulSoup

# =========================
# 設定
# =========================

PAPER_URL = "https://www.amazon.co.jp/dp/4798183180"
AUDIBLE_URL = "https://www.amazon.co.jp/dp/B0G66DNXDH"

CREDENTIALS_PATH = os.path.expanduser("~/credentials.json")
SPREADSHEET_ID = "1DSn3IK9ebd0apbqe2WIXKaRGrDVg7XhaK1jlQZrjBk8"
SHEET_NAME = "Amazon 売れ筋ランキング"

# プロキシ設定（IPRoyal Residential）
PROXY_HOST = "geo.iproyal.com"
PROXY_PORT = "12321"
PROXY_USER = "VjolpstNX9HENvOY"
PROXY_PASS = "PCiV3IQ6N2iDLQnX"
PROXY_COUNTRY = "jp"

PROXIES = {
    "http": f"http://{PROXY_USER}:{PROXY_PASS}_country-{PROXY_COUNTRY}@{PROXY_HOST}:{PROXY_PORT}",
    "https": f"http://{PROXY_USER}:{PROXY_PASS}_country-{PROXY_COUNTRY}@{PROXY_HOST}:{PROXY_PORT}",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Cache-Control": "max-age=0",
}

LOG_PATH = os.path.expanduser("~/amazonranking_log.txt")

def log(msg):
    ts = time.strftime("%Y/%m/%d %H:%M:%S")
    log_msg = f"[{ts}] {msg}"
    print(log_msg)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(log_msg + "\n")
    except Exception:
        pass

# =========================
# ランキング取得
# =========================

def create_session():
    """クッキーを保持するセッションを作成し、まずトップページにアクセスする"""
    session = requests.Session()
    session.headers.update(HEADERS)
    session.proxies.update(PROXIES)
    try:
        session.get("https://www.amazon.co.jp/", timeout=30)
        log("セッション作成完了（トップページアクセス済み）")
    except Exception as e:
        log(f"セッション作成時の警告: {e}")
    return session

def fetch_rankings(session, url, keyword):
    """requestsでHTMLを取得し、BeautifulSoupでランキングを抽出"""
    # ランダムな遅延（3〜8秒）
    delay = random.uniform(3, 8)
    log(f"{keyword} {delay:.1f}秒待機後にアクセス")
    time.sleep(delay)

    log(f"{keyword} ページ取得開始")
    try:
        r = session.get(url, timeout=30)
        html = r.text
        log(f"{keyword} ページ取得完了（HTMLサイズ: {len(html)} bytes）")
    except Exception as e:
        log(f"{keyword} 取得エラー: {e}")
        return ["-"] * 4

    if "売れ筋ランキング" not in html:
        if "captcha" in html.lower():
            log(f"{keyword} CAPTCHAでブロックされました")
        else:
            log(f"{keyword} ランキング情報が見つかりません")
        return ["-"] * 4

    soup = BeautifulSoup(html, "html.parser")
    rankings = []

    # 方法A: detailBullets から全体ランキングを取得（紙書籍向け）
    bullets = soup.select("#detailBulletsWrapper_feature_div li span.a-list-item")
    for b in bullets:
        text = b.get_text()
        if "売れ筋ランキング" in text:
            match = re.search(r'[：:]\s*(.+?)\s+-\s+(\d{1,3}(?:,\d{3})*位)', text)
            if match:
                rankings.append(f"{match.group(2).strip()} {match.group(1).strip()}")
            break

    # 方法B: zg_hrsr からカテゴリ別ランキングを取得（紙書籍向け）
    items = soup.select("ul.zg_hrsr li")
    for item in items:
        text = item.get_text().replace("\n", " ").strip()
        match = re.search(r'(.+?)\s+-\s+(\d{1,3}(?:,\d{3})*位)', text)
        if match:
            category = match.group(1).strip()
            rank = match.group(2).strip()
            rankings.append(f"{rank} {category}")

    # 方法C: ページ全体のテキストから抽出（Audible向けフォールバック）
    if not rankings:
        body_text = soup.get_text(separator="\n")
        idx = body_text.find("売れ筋ランキング")
        if idx != -1:
            block = body_text[idx:idx+500]
            lines = [l.strip() for l in block.split("\n") if l.strip()]
            # 「カテゴリ」と「- 順位」が別行の場合があるので結合
            merged = " ".join(lines)
            # ノイズ除去
            merged = re.sub(r"\(\s*[^)]*売れ筋ランキングを見る\s*\)", "", merged)
            merged = re.sub(r"売れ筋ランキング\s*", "", merged, count=1)
            # 抽出
            for m in re.finditer(r"(.+?)\s+-\s+(\d{1,3}(?:,\d{3})*位)", merged):
                category = m.group(1).strip()
                rank = m.group(2).strip()
                if not category or "関連" in category or "スポンサー" in category:
                    continue
                rankings.append(f"{rank} {category}")
                if len(rankings) >= 4:
                    break

    # 重複除去
    seen = set()
    unique = []
    for r in rankings:
        if r not in seen:
            unique.append(r)
            seen.add(r)

    # 4つに揃える
    while len(unique) < 4:
        unique.append("-")
    unique = unique[:4]

    log(f"{keyword} 抽出完了: {unique}")
    return unique

# =========================
# スプレッドシート書き込み
# =========================

def write_to_spreadsheet(row_data):
    if not os.path.exists(CREDENTIALS_PATH):
        log("認証ファイルが見つかりません（スキップ）")
        return

    try:
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials

        scope = ["https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, scope)
        client = gspread.authorize(creds)
        log("認証成功")

        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        ws = spreadsheet.worksheet(SHEET_NAME)

        # 重複チェック
        existing = ws.get_all_values()
        existing_dates = set()
        for row in existing[1:]:
            if row and row[0]:
                existing_dates.add(row[0])

        if row_data[0] in existing_dates:
            log("この日時のデータは既に存在します（スキップ）")
            return

        ws.append_row(row_data, value_input_option="USER_ENTERED")
        ws.sort((1, "des"))
        log("スプレッドシート書き込み完了")

    except Exception as e:
        log(f"スプレッドシートエラー: {e}")
        import traceback
        log(f"エラー詳細: {traceback.format_exc()}")

# =========================
# 実行
# =========================

def main():
    # 起動時にランダム遅延（0〜300秒）で実行時刻をずらす
    startup_delay = random.uniform(0, 300)
    log(f"起動遅延: {startup_delay:.0f}秒")
    time.sleep(startup_delay)

    log("=" * 60)
    log("処理開始")
    log("=" * 60)

    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y/%m/%d %H:%M")
    log(f"実行日時: {now}")

    session = create_session()
    paper = fetch_rankings(session, PAPER_URL, "紙書籍")
    audible = fetch_rankings(session, AUDIBLE_URL, "Audible")

    row_data = [now] + paper + audible
    log(f"構築データ: {row_data}")

    write_to_spreadsheet(row_data)

    log("=" * 60)
    log("すべての処理が完了しました")
    log("=" * 60)

if __name__ == "__main__":
    main()
