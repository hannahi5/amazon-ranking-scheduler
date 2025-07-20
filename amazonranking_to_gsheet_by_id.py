import gspread
from oauth2client.service_account import ServiceAccountCredentials
import urllib.request
import re
import datetime
import time
import pytz  # JST用

def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")

# ==== Google Sheets API 認証設定 ====
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

# ==== スプレッドシートを開く（ID指定） ====
spreadsheet = gc.open_by_key("1Z8ajNuy4Q6Voh1P_EjqweRQnI_Obhk-SPCvAPCepIMc")
sheet_normal = spreadsheet.worksheet("紙書籍")
sheet_kindle = spreadsheet.worksheet("Kindle")

def clean_rankings(rankings):
    cleaned = []
    for r in rankings:
        r = r.replace("Amazon 売れ筋ランキング:", "").strip()
        r = r.replace("(本の売れ筋ランキングを見る)", "").strip()
        r = r.replace("(Kindleストアの売れ筋ランキングを見る)", "").strip()
        if r:
            cleaned.append(r)
    return cleaned

def get_rankings_from_url(url, keyword):
    log(f"{keyword}ページ取得開始")

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/115.0.0.0 Safari/537.36'
    }

    html = None
    for attempt in range(3):  # 最大3回リトライ
        try:
            req = urllib.request.Request(url, headers=headers)
            res = urllib.request.urlopen(req, timeout=15)
            html = res.read().decode('utf-8', errors='ignore')
            log(f"{keyword}ページ取得成功（試行{attempt+1}回目）")
            break
        except Exception as e:
            log(f"{keyword}ページ取得エラー（試行{attempt+1}回目）: {e}")
            if attempt < 2:
                time.sleep(3)  # 3秒待って再試行
    if html is None:
        return ['ランキング情報なし']

    # ==== デバッグログ: HTML先頭2000文字出力 ====
    print("----- HTML DEBUG START -----")
    print(html[:2000])
    print("----- HTML DEBUG END -----")

    # 「売れ筋ランキング」ブロックを正規表現で抽出
    pattern = r"Amazon 売れ筋ランキング:(.*?)カスタマーレビュー"
    match = re.search(pattern, html, re.S)
    if match:
        block = match.group(1)
        block = re.sub('<.*?>', '', block)
        block = re.sub(r'\s+', ' ', block)
        rankings = [r.strip() for r in block.split('-') if '位' in r]
        rankings = clean_rankings(rankings)
        log(f"{keyword}ランキング抽出完了: {rankings}")
        return rankings if rankings else ['ランキング情報なし']
    else:
        log(f"{keyword}ランキング情報なし")
        return ['ランキング情報なし']

log("処理開始")

# ==== JSTの現在時刻を取得 ====
jst = pytz.timezone('Asia/Tokyo')
now = datetime.datetime.now(jst).strftime('%Y/%m/%d %H:%M')

normal_url = 'https://www.amazon.co.jp/dp/4798183180'
kindle_url = 'https://www.amazon.co.jp/dp/B0CYPMKYM3'

normal_rankings = [now] + get_rankings_from_url(normal_url, '紙書籍')
kindle_rankings = [now] + get_rankings_from_url(kindle_url, 'Kindle')

try:
    log("スプレッドシート書き込み開始")
    sheet_normal.append_row(normal_rankings)
    sheet_kindle.append_row(kindle_rankings)
    log("スプレッドシート書き込み完了")
except Exception as e:
    log(f"スプレッドシート更新エラー: {e}")

log("処理完了")
