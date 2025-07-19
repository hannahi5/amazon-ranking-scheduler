import gspread
from oauth2client.service_account import ServiceAccountCredentials
import urllib.request
import re
import datetime
import time


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
    try:
        res = urllib.request.urlopen(url, timeout=15)
        html = res.read().decode('utf-8')
        log(f"{keyword}ページ取得完了（HTMLサイズ: {len(html)} bytes）")
    except Exception as e:
        log(f"{keyword}ページ取得エラー: {e}")
        return ['ランキング情報なし']

    start = html.find("Amazon 売れ筋ランキング:")
    end = html.find("カスタマーレビュー", start)
    if start != -1 and end != -1:
        block = html[start:end]
        block = re.sub('<.*?>', '', block)
        block = re.sub(r'\s+', ' ', block)
        rankings = [r.strip() for r in block.split('-') if r.strip()]
        rankings = clean_rankings(rankings)
        log(f"{keyword}ランキング抽出完了: {rankings}")
        return rankings
    else:
        log(f"{keyword}ランキング情報なし")
        return ['ランキング情報なし']

log("処理開始")
now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')

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
