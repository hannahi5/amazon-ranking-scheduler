
import re
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright

def log(msg):
    now = datetime.datetime.now().strftime('%H:%M:%S')
    print(f"[{now}] {msg}")

# ==== Google Sheets API 認証設定 ====
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

spreadsheet = gc.open_by_key("1Z8ajNuy4Q6Voh1P_EjqweRQnI_Obhk-SPCvAPCepIMc")
sheet_normal = spreadsheet.worksheet("紙書籍")
sheet_kindle = spreadsheet.worksheet("Kindle")

def extract_normal_rankings(html):
    block_start = html.find("Amazon 売れ筋ランキング")
    if block_start == -1:
        return ["ランキング情報なし"]

    block_end = html.find("カスタマーレビュー", block_start)
    block = html[block_start:block_end] if block_end != -1 else html[block_start:]

    text = re.sub('<.*?>', '', block)
    text = text.replace("Amazon 売れ筋ランキング:", "").replace("本の売れ筋ランキングを見る", "")
    text = re.sub(r'\s+', ' ', text).strip()
    rankings = [r.strip() for r in text.split('-') if r.strip()]
    return rankings if rankings else ["ランキング情報なし"]

def extract_kindle_rankings(html):
    block_start = html.find("Amazon 売れ筋ランキング")
    if block_start == -1:
        return ["ランキング情報なし"]

    block_end = html.find("カスタマーレビュー", block_start)
    block = html[block_start:block_end] if block_end != -1 else html[block_start:]

    text = re.sub('<.*?>', '', block)
    text = text.replace("Amazon 売れ筋ランキング:", "").replace("Kindleストアの売れ筋ランキングを見る", "")
    text = re.sub(r'\s+', ' ', text).strip()
    rankings = [r.strip() for r in text.split('-') if r.strip()]
    return rankings if rankings else ["ランキング情報なし"]

def get_html_via_playwright(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, timeout=60000)
        
        page.wait_for_selector("#zg-rank-ctnr", timeout=10000)  # ランキング要素が出るまで最大10秒待つ
        page.screenshot(path="debug.png")  # スクショ保存
        
        html = page.content()
        browser.close()
        return html

def main():
    log("処理開始")
    now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')

    normal_url = 'https://www.amazon.co.jp/dp/4798183180'
    kindle_url = 'https://www.amazon.co.jp/dp/B0CYPMKYM3'

    # 紙書籍ランキング取得
    log("紙書籍ページ取得開始")
    normal_html = get_html_via_playwright(normal_url)
    normal_rankings = [now] + extract_normal_rankings(normal_html)

    # Kindleランキング取得
    log("Kindleページ取得開始")
    kindle_html = get_html_via_playwright(kindle_url)
    kindle_rankings = [now] + extract_kindle_rankings(kindle_html)

    try:
        log("スプレッドシート書き込み開始")
        sheet_normal.append_row(normal_rankings)
        sheet_kindle.append_row(kindle_rankings)
        log("スプレッドシート書き込み完了")
    except Exception as e:
        log(f"スプレッドシート更新エラー: {e}")

    log("処理完了")

if __name__ == "__main__":
    main()
