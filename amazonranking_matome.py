import urllib.request
import re
import datetime
import pytz
import openpyxl
import time
import os
import json

import gspread
from oauth2client.service_account import ServiceAccountCredentials

BASE_DIR = os.path.dirname(__file__)


def log(msg):
    """ログ出力（コンソール＋ファイル）"""
    timestamp = time.strftime('%H:%M:%S')
    print(f"[{timestamp}] {msg}")
    try:
        with open('amazonranking_log.txt', 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {msg}\n")
    except Exception as e:
        print(f"ログ書き込みエラー: {e}")


def extract_rankings_from_html(html, keyword):
    """HTMLからランキング情報を可変カテゴリ対応で抽出（不要語除去版）"""
    rankings = []

    # 売れ筋ランキング部分を検出
    start = html.find("Amazon 売れ筋ランキング:")
    if start == -1:
        log(f"{keyword}ランキング情報が見つかりません")
        return ['-'] * (4 if keyword == '紙書籍' else 2)

    end = html.find("カスタマーレビュー", start)
    if end == -1:
        end = len(html)
    block = html[start:end]

    # HTMLタグを除去
    block = re.sub('<.*?>', '', block)
    # 余分な空白除去
    block = re.sub(r'\s+', ' ', block)

    # 不要語削除
    block = re.sub(r'（?本の売れ筋ランキングを見る）?', '', block)
    block = re.sub(r'\(Kindleストアの売れ筋ランキングを見る\)', '', block)
    block = re.sub(r'本の売れ筋ランキングを見る', '', block)
    block = re.sub(r'Kindleストアの売れ筋ランキングを見る', '', block)

    log(f"{keyword}処理前のブロック: {block[:200]}")

    # 汎用正規表現：「カテゴリ名 - 順位」
    pattern = r'([^\-:：]{2,80}?)\s*[-−]\s*(\d{1,3}(?:,\d{3})*位)'

    matches = re.findall(pattern, block)
    if not matches:
        log(f"{keyword}ランキングパターンに一致なし")
        return ['-'] * (4 if keyword == '紙書籍' else 2)

    for name, rank in matches:
        name = name.strip()
        rank = rank.strip()
        # ノイズ除去
        if "Amazon" in name or "見る" in name:
            continue
        text = f"{rank}{name}"
        # 空のかっこ「()」を削除
        text = re.sub(r'\(\s*\)', '', text)
        rankings.append(text)

    # 列数統一（紙書籍4列、Kindle2列）
    expected_len = 4 if keyword == '紙書籍' else 2
    if len(rankings) < expected_len:
        rankings += ['-'] * (expected_len - len(rankings))
    elif len(rankings) > expected_len:
        rankings = rankings[:expected_len]

    log(f"{keyword}ランキング抽出完了: {rankings}")
    return rankings


def get_rankings_from_url(url, keyword):
    """Amazonページからランキングを取得"""
    log(f"{keyword}ページ取得開始")

    try:
        res = urllib.request.urlopen(url, timeout=15)
        html = res.read().decode('utf-8')
        log(f"{keyword}ページ取得完了（HTMLサイズ: {len(html)} bytes）")
    except Exception as e:
        log(f"{keyword}ページ取得エラー: {e}")
        if keyword == '紙書籍':
            return ['-'] * 4
        else:
            return ['-'] * 2

    return extract_rankings_from_html(html, keyword)


def save_to_excel_with_retry(excel_path, row_data, max_retries=3):
    """Excelファイルに保存（リトライ対応）"""
    for attempt in range(max_retries):
        try:
            log(f"Excel保存試行 {attempt + 1}/{max_retries}")

            if os.path.exists(excel_path):
                wb = openpyxl.load_workbook(excel_path)
            else:
                wb = openpyxl.Workbook()
                if 'Sheet' in wb.sheetnames:
                    wb.remove(wb['Sheet'])

            if 'Sheet1' not in wb.sheetnames:
                wb.create_sheet('Sheet1')
            ws = wb['Sheet1']

            # 行追加
            ws.append(row_data)

            # 列幅調整
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(excel_path)
            log("Excel書き込み完了")
            return True

        except PermissionError as e:
            log(f"権限エラー (試行 {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                log("5秒待機してリトライします...")
                time.sleep(5)
            else:
                log("Excelファイルが開かれている可能性があります。")
                return False
        except Exception as e:
            log(f"予期しないエラー: {e}")
            return False

    return False

def append_to_google_sheet(row_data):
    """Googleスプレッドシートに1行追記し、日時列で降順ソートする"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        log("環境変数 GOOGLE_CREDENTIALS が見つかりません")
        return

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        info = json.loads(creds_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
        client = gspread.authorize(creds)

        SPREADSHEET_ID = "1Rv5jj-ix-klCYptudYcxfd-j1jnOPeML"

        sheet = client.open_by_key(SPREADSHEET_ID).sheet1

        # ▼ まず行追加
        sheet.append_row(row_data, value_input_option='USER_ENTERED')
        log("Googleスプレッドシートに追記完了")

        # ▼ 日付のある列（A列）で降順ソート
        sheet.sort((1, 'des'))
        log("スプレッドシートを日時降順で並べ替えました")

    except Exception as e:
        log(f"Googleスプレッドシート書き込みエラー: {e}")


# -------------------------------
# メイン処理
# -------------------------------
log("処理開始")
JST = pytz.timezone('Asia/Tokyo')

now_dt = datetime.datetime.now(JST)

# 分以下を 00:00 に丸める
rounded = now_dt.replace(minute=0, second=0, microsecond=0)

now = rounded.strftime('%Y/%m/%d %H:%M')

normal_url = 'https://www.amazon.co.jp/dp/4798183180'
kindle_url = 'https://www.amazon.co.jp/dp/B0CYPMKYM3'

normal_rankings = get_rankings_from_url(normal_url, '紙書籍')
kindle_rankings = get_rankings_from_url(kindle_url, 'Kindle')

row_data = [now] + normal_rankings + kindle_rankings

log(f"構築された行データ: {row_data}")
log(f"データ列数: {len(row_data)}")

# ① Googleスプレッドシートへ追記
append_to_google_sheet(row_data)

# ② （オプション）Excelにも保存しておきたい場合はそのまま残す
try:
    log("Excel書き込み開始")
    excel_path = os.path.join(BASE_DIR, 'amazonranking_matome.xlsx')

    if not save_to_excel_with_retry(excel_path, row_data):
        log("Excel保存に失敗。バックアップに切替。")
        backup_path = os.path.join(BASE_DIR, 'amazonranking_matome_backup.xlsx')
        if save_to_excel_with_retry(backup_path, row_data):
            log("バックアップファイルに保存しました。")
        else:
            log("バックアップファイルへの保存も失敗しました。")

except Exception as e:
    log(f"Excel更新エラー: {e}")
    import traceback
    log(f"エラー詳細: {traceback.format_exc()}")

log("処理完了")
