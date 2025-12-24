"""Amazonランキング集計スクリプト（紙書籍・Kindle・Audible）

概要:
    指定したAmazon商品ページ（紙書籍・Kindle・Audible）の売れ筋ランキング情報を取得し、
    ローカルのExcelファイルおよびGoogleスプレッドシートに1行として追記するスクリプト。

主な仕様:
    - Amazon商品ページから「売れ筋ランキング」ブロックを抽出し、カテゴリ名と順位を解析する。
    - 紙書籍は最大4カテゴリ、それ以外（Kindle・Audible）は最大2カテゴリのランキングを取得する。
    - JSTの現在時刻（時単位に丸め）を先頭列に付与し、ランキング情報を右方向に並べて保存する。
    - Excel書き込み時にはリトライ処理と列幅自動調整を行う。
    - Googleスプレッドシートには追記後、日時列（A列）で降順ソートを実施する。

制限事項:
    - Amazonサイト側のHTML構造や文言仕様が変更された場合、ランキング抽出が正しく動作しない可能性がある。
    - ネットワーク環境や認証情報（Google API）に依存しており、環境が整っていない場合は処理が失敗する。
"""

import urllib.request
import re
import datetime
import pytz
import openpyxl
import time
import os
import json

import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials

BASE_DIR = os.path.dirname(__file__)
SHEET_NAME = 'Amazon 売れ筋ランキング'


def log(msg):
    """
    概要:
        ログメッセージをコンソールおよびログファイルに出力する。

    Args:
        msg (str): 出力したいログメッセージ。

    Returns:
        None
    """
    timestamp = time.strftime('%H:%M:%S')
    print(f"[{timestamp}] {msg}")
    try:
        with open('amazonranking_log.txt', 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {msg}\n")
    except Exception as e:
        print(f"ログ書き込みエラー: {e}")


def extract_rankings_from_html(html, keyword):
    """
    概要:
        Amazon商品ページのHTML文字列から、売れ筋ランキング情報を抽出する。
        紙書籍は最大4件、それ以外（Kindle・Audible）は最大2件に正規化して返却する。

    Args:
        html (str): 対象となるAmazon商品ページのHTML文字列。
        keyword (str): ランキング種別を示すキーワード（'紙書籍' / 'Kindle' / 'Audible' など）。

    Returns:
        list[str]: 抽出されたランキング情報のリスト。
                   - keyword が '紙書籍' の場合: 長さ4のリスト。
                   - 上記以外の場合（Kindle・Audible等）: 長さ2のリスト。
                   - 抽出に失敗した場合は、上記長さ分の '-' を要素とするリスト。
    """
    rankings = []

    # 売れ筋ランキング部分を検出（文言揺れにある程度耐えるように検索語を緩める）
    start = html.find("Amazon 売れ筋ランキング:")
    if start == -1:
        start = html.find("売れ筋ランキング:")
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
    block = re.sub(r'\(Audibleの売れ筋ランキングを見る\)', '', block)
    block = re.sub(r'本の売れ筋ランキングを見る', '', block)
    block = re.sub(r'Kindleストアの売れ筋ランキングを見る', '', block)
    block = re.sub(r'Audibleの売れ筋ランキングを見る', '', block)

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

    # 列数統一（紙書籍4列、それ以外2列: Kindle / Audible 等）
    expected_len = 4 if keyword == '紙書籍' else 2
    if len(rankings) < expected_len:
        rankings += ['-'] * (expected_len - len(rankings))
    elif len(rankings) > expected_len:
        rankings = rankings[:expected_len]

    log(f"{keyword}ランキング抽出完了: {rankings}")
    return rankings


def get_rankings_from_url(url, keyword):
    """
    概要:
        指定されたAmazon商品ページURLからHTMLを取得し、ランキング情報を抽出する。

    Args:
        url (str): 対象となるAmazon商品ページのURL。
        keyword (str): ランキング種別を示すキーワード（'紙書籍' / 'Kindle' / 'Audible' など）。

    Returns:
        list[str]: 抽出されたランキング情報のリスト。
                   - keyword が '紙書籍' の場合: 長さ4のリスト。
                   - 上記以外の場合（Kindle・Audible等）: 長さ2のリスト。
                   - HTML取得や解析に失敗した場合は、上記長さ分の '-' を要素とするリスト。
    """
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
    """
    概要:
        指定されたExcelファイルパスに対して、行データを追記保存する。
        PermissionError 等で書き込みに失敗した場合は、指定回数までリトライを行う。

    Args:
        excel_path (str): 書き込み対象のExcelファイルパス。
        row_data (list): 1行分の書き込みデータ（セル値のリスト）。
        max_retries (int): 最大リトライ回数。

    Returns:
        bool: 書き込みに成功した場合は True、全ての試行で失敗した場合は False。
    """
    for attempt in range(max_retries):
        try:
            log(f"Excel保存試行 {attempt + 1}/{max_retries}")

            if os.path.exists(excel_path):
                wb = openpyxl.load_workbook(excel_path)
            else:
                wb = openpyxl.Workbook()
                if 'Sheet' in wb.sheetnames:
                    wb.remove(wb['Sheet'])

            # 既存のSheet1のみが存在する場合はターゲット名にリネームし、なければ作成
            if SHEET_NAME not in wb.sheetnames:
                if 'Sheet1' in wb.sheetnames and len(wb.sheetnames) == 1:
                    wb['Sheet1'].title = SHEET_NAME
                    log(f"Excelシート名を {SHEET_NAME} に変更しました")
                else:
                    wb.create_sheet(SHEET_NAME)
                    log(f"Excelにシート {SHEET_NAME} を新規作成しました")

            ws = wb[SHEET_NAME]

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
    """
    概要:
        Googleスプレッドシートに対して、1行分のデータを追記し、
        その後、日時列（A列）をキーとして降順ソートを行う。

    Args:
        row_data (list): 1行分の書き込みデータ（セル値のリスト）。

    Returns:
        None
    """
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

        workbook = client.open_by_key(SPREADSHEET_ID)
        try:
            worksheet = workbook.worksheet(SHEET_NAME)
        except WorksheetNotFound:
            log(f"{SHEET_NAME} シートが存在しないため新規作成します")
            worksheet = workbook.add_worksheet(title=SHEET_NAME, rows="10000", cols="20")

        # ▼ まず行追加
        worksheet.append_row(row_data, value_input_option='USER_ENTERED')
        log("Googleスプレッドシートに追記完了")

        # ▼ 日付のある列（A列）で降順ソート
        worksheet.sort((1, 'des'))
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
audible_url = 'https://www.amazon.co.jp/dp/B0G66DNXDH'

normal_rankings = get_rankings_from_url(normal_url, '紙書籍')
kindle_rankings = get_rankings_from_url(kindle_url, 'Kindle')
audible_rankings = get_rankings_from_url(audible_url, 'Audible')

row_data = [now] + normal_rankings + kindle_rankings + audible_rankings

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
