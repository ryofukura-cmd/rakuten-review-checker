#!/usr/bin/env python3
"""楽天商品レビュー定期チェッカー（Playwright版）"""

import os
import json
import hashlib
import time
import re
import requests
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread
from playwright.sync_api import sync_playwright

JST = ZoneInfo('Asia/Tokyo')

CHATWORK_TOKEN = os.environ['CHATWORK_API_TOKEN']
CHATWORK_ROOM  = os.environ.get('CHATWORK_ROOM_ID', '436382401')
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'ja,en-US;q=0.9',
}


# ── Google Sheets ──────────────────────────────────────────────────

def setup_gspread():
    import google.auth
    creds, _ = google.auth.default(scopes=[
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ])
    return gspread.authorize(creds)


def ensure_sheets(gc):
    sh = gc.open_by_key(SPREADSHEET_ID)
    existing = {ws.title for ws in sh.worksheets()}

    if '商品リスト' not in existing:
        ws = sh.add_worksheet('商品リスト', rows=200, cols=4)
        ws.update('A1:D1', [['商品名', '商品URL（サムネ用）', 'レビューURL', '監視(ON/OFF)']])
        ws.format('A1:D1', {'textFormat': {'bold': True}})
        print('シート「商品リスト」を作成しました')

    if '通知済み' not in existing:
        ws = sh.add_worksheet('通知済み', rows=5000, cols=6)
        ws.update('A1:F1', [['商品名', 'レビューハッシュ', '通知日時', 'レビュー日付', '評価', 'レビュー内容']])
        ws.format('A1:F1', {'textFormat': {'bold': True}})
        print('シート「通知済み」を作成しました')


def load_products(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('商品リスト').get_all_values()
    products = []
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        active = True
        if len(r) >= 4 and r[3].strip().upper() == 'OFF':
            active = False
        if not active:
            continue
        products.append({
            'name':        r[0].strip(),
            'product_url': r[1].strip() if len(r) > 1 else '',
            'review_url':  r[2].strip() if len(r) > 2 else '',
        })
    return products


def load_notified(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').get_all_values()
    return set(r[1] for r in rows[1:] if len(r) >= 2)


def save_notified(gc, product_name, h, review):
    stars = ('★' * review['rating'] + '☆' * (5 - review['rating'])) if review.get('rating') else '未評価'
    gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').append_row(
        [product_name, h, datetime.now(JST).strftime('%Y-%m-%d %H:%M'),
         review.get('date', ''), stars, review.get('text', '')]
    )


# ── チェック対象期間 ──────────────────────────────────────────────

def get_check_since():
    """過去30日分のレビューを対象にする（テスト用・本番は2日に戻す）"""
    return datetime.now(JST) - timedelta(days=30)


# ── 日付パース ────────────────────────────────────────────────────

def parse_date(text):
    if not text:
        return None
    m = re.search(
        r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日(?:[^\d]*(\d{1,2}):(\d{2}))?',
        text,
    )
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    h, mi = (int(m.group(4)), int(m.group(5))) if m.group(4) else (0, 0)
    return datetime(y, mo, d, h, mi, tzinfo=JST)


# ── サムネイル取得 ────────────────────────────────────────────────

def fetch_thumbnail(url):
    if not url:
        return None
    try:
        from bs4 import BeautifulSoup
        r = requests.get(url, headers=HEADERS, timeout=20)
        soup = BeautifulSoup(r.text, 'html.parser')
        og = soup.find('meta', property='og:image')
        return og['content'] if og else None
    except Exception as e:
        print(f'  サムネイル取得失敗: {e}')
        return None


# ── レビュースクレイピング（Playwright） ──────────────────────────

def build_review_url(review_url):
    """sort=6（新着順）を付与"""
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    return f'{base}?sort=6'


def _parse_review_json(data, since_dt, notified):
    """楽天レビューAPIのJSONからレビューを抽出する"""
    results = []
    # レビュー一覧が入っているキーを再帰的に探す
    def find_reviews(obj, depth=0):
        if depth > 6:
            return []
        found = []
        if isinstance(obj, list):
            for item in obj:
                found.extend(find_reviews(item, depth + 1))
        elif isinstance(obj, dict):
            # 日付っぽいキーがあればレビューアイテムの可能性
            date_keys = [k for k in obj if re.search(r'date|Date|time|Time|포스트|日時', k)]
            text_keys = [k for k in obj if re.search(r'text|body|comment|review|内容|口コミ', k, re.I)]
            if date_keys and text_keys:
                found.append(obj)
            else:
                for v in obj.values():
                    found.extend(find_reviews(v, depth + 1))
        return found

    candidates = find_reviews(data)
    print(f'    JSONレビュー候補: {len(candidates)} 件')

    for item in candidates:
        try:
            # 日付
            dt_text = ''
            for k in item:
                if re.search(r'date|Date|time|Time', k):
                    val = str(item[k])
                    if re.search(r'\d{4}年|\d{4}-\d{2}-\d{2}', val):
                        dt_text = val
                        break

            # YYYY-MM-DD 形式を変換
            if re.match(r'\d{4}-\d{2}-\d{2}', dt_text):
                ymd = dt_text[:10].split('-')
                dt_text = f'{ymd[0]}年{int(ymd[1])}月{int(ymd[2])}日'

            rev_dt = parse_date(dt_text)
            if not rev_dt:
                continue
            if rev_dt < since_dt:
                continue

            # 本文
            text = ''
            for k in item:
                if re.search(r'text|body|comment|review|内容|口コミ', k, re.I):
                    v = str(item[k]).strip()
                    if len(v) > len(text):
                        text = v
            if not text:
                continue

            # 評価
            rating = 0
            for k in item:
                if re.search(r'rating|score|star|評価|点', k, re.I):
                    try:
                        rating = int(float(str(item[k])))
                    except Exception:
                        pass

            h = hashlib.md5(f'{dt_text}{text}'.encode()).hexdigest()
            if h not in notified:
                results.append({'date': dt_text, 'text': text, 'rating': rating, 'hash': h})
        except Exception:
            continue
    return results


def scrape_reviews(review_url, since_dt, notified):
    results = []
    url = build_review_url(review_url)
    print(f'  URL: {url}')

    captured_jsons = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage',
            ]
        )
        context = browser.new_context(
            user_agent=HEADERS['User-Agent'],
            locale='ja-JP',
            viewport={'width': 1280, 'height': 900},
            extra_http_headers={'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'},
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = context.new_page()

        # レビュー関連のJSONレスポンスをすべて捕捉
        def on_response(response):
            try:
                if response.status != 200:
                    return
                ct = response.headers.get('content-type', '')
                if 'json' not in ct:
                    return
                rurl = response.url
                if any(kw in rurl for kw in ['review', 'Review', 'rvw', 'comment', 'rating']):
                    body = response.json()
                    captured_jsons.append({'url': rurl, 'data': body})
                    print(f'  JSON捕捉: {rurl[:120]}')
            except Exception:
                pass

        page.on('response', on_response)

        try:
            page.goto(url, wait_until='domcontentloaded', timeout=60000)
            # スクロールしてAJAX読み込みを誘発
            for i in range(8):
                page.evaluate(f'window.scrollTo(0, {(i+1) * 400})')
                time.sleep(1)
            time.sleep(5)

            print(f'  捕捉JSON数: {len(captured_jsons)}')

            # JSONから解析を試みる
            for cj in captured_jsons:
                r = _parse_review_json(cj['data'], since_dt, notified)
                if r:
                    results.extend(r)

            if results:
                browser.close()
                return results

            # フォールバック: HTML本文から日付パターン検索
            body_text = page.inner_text('body')
            print(f'  ページ文字数: {len(body_text)}')
            dates_in_body = re.findall(r'\d{4}年\d{1,2}月\d{1,2}日', body_text)
            print(f'  本文中の日付: {dates_in_body[:5]}')

            if not dates_in_body:
                print('  レビューデータ取得不可')
                browser.close()
                return []

            # 日付前後のテキストをレビューとして抽出
            for m in re.finditer(r'(\d{4}年\d{1,2}月\d{1,2}日)', body_text):
                dt_text = m.group(1)
                rev_dt = parse_date(dt_text)
                if not rev_dt or rev_dt < since_dt:
                    continue
                snippet = body_text[m.start():m.start() + 300]
                lines = [l.strip() for l in snippet.split('\n') if len(l.strip()) > 15]
                text = lines[0] if lines else ''
                if not text:
                    continue
                h = hashlib.md5(f'{dt_text}{text}'.encode()).hexdigest()
                if h not in notified:
                    results.append({'date': dt_text, 'text': text, 'rating': 0, 'hash': h})

        except Exception as e:
            print(f'  ページ読み込みエラー: {e}')
        finally:
            browser.close()

    return results


# ── Chatwork通知 ──────────────────────────────────────────────────

def notify_chatwork(product_name, review, thumbnail_url):
    stars  = ('★' * review['rating'] + '☆' * (5 - review['rating'])) if review['rating'] else '未評価'
    msg = (
        f"[info][title]🛒 新しいレビューが届きました！[/title]"
        f"📦 商品：{product_name}\n"
        f"📅 投稿日：{review['date']}\n"
        f"⭐ 評価：{stars}\n"
        f"💬 {review['text']}"
        f"[/info]"
    )

    hdrs = {'X-ChatWorkToken': CHATWORK_TOKEN}

    if thumbnail_url:
        try:
            img = requests.get(thumbnail_url, headers=HEADERS, timeout=20)
            if img.status_code == 200:
                r = requests.post(
                    f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM}/files',
                    headers=hdrs,
                    files={'file': ('thumbnail.jpg', img.content, 'image/jpeg')},
                    data={'message': msg},
                    timeout=30,
                )
                print(f'  Chatwork画像送信: {r.status_code}')
                if r.status_code in (200, 201):
                    return
        except Exception as e:
            print(f'  画像送信失敗（テキストで送信）: {e}')

    r = requests.post(
        f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM}/messages',
        headers=hdrs, data={'body': msg}, timeout=30,
    )
    print(f'  Chatworkテキスト送信: {r.status_code} {r.text[:200]}')


# ── メイン ────────────────────────────────────────────────────────

def main():
    since_dt = get_check_since()
    print(f'チェック対象: {since_dt:%Y-%m-%d %H:%M} 以降のレビュー\n')

    gc = setup_gspread()
    ensure_sheets(gc)

    products = load_products(gc)
    notified = load_notified(gc)
    print(f'{len(products)} 商品を処理します\n')

    for p in products:
        name        = p['name']
        product_url = p['product_url']
        review_url  = p['review_url']
        print(f'▶ {name}')

        if not review_url:
            print('  レビューURLが未設定（スプレッドシートのC列に入力してください）')
            continue

        thumb   = fetch_thumbnail(product_url)
        reviews = scrape_reviews(review_url, since_dt, notified)

        if not reviews:
            print('  新着レビューなし')
            continue

        print(f'  {len(reviews)} 件を通知')
        for rv in reviews:
            notify_chatwork(name, rv, thumb)
            save_notified(gc, name, rv['hash'], rv)
            notified.add(rv['hash'])
            time.sleep(0.5)

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
