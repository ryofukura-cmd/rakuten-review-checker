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
        ws = sh.add_worksheet('通知済み', rows=5000, cols=3)
        ws.update('A1:C1', [['商品名', 'レビューハッシュ', '通知日時']])
        ws.format('A1:C1', {'textFormat': {'bold': True}})
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


def save_notified(gc, product_name, h):
    gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').append_row(
        [product_name, h, datetime.now(JST).strftime('%Y-%m-%d %H:%M')]
    )


# ── 時間ウィンドウ ────────────────────────────────────────────────

def get_time_window():
    now  = datetime.now(JST)
    base = now.replace(minute=0, second=0, microsecond=0)
    if now.hour == 9:
        start = (now - timedelta(days=1)).replace(hour=18, minute=0, second=0, microsecond=0)
    else:
        start = base - timedelta(hours=1)
    return start, base


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


def scrape_reviews(review_url, start_dt, end_dt, notified):
    results = []
    url = build_review_url(review_url)
    print(f'  URL: {url}')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox'])
        context = browser.new_context(
            user_agent=HEADERS['User-Agent'],
            locale='ja-JP',
        )
        page = context.new_page()

        try:
            page.goto(url, wait_until='domcontentloaded', timeout=30000)
            # ページが安定するまで少し待つ
            time.sleep(3)

            # レビュー要素を探す（複数セレクタで試行）
            review_items = []
            selectors = [
                '[class*="review-item"]',
                '[class*="reviewItem"]',
                '[class*="ReviewItem"]',
                '[class*="review_item"]',
                'li[class*="review"]',
                'article[class*="review"]',
                '[data-review-id]',
                '[class*="revRvw"]',
            ]

            found_selector = None
            for sel in selectors:
                try:
                    items = page.query_selector_all(sel)
                    if items:
                        review_items = items
                        found_selector = sel
                        print(f'  セレクタ「{sel}」で {len(items)} 件検出')
                        break
                except Exception:
                    continue

            if not review_items:
                # フォールバック: 日付パターンを含む要素を探す
                print('  セレクタ未検出。テキストから日付を探します')
                body = page.inner_text('body')
                dates = re.findall(r'\d{4}年\d{1,2}月\d{1,2}日', body)
                print(f'  本文中の日付: {dates[:5]}')
                browser.close()
                return []

            for item in review_items:
                try:
                    full_text = item.inner_text()

                    # 日付
                    dt_text = ''
                    for sel in ['[class*="date"]', '[class*="Date"]', 'time', '[class*="post"]']:
                        el = item.query_selector(sel)
                        if el:
                            dt_text = el.get_attribute('datetime') or el.inner_text()
                            if re.search(r'\d{4}年', dt_text):
                                break
                    if not dt_text:
                        m = re.search(r'\d{4}年\d{1,2}月\d{1,2}日', full_text)
                        if m:
                            dt_text = m.group(0)

                    rev_dt = parse_date(dt_text)
                    if not rev_dt:
                        continue

                    # 新着順なので古いレビューに達したら終了
                    if rev_dt < start_dt:
                        print(f'  {dt_text} は対象期間より古い → 終了')
                        results = results  # 現在の結果を保持
                        browser.close()
                        return results

                    if rev_dt >= end_dt:
                        continue

                    # レビュー本文（長い行を優先）
                    lines = [l.strip() for l in full_text.split('\n') if len(l.strip()) > 15]
                    text = lines[0] if lines else ''

                    # 評価（数字）
                    rating = 0
                    for sel in ['[class*="rating"]', '[class*="star"]', '[aria-label*="点"]']:
                        el = item.query_selector(sel)
                        if el:
                            aria = el.get_attribute('aria-label') or ''
                            txt  = el.inner_text()
                            m = re.search(r'([1-5])', aria + txt)
                            if m:
                                rating = int(m.group(1))
                                break

                    if not text:
                        continue

                    h = hashlib.md5(f'{dt_text}{text}'.encode()).hexdigest()
                    if h not in notified:
                        results.append({
                            'date': dt_text, 'text': text,
                            'rating': rating, 'hash': h,
                        })

                except Exception as e:
                    print(f'  レビュー要素パースエラー: {e}')
                    continue

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
                if r.status_code in (200, 201):
                    return
        except Exception as e:
            print(f'  画像送信失敗（テキストで送信）: {e}')

    requests.post(
        f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM}/messages',
        headers=hdrs, data={'body': msg}, timeout=30,
    )


# ── メイン ────────────────────────────────────────────────────────

def main():
    start_dt, end_dt = get_time_window()
    print(f'チェック期間: {start_dt:%Y-%m-%d %H:%M} ～ {end_dt:%Y-%m-%d %H:%M} JST\n')

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
        reviews = scrape_reviews(review_url, start_dt, end_dt, notified)

        if not reviews:
            print('  新着レビューなし')
            continue

        print(f'  {len(reviews)} 件を通知')
        for rv in reviews:
            notify_chatwork(name, rv, thumb)
            save_notified(gc, name, rv['hash'])
            notified.add(rv['hash'])
            time.sleep(0.5)

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
