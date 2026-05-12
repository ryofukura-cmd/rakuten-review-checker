#!/usr/bin/env python3
"""楽天商品レビュー定期チェッカー"""

import os
import json
import hashlib
import time
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread

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
    'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
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
    """シートが存在しない場合は自動作成"""
    sh = gc.open_by_key(SPREADSHEET_ID)
    existing = {ws.title for ws in sh.worksheets()}

    if '商品リスト' not in existing:
        ws = sh.add_worksheet('商品リスト', rows=200, cols=3)
        ws.update('A1:C1', [['商品名', '楽天商品URL', '監視(ON/OFF)']])
        ws.format('A1:C1', {'textFormat': {'bold': True}})
        print('シート「商品リスト」を作成しました')

    if '通知済み' not in existing:
        ws = sh.add_worksheet('通知済み', rows=5000, cols=3)
        ws.update('A1:C1', [['商品URL', 'レビューハッシュ', '通知日時']])
        ws.format('A1:C1', {'textFormat': {'bold': True}})
        print('シート「通知済み」を作成しました')


def load_products(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('商品リスト').get_all_values()
    return [
        {'name': r[0].strip(), 'url': r[1].strip()}
        for r in rows[1:]
        if len(r) >= 2
        and r[1].strip()
        and (len(r) < 3 or r[2].strip().upper() != 'OFF')
    ]


def load_notified(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').get_all_values()
    return set(r[1] for r in rows[1:] if len(r) >= 2)


def save_notified(gc, url, h):
    gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').append_row(
        [url, h, datetime.now(JST).strftime('%Y-%m-%d %H:%M')]
    )


# ── 時間ウィンドウ ────────────────────────────────────────────────

def get_time_window():
    """
    9時  → 前日18:00〜当日09:00
    10〜18時 → 直前1時間
    """
    now  = datetime.now(JST)
    base = now.replace(minute=0, second=0, microsecond=0)
    if now.hour == 9:
        start = (now - timedelta(days=1)).replace(hour=18, minute=0, second=0, microsecond=0)
    else:
        start = base - timedelta(hours=1)
    return start, base


# ── スクレイピング ────────────────────────────────────────────────

def parse_rakuten_url(url):
    m = re.search(r'item\.rakuten\.co\.jp/([^/?#]+)/([^/?#]+)', url.rstrip('/'))
    return (m.group(1), m.group(2)) if m else (None, None)


def fetch_thumbnail(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        soup = BeautifulSoup(r.text, 'html.parser')
        og = soup.find('meta', property='og:image')
        return og['content'] if og else None
    except Exception as e:
        print(f'  サムネイル取得失敗: {e}')
        return None


def parse_date(text):
    m = re.search(
        r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日(?:[^\d]*(\d{1,2}):(\d{2}))?',
        text,
    )
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    h, mi = (int(m.group(4)), int(m.group(5))) if m.group(4) else (0, 0)
    return datetime(y, mo, d, h, mi, tzinfo=JST)


def scrape_reviews(shop_id, item_id, start_dt, end_dt, notified):
    results = []

    for page in range(1, 11):
        url = f'https://review.rakuten.co.jp/item/1/{shop_id}.{item_id}/{page}/'
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 404:
                break
            if r.status_code != 200:
                print(f'  HTTP {r.status_code}')
                break

            soup = BeautifulSoup(r.text, 'html.parser')

            # レビューブロックを探す（セレクタ優先順）
            blocks = (
                soup.select('.revRvwUserEntry')
                or soup.select('[class*="revRvwUser"]')
                or soup.select('.review-item')
                or soup.select('[itemtype*="Review"]')
            )
            if not blocks:
                print(f'  レビューブロック未検出 (p{page})')
                break

            stop = False
            for b in blocks:
                # 日付
                dt_text = ''
                for sel in ['.revDate', 'time', '[class*="date"]', '[class*="Date"]']:
                    el = b.select_one(sel)
                    if el:
                        dt_text = el.get('datetime', '') or el.get_text(strip=True)
                        break
                if not dt_text:
                    m = re.search(r'\d{4}年\d{1,2}月\d{1,2}日', b.get_text())
                    dt_text = m.group(0) if m else ''

                rev_dt = parse_date(dt_text)
                if not rev_dt:
                    continue
                if rev_dt < start_dt:
                    stop = True
                    break
                if rev_dt >= end_dt:
                    continue

                # レビュー本文
                text = ''
                for sel in ['.revComment', '[class*="comment"]', '[class*="Comment"]', 'p']:
                    el = b.select_one(sel)
                    if el and len(el.get_text(strip=True)) > 5:
                        text = el.get_text(strip=True)
                        break
                if not text:
                    continue

                # タイトル
                title = ''
                for sel in ['.revTitle', '[class*="title"]', '[class*="Title"]']:
                    el = b.select_one(sel)
                    if el:
                        title = el.get_text(strip=True)
                        break

                # 評価（1〜5）
                rating = 0
                for sel in ['[class*="rating"]', '[class*="Rating"]', '[class*="star"]', '[itemprop="ratingValue"]']:
                    el = b.select_one(sel)
                    if el:
                        val = el.get('content', '') or el.get_text()
                        m = re.search(r'[1-5]', str(val))
                        if m:
                            rating = int(m.group(0))
                            break

                h = hashlib.md5(f'{dt_text}{text}'.encode()).hexdigest()
                if h not in notified:
                    results.append({
                        'date': dt_text,
                        'title': title,
                        'text': text,
                        'rating': rating,
                        'hash': h,
                    })

            if stop:
                break
            time.sleep(1)

        except Exception as e:
            print(f'  取得エラー p{page}: {e}')
            break

    return results


# ── Chatwork通知 ──────────────────────────────────────────────────

def notify_chatwork(product_name, review, thumbnail_url):
    stars = ('★' * review['rating'] + '☆' * (5 - review['rating'])) if review['rating'] else '未評価'
    title_line = f"📝 {review['title']}\n" if review['title'] else ''

    msg = (
        f"[info][title]🛒 新しいレビューが届きました！[/title]"
        f"📦 商品：{product_name}\n"
        f"📅 投稿日：{review['date']}\n"
        f"⭐ 評価：{stars}\n"
        f"{title_line}"
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
        headers=hdrs,
        data={'body': msg},
        timeout=30,
    )


# ── メイン ────────────────────────────────────────────────────────

def main():
    start_dt, end_dt = get_time_window()
    print(f'チェック期間: {start_dt:%Y-%m-%d %H:%M} ～ {end_dt:%Y-%m-%d %H:%M} JST')

    gc = setup_gspread()
    ensure_sheets(gc)

    products = load_products(gc)
    notified = load_notified(gc)
    print(f'{len(products)} 商品を処理します\n')

    for p in products:
        name, url = p['name'], p['url']
        print(f'▶ {name}')

        shop_id, item_id = parse_rakuten_url(url)
        if not shop_id:
            print('  URLパース失敗（item.rakuten.co.jp 形式のURLを確認してください）')
            continue

        thumb   = fetch_thumbnail(url)
        reviews = scrape_reviews(shop_id, item_id, start_dt, end_dt, notified)

        if not reviews:
            print('  新着レビューなし')
            continue

        print(f'  {len(reviews)} 件を通知')
        for rv in reviews:
            notify_chatwork(name, rv, thumb)
            save_notified(gc, url, rv['hash'])
            notified.add(rv['hash'])
            time.sleep(0.5)

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
