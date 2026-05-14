#!/usr/bin/env python3
"""楽天市場 低評価（★1-2）レビュー 毎時アラート"""

import os
import re
import json
import hashlib
import time
import requests
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread

JST = ZoneInfo('Asia/Tokyo')

CHATWORK_TOKEN = os.environ['CHATWORK_API_TOKEN']
CHATWORK_ROOM  = os.environ.get('CHATWORK_ROOM_ID', '436633458')
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']
COMPANY_NAME   = os.environ.get('COMPANY_NAME', 'さくらフォレスト株式会社')
LOW_RATING_MAX = 2

WEEKDAYS_JA = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
ALERT_SHEET  = '低評価通知済み'

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


def ensure_alert_sheet(gc):
    sh = gc.open_by_key(SPREADSHEET_ID)
    existing = {ws.title for ws in sh.worksheets()}
    if ALERT_SHEET not in existing:
        ws = sh.add_worksheet(ALERT_SHEET, rows=5000, cols=10)
        print(f'シート「{ALERT_SHEET}」を作成しました')
    else:
        ws = sh.worksheet(ALERT_SHEET)
    ws.update(
        [['商品名', 'レビュー日付', '評価', 'タイトル', '本文', '投稿者名', '性別', '年齢', '通知日時', 'レビューハッシュ']],
        'A1:J1',
    )
    ws.format('A1:J1', {'textFormat': {'bold': True}})
    return ws


def load_products(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('商品リスト').get_all_values()
    products = []
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        if len(r) >= 4 and r[3].strip().upper() == 'OFF':
            continue
        products.append({
            'name':       r[0].strip(),
            'review_url': r[2].strip() if len(r) > 2 else '',
        })
    return products


def load_notified(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet(ALERT_SHEET).get_all_values()
    return set(r[9] for r in rows[1:] if len(r) >= 10)


def save_notified(ws, product_name, rv):
    now    = datetime.now(JST).strftime('%Y-%m-%d %H:%M')
    rating = rv.get('rating', 0)
    stars  = f'{"★" * rating}{"☆" * (5 - rating)} {rating}点' if rating else '未評価'
    ws.append_rows([[
        product_name,
        rv.get('date', ''),
        stars,
        rv.get('title', ''),
        rv.get('body', ''),
        rv.get('reviewer_name', ''),
        rv.get('gender', ''),
        rv.get('age', ''),
        now,
        rv['hash'],
    ]], value_input_option='USER_ENTERED')


# ── 稼働日チェック ────────────────────────────────────────────────

def is_business_day():
    import jpholiday
    today = datetime.now(JST).date()
    if today.weekday() >= 5:
        return False
    if jpholiday.is_holiday(today):
        return False
    if (today.month == 12 and today.day >= 28) or (today.month == 1 and today.day <= 4):
        return False
    if today.month == 8 and 10 <= today.day <= 16:
        return False
    return True


# ── URL ユーティリティ ─────────────────────────────────────────────

def build_scrape_url(review_url):
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    if not re.search(r'/\d+/?$', base):
        base = f'{base}/1'
    return f'{base}/?sort=6'


def build_display_url(review_url):
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    base = re.sub(r'/\d+$', '', base)
    return f'{base}?sort=6'


def parse_date(text):
    if not text:
        return None
    m = re.search(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
    if not m:
        return None
    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), tzinfo=JST)


# ── スクレイピング（★1-2 のみ）────────────────────────────────────

def scrape_low_reviews(review_url, since_dt, notified):
    results = []
    url = build_scrape_url(review_url)
    print(f'  URL: {url}')

    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        html = resp.text

        m = re.search(r'window\.__INITIAL_STATE__\s*=\s*', html)
        if not m:
            print('  __INITIAL_STATE__ が見つかりません')
            return []

        start = m.end()
        while start < len(html) and html[start] in ' \t\n\r':
            start += 1
        decoder = json.JSONDecoder()
        data, _ = decoder.raw_decode(html, start)

        reviews_section = data.get('reviews', {})
        item_reviews    = reviews_section.get('itemReviews', {})
        reviews_data    = reviews_section.get('data', {})
        uuids           = item_reviews.get('keys', [])
        total           = item_reviews.get('count', 0)
        print(f'  全レビュー数: {total}, このページ: {len(uuids)}件')

        for uuid in uuids:
            rv = reviews_data.get(uuid, {})
            if not isinstance(rv, dict):
                continue

            rating = rv.get('rating', 0)
            if rating > LOW_RATING_MAX:
                continue

            post_date = rv.get('postDate', '')
            dm = re.match(r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})', post_date)
            if not dm:
                continue
            dt_text = f'{dm.group(1)}年{int(dm.group(2))}月{int(dm.group(3))}日'
            rev_dt = parse_date(dt_text)
            if not rev_dt or rev_dt < since_dt:
                continue

            body  = rv.get('body', '').strip()
            title = rv.get('title', '').strip()
            if not body and not title:
                continue

            nickname   = re.sub(r'さん$', '', rv.get('nickname', '')).strip()
            sex        = rv.get('sex', '')
            gender     = {'male': '男性', 'female': '女性'}.get(sex, sex)
            age_range  = str(rv.get('ageRange', ''))
            age_suffix = rv.get('ageSuffix', '代')
            age        = f'{age_range}{age_suffix}' if age_range else ''

            h = hashlib.md5(f'{dt_text}{body or title}'.encode()).hexdigest()
            if h not in notified:
                results.append({
                    'date':          dt_text,
                    'title':         title,
                    'body':          body,
                    'rating':        rating,
                    'reviewer_name': nickname,
                    'gender':        gender,
                    'age':           age,
                    'hash':          h,
                })

    except Exception as e:
        print(f'  レビュー取得エラー: {e}')

    return results


# ── Chatwork通知（1件ずつ）────────────────────────────────────────

def build_alert_message(product_name, rv, display_url):
    today    = datetime.now(JST)
    date_str = f'[ {today.year}年{today.month}月{today.day}日 {WEEKDAYS_JA[today.weekday()]} ]'

    rating = rv.get('rating', 0)
    stars  = '★' * rating + '☆' * (5 - rating)

    name   = rv.get('reviewer_name', '')
    gender = rv.get('gender', '')
    age    = rv.get('age', '')
    if name:
        profile = f'{name} 様'
        parts = [x for x in [gender, age] if x]
        if parts:
            profile += f'（{"・".join(parts)}）'
    else:
        profile = '購入者様'

    title = rv.get('title', '')
    body  = rv.get('body', '')

    lines = [f'[info][title]⚠ 低評価レビュー {stars} {rating}/5[/title]']
    lines += [f'【{COMPANY_NAME}】', date_str, '']
    lines += [f'■ {product_name}', '']
    lines.append(f'投稿日：{rv.get("date", "")}')
    lines.append(profile)
    if title:
        lines.append(f'「{title}」')
    if body:
        lines.append(body)
    lines += ['', '詳細はこちら', display_url, '[/info]']
    return '\n'.join(lines)


def notify_chatwork(product_name, rv, display_url):
    msg  = build_alert_message(product_name, rv, display_url)
    hdrs = {'X-ChatWorkToken': CHATWORK_TOKEN}
    r = requests.post(
        f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM}/messages',
        headers=hdrs, data={'body': msg}, timeout=30,
    )
    print(f'    Chatwork送信: {r.status_code}')
    return r.status_code in (200, 201)


# ── メイン ────────────────────────────────────────────────────────

def main():
    if not is_business_day():
        today = datetime.now(JST)
        print(f'本日（{today:%Y-%m-%d}）は稼働対象外のためスキップします')
        return

    since_dt = datetime.now(JST) - timedelta(hours=2)
    print(f'チェック対象: {since_dt:%Y-%m-%d %H:%M} 以降の低評価（★1-2）レビュー\n')

    gc       = setup_gspread()
    ws_alert = ensure_alert_sheet(gc)
    products = load_products(gc)
    notified = load_notified(gc)
    print(f'{len(products)} 商品を処理します\n')

    total_sent = 0
    for p in products:
        name       = p['name']
        review_url = p['review_url']
        print(f'▶ {name}')

        if not review_url:
            print('  レビューURLが未設定')
            continue

        reviews = scrape_low_reviews(review_url, since_dt, notified)

        if not reviews:
            print('  新着低評価レビューなし')
            continue

        display_url = build_display_url(review_url)
        for rv in reviews:
            print(f'  ★{rv["rating"]} 検出 → 通知送信')
            if notify_chatwork(name, rv, display_url):
                save_notified(ws_alert, name, rv)
                notified.add(rv['hash'])
                total_sent += 1
            time.sleep(0.5)

    if total_sent:
        print(f'\n✅ 合計 {total_sent} 件の低評価レビューを通知しました')
    else:
        print('\n低評価の新着レビューなし')

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
