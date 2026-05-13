#!/usr/bin/env python3
"""楽天商品レビュー定期チェッカー"""

import os
import json
import hashlib
import time
import re
import requests
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
        ws = sh.add_worksheet('通知済み', rows=5000, cols=10)
    else:
        ws = sh.worksheet('通知済み')
    ws.update('A1:J1', [['商品名', 'レビュー日付', '評価', 'タイトル', '本文', '投稿者名', '性別', '年齢', '通知日時', 'レビューハッシュ']])
    ws.format('A1:J1', {'textFormat': {'bold': True}})
    if '通知済み' not in existing:
        print('シート「通知済み」を作成しました')


def load_products(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('商品リスト').get_all_values()
    products = []
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        if len(r) >= 4 and r[3].strip().upper() == 'OFF':
            continue
        products.append({
            'name':        r[0].strip(),
            'product_url': r[1].strip() if len(r) > 1 else '',
            'review_url':  r[2].strip() if len(r) > 2 else '',
        })
    return products


def load_notified(gc):
    rows = gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').get_all_values()
    return set(r[9] for r in rows[1:] if len(r) >= 10)


def save_notified(gc, product_name, h, review):
    rating = review.get('rating', 0)
    stars = ('★' * rating + '☆' * (5 - rating)) if rating else '未評価'
    gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み').append_row([
        product_name,
        review.get('date', ''),
        stars,
        review.get('title', ''),
        review.get('body', ''),
        review.get('reviewer_name', ''),
        review.get('gender', ''),
        review.get('age', ''),
        datetime.now(JST).strftime('%Y-%m-%d %H:%M'),
        h,
    ])


# ── チェック対象期間 ──────────────────────────────────────────────

def get_check_since():
    return datetime.now(JST) - timedelta(days=30)


# ── 日付パース ────────────────────────────────────────────────────

def parse_date(text):
    if not text:
        return None
    m = re.search(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
    if not m:
        return None
    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), tzinfo=JST)


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


# ── レビュースクレイピング ─────────────────────────────────────────

def build_review_url(review_url):
    """ページ番号とsort=6（新着順）を付与"""
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    # /1/ が末尾にない場合は追加
    if not re.search(r'/\d+/?$', base):
        base = f'{base}/1'
    return f'{base}/?sort=6'


def scrape_reviews(review_url, since_dt, notified):
    results = []
    url = build_review_url(review_url)
    print(f'  URL: {url}')

    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        html = resp.text

        # window.__INITIAL_STATE__ を抽出
        m = re.search(r'window\.__INITIAL_STATE__\s*=\s*', html)
        if not m:
            print('  __INITIAL_STATE__ が見つかりません')
            return []

        start = m.end()
        end = html.find('</script>', start)
        json_str = html[start:end].strip().rstrip(';')
        data = json.loads(json_str)

        item_reviews = data.get('reviews', {}).get('itemReviews', {})
        print(f'  レビュー件数（全体）: {len(item_reviews)}')

        for key, rv in item_reviews.items():
            if not isinstance(rv, dict):
                continue

            # 日付（YYYY/MM/DD 形式）
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

            rating        = rv.get('rating', 0)
            nickname      = rv.get('nickname', '')
            sex           = rv.get('sex', '')
            gender        = {'male': '男性', 'female': '女性'}.get(sex, sex)
            age_range     = str(rv.get('ageRange', ''))
            age_suffix    = rv.get('ageSuffix', '代')
            age           = f'{age_range}{age_suffix}' if age_range else ''

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


# ── Chatwork通知 ──────────────────────────────────────────────────

def notify_chatwork(product_name, review, thumbnail_url, review_url=''):
    rating = review.get('rating', 0)
    stars  = '★' * rating + '☆' * (5 - rating)
    title  = review.get('title', '')
    body   = review.get('body', '')

    lines = [
        '【楽天市場レビュー通知】',
        '\\お客様よりレビューをいただきました/',
        '',
        '■ 商品名',
        product_name,
        '',
        '■ レビュー点数',
        f'{stars}{rating} / 5',
    ]
    if title:
        lines += ['', '■ レビュータイトル文', title]

    reviewer_name = review.get('reviewer_name', '')
    gender        = review.get('gender', '')
    age           = review.get('age', '')
    if reviewer_name:
        profile = f'{reviewer_name} 様'
        if gender or age:
            profile += f'（{gender}・{age}）' if gender and age else f'（{gender or age}）'
        lines += ['', '■ 投稿者情報', profile]

    lines += [
        '',
        '■ レビュー本文',
        body,
        '',
        '――――――――――',
        '',
        '詳細はレビューページをご確認ください',
        review_url,
    ]
    msg = '\n'.join(lines)

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
            notify_chatwork(name, rv, thumb, review_url)
            save_notified(gc, name, rv['hash'], rv)
            notified.add(rv['hash'])
            time.sleep(0.5)

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
