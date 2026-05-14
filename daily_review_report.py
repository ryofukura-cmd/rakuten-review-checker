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
COMPANY_NAME   = os.environ.get('COMPANY_NAME', 'さくらフォレスト株式会社')

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
        ws.update([['商品名', '商品URL（サムネ用）', 'レビューURL', '監視(ON/OFF)']], 'A1:D1')
        ws.format('A1:D1', {'textFormat': {'bold': True}})
        print('シート「商品リスト」を作成しました')

    if '通知済み' not in existing:
        ws = sh.add_worksheet('通知済み', rows=5000, cols=10)
    else:
        ws = sh.worksheet('通知済み')
    ws.update([['商品名', 'レビュー日付', '評価', 'タイトル', '本文', '投稿者名', '性別', '年齢', '通知日時', 'レビューハッシュ']], 'A1:J1')
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


def save_notified_batch(ws, product_name, reviews):
    """通知済みシートに複数行をまとめて追記"""
    now = datetime.now(JST).strftime('%Y-%m-%d %H:%M')
    rows = []
    for rv in reviews:
        rating = rv.get('rating', 0)
        stars  = (f'{"★" * rating}{"☆" * (5 - rating)} {rating}点') if rating else '未評価'
        rows.append([
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
        ])
    if rows:
        ws.append_rows(rows, value_input_option='USER_ENTERED')


# ── 稼働日チェック ────────────────────────────────────────────────

def is_business_day():
    """土日・祝日・長期休暇期間はFalseを返す"""
    import jpholiday
    today = datetime.now(JST).date()

    # 土曜・日曜
    if today.weekday() >= 5:
        return False

    # 日本の祝日（jpholidayが自動判定）
    if jpholiday.is_holiday(today):
        return False

    # 年末年始（12/28〜1/4）
    if (today.month == 12 and today.day >= 28) or \
       (today.month == 1  and today.day <= 4):
        return False

    # お盆（8/10〜8/16）
    if today.month == 8 and 10 <= today.day <= 16:
        return False

    return True


# ── チェック対象期間 ──────────────────────────────────────────────

def get_check_since():
    return datetime.now(JST) - timedelta(days=2)


# ── 日付パース ────────────────────────────────────────────────────

def parse_date(text):
    if not text:
        return None
    m = re.search(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
    if not m:
        return None
    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), tzinfo=JST)


# ── レビュースクレイピング ─────────────────────────────────────────

def build_scrape_url(review_url):
    """スクレイピング用URL（ページ番号 + sort=6）"""
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    if not re.search(r'/\d+/?$', base):
        base = f'{base}/1'
    return f'{base}/?sort=6'


def build_display_url(review_url):
    """通知表示用URL（ページ番号なし + sort=6）"""
    base = re.sub(r'[?#].*', '', review_url.rstrip('/'))
    base = re.sub(r'/\d+$', '', base)
    return f'{base}?sort=6'


def scrape_reviews(review_url, since_dt, notified):
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
        print(f'  レビュー件数（全体）: {total}, このページ: {len(uuids)}件')

        for uuid in uuids:
            rv = reviews_data.get(uuid, {})
            if not isinstance(rv, dict):
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

            rating     = rv.get('rating', 0)
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


# ── Chatwork通知（全商品まとめて1通）────────────────────────────────

def build_message(product_reviews):
    """
    product_reviews: [{'name': str, 'display_url': str, 'reviews': [...]}, ...]
    """
    SEP = '━━━━━━━━━━━━━━'
    total = sum(len(p['reviews']) for p in product_reviews)

    today = datetime.now(JST)
    weekdays = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
    date_str = f'[ {today.year}年{today.month}月{today.day}日 {weekdays[today.weekday()]} ]'

    lines = [f'[info][title]【{COMPANY_NAME}】楽天市場レビュー通知 - 新しいレビューが {total}件 届きました[/title]', '']
    lines += [date_str, '']
    lines += ['\\ お客様のレビュー速報 /', '']

    # 全体サマリー
    lines += [SEP, '【全体】', SEP, '']
    lines.append('■ 対象商品')
    for p in product_reviews:
        lines.append(f'・{p["name"]}：{len(p["reviews"])}件')
    lines.append('')

    rating_count = {}
    for p in product_reviews:
        for rv in p['reviews']:
            r = rv.get('rating', 0)
            rating_count[r] = rating_count.get(r, 0) + 1

    lines.append('■ 評価内訳')
    for r in sorted(rating_count.keys(), reverse=True):
        stars = '★' * r + '☆' * (5 - r)
        lines.append(f'{stars}：{rating_count[r]}件')
    lines.append('')

    # 商品別レビュー
    for p in product_reviews:
        lines += [SEP, f'【{p["name"]}】', SEP, '']

        for rv in p['reviews']:
            rating = rv.get('rating', 0)
            stars  = '★' * rating + '☆' * (5 - rating)
            lines.append(f'{stars} {rating} / 5')

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
            lines.append(profile)

            title = rv.get('title', '')
            body  = rv.get('body', '')
            if title:
                lines.append(f'『{title}』')
            if body:
                lines.append(f'「{body}」')
            lines.append('')

        lines += ['詳細はこちら', p['display_url'], '']

    lines += ['――――――――――', '[/info]']
    return '\n'.join(lines)


def notify_chatwork(product_reviews):
    msg  = build_message(product_reviews)
    hdrs = {'X-ChatWorkToken': CHATWORK_TOKEN}
    r = requests.post(
        f'https://api.chatwork.com/v2/rooms/{CHATWORK_ROOM}/messages',
        headers=hdrs, data={'body': msg}, timeout=30,
    )
    print(f'  Chatwork送信: {r.status_code} {r.text[:200]}')
    return r.status_code in (200, 201)


# ── メイン ────────────────────────────────────────────────────────

def main():
    if not is_business_day():
        today = datetime.now(JST)
        print(f'本日（{today:%Y-%m-%d %A}）は稼働対象外のためスキップします')
        return

    since_dt = get_check_since()
    print(f'チェック対象: {since_dt:%Y-%m-%d %H:%M} 以降のレビュー\n')

    gc = setup_gspread()
    ensure_sheets(gc)

    products = load_products(gc)
    notified = load_notified(gc)
    print(f'{len(products)} 商品を処理します\n')

    product_reviews = []

    for p in products:
        name       = p['name']
        review_url = p['review_url']
        print(f'▶ {name}')

        if not review_url:
            print('  レビューURLが未設定（スプレッドシートのC列に入力してください）')
            continue

        reviews = scrape_reviews(review_url, since_dt, notified)

        if not reviews:
            print('  新着レビューなし')
            continue

        print(f'  {len(reviews)} 件の新着レビュー')
        product_reviews.append({
            'name':        name,
            'display_url': build_display_url(review_url),
            'reviews':     reviews,
        })

    # 同名商品をマージしてレビュー件数の多い順に並べる
    merged: dict[str, dict] = {}
    for p in product_reviews:
        if p['name'] in merged:
            merged[p['name']]['reviews'].extend(p['reviews'])
        else:
            merged[p['name']] = {**p}
    product_reviews = sorted(merged.values(), key=lambda x: len(x['reviews']), reverse=True)

    if product_reviews:
        total = sum(len(p['reviews']) for p in product_reviews)
        print(f'\n合計 {total} 件を通知送信中...')
        notify_chatwork(product_reviews)

        print('スプレッドシートに記録中...')
        ws_notified = gc.open_by_key(SPREADSHEET_ID).worksheet('通知済み')
        for p in product_reviews:
            save_notified_batch(ws_notified, p['name'], p['reviews'])
            for rv in p['reviews']:
                notified.add(rv['hash'])
    else:
        print('\n新着レビューなし')

    print('\n✅ 完了')


if __name__ == '__main__':
    main()
