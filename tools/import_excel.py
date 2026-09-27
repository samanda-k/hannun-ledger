"""
월별 수입지출 엑셀 → 한눈 가계부(구글 시트) 가져오기

사용법 (PowerShell):
  python tools/import_excel.py "D:/AI자동화/가계부_가져오기/월별수입지출내역_2026.xlsx"          # 확인만 (시트에 안 넣음)
  python tools/import_excel.py "...xlsx" --push                                                  # 시트에 넣기

엑셀 구조 (월 시트 하나에 한 달):
  - 매출: '날짜 | 인원 | 결제구분 | 적요 | 공급가 | 부가세 | 입금액 | 일 합계' 머리줄 아래, 날짜·인원은 그날 첫 줄에만
  - '모으기' 줄 아래: 모으기 내역
  - '지출' → '고정'(비고가 '새출발'이면 새출발기금) / '사업자'(결제구분이 '소모품'이면 소모품) / '개인'
연결 정보는 엑셀과 같은 폴더의 연결정보.txt (URL=..., CODE=...) 에서 읽어요. 화면에 출력하지 않아요.
"""
import argparse
import calendar
import datetime as dt
import hashlib
import json
import os
import re
import sys
import time
import warnings
from collections import defaultdict

import openpyxl
import requests

warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

SECTION_MARKERS = {'고정': 'fix', '사업자': 'biz', '개인': 'per'}


def s(v):
    return '' if v is None else str(v).strip()


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return None


def is_date(v):
    return isinstance(v, (dt.datetime, dt.date))


def dstr(v):
    return v.strftime('%Y-%m-%d')


def parse_month(ws, year, month):
    """한 달 시트 → (entries, visits, check) ; check는 엑셀 안의 소계들"""
    entries, visits = [], {}
    check = {}
    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=8, values_only=True))
    first_day = f'{year}-{month:02d}-01'
    mode = None
    cur_date = None
    for i, r in enumerate(rows, start=1):
        a, b, c, d, e, f, g, h = (list(r) + [None] * 8)[:8]
        A = s(a)
        # 구역 바꾸기
        if A == '날짜' and mode is None:
            mode, cur_date = 'sales', None
            continue
        if A == '모으기':
            mode, cur_date = 'save', None
            check['모으기'] = num(g) or 0
            continue
        if A == '지출':
            mode = 'exp_head'
            continue
        if A == '날짜' and mode in ('exp_head', 'fix', 'biz', 'per'):
            continue
        if A in SECTION_MARKERS and mode in ('exp_head', 'fix', 'biz', 'per'):
            mode, cur_date = SECTION_MARKERS[A], None
            check[A] = num(g) or 0
            continue
        if mode is None:
            continue

        if is_date(a):
            # 지난달 시트를 복사하면서 남은 옛 날짜도 이 시트의 달로 맞춰요 (일은 그대로)
            day = min(a.day, calendar.monthrange(year, month)[1])
            cur_date = f'{year}-{month:02d}-{day:02d}'
        amt = num(g)

        if mode == 'sales':
            if not is_date(a) and a is not None and not s(c) and not s(d):
                # 월 합계 줄 (A칸에 평균 등) → 소계 확인용
                check['매출'] = num(h) or num(g) or 0
                check['방문'] = num(b) or 0
                mode = 'after_sales'
                continue
            if is_date(a):
                n = num(b)
                if n:
                    visits[cur_date] = visits.get(cur_date, 0) + int(round(n))
            if amt and s(c) and cur_date:
                sup = num(e)
                sup = int(round(sup)) if sup else 0
                total = int(round(amt))
                entries.append(dict(d=cur_date, t='in', biz=True, cat=s(d) or '기타', pay=s(c),
                                    sup=sup, vat=(total - sup) if sup else 0, amt=total, memo='', src=f'{ws.title}!{i}'))
            continue

        if mode == 'save':
            if amt and s(d):
                if '빌린' in s(d):
                    cat, memo = '빌린 돈 갚기', s(b) or s(d)
                elif '노란우산' in s(d):
                    cat, memo = '노란우산공제', s(d) + (f' ({s(b)})' if s(b) else '')
                else:
                    cat, memo = '모으기', s(d) + (f' ({s(b)})' if s(b) else '')
                entries.append(dict(d=cur_date or first_day, t='move', biz=False, cat=cat, pay='',
                                    sup=0, vat=0, amt=int(round(amt)), memo=memo, src=f'{ws.title}!{i}'))
            continue

        if mode in ('fix', 'biz', 'per'):
            if not amt or not s(d):
                continue  # 합계 줄·빈 줄
            total = int(round(amt))
            sup = num(e)
            sup = int(round(sup)) if sup else 0
            vat = (total - sup) if sup else 0
            tag, note = s(b), s(h)
            date = cur_date or first_day
            if mode == 'fix':
                if note == '새출발' or '새출발' in s(d):
                    ent = dict(t='move', biz=False, cat='새출발기금', memo=s(d))
                elif tag == '소모품':
                    ent = dict(t='out', biz=True, cat='소모품', memo=s(d))
                else:
                    ent = dict(t='out', biz=True, cat='고정비', memo=s(d) + (f' ({tag})' if tag else ''))
            elif mode == 'biz':
                if tag == '소모품':
                    ent = dict(t='out', biz=True, cat='소모품', memo=s(d))
                else:
                    ent = dict(t='out', biz=True, cat='사업자', memo=s(d) + (f' ({tag})' if tag else ''))
            else:
                cat = {'식대': '식비'}.get(tag, tag) if tag else '개인지출'
                ent = dict(t='out', biz=False, cat=cat, memo=s(d))
            ent.update(d=date, pay='', sup=sup if ent['t'] == 'out' else 0, vat=vat if ent['t'] == 'out' else 0,
                       amt=total, src=f'{ws.title}!{i}')
            entries.append(ent)
    return entries, visits, check


def assign_ids(entries):
    seen = defaultdict(int)
    for e in entries:
        key = '|'.join(str(e[k]) for k in ('d', 't', 'biz', 'cat', 'pay', 'amt', 'memo'))
        seen[key] += 1
        e['id'] = 'xl' + hashlib.sha1(f'{key}|{seen[key]}'.encode()).hexdigest()[:12]
        e['ts'] = int(time.mktime(time.strptime(e['d'], '%Y-%m-%d')) * 1000) + seen[key]


def month_stats(entries, visits, ym):
    r = dict(수입=0, 방문=0, 고정비=0, 소모품=0, 사업자=0, 개인지출=0, 새출발=0, 빌린돈=0, 모으기=0, 건수=0)
    for e in entries:
        if not e['d'].startswith(ym):
            continue
        r['건수'] += 1
        if e['t'] == 'in' and e['biz']:
            r['수입'] += e['amt']
        elif e['t'] == 'out' and e['biz']:
            r[{'고정비': '고정비', '소모품': '소모품'}.get(e['cat'], '사업자')] += e['amt']
        elif e['t'] == 'out':
            r['개인지출'] += e['amt']
        elif e['cat'] == '새출발기금':
            r['새출발'] += e['amt']
        elif e['cat'] in ('모으기', '노란우산공제'):
            r['모으기'] += e['amt']
        elif e['cat'] == '빌린 돈 갚기':
            r['빌린돈'] += e['amt']
    r['방문'] = sum(n for d, n in visits.items() if d.startswith(ym))
    r['지출계'] = r['고정비'] + r['소모품'] + r['사업자']
    r['월급여'] = r['수입'] - r['지출계']
    r['남는돈'] = r['월급여'] - r['개인지출'] - r['새출발'] - r['빌린돈'] - r['모으기']
    r['모으기+빌린돈'] = r['모으기'] + r['빌린돈']
    return r


def total_sheet_values(wb, year):
    """'토탈' 시트의 월별 값 (엑셀이 계산해 둔 값)"""
    ws = wb.worksheets[0]
    labels = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=14, values_only=True):
        lab = s(row[0])
        if lab:
            labels[lab] = [num(v) or 0 for v in row[1:13]]
    pick = lambda *names: next((labels[n] for n in names if n in labels), [0] * 12)
    return {
        '수입': pick(f'{year}년 수입'), '방문': pick(f'{year}년 방문'),
        '고정비': pick('고정비'), '소모품': pick('소모품'), '사업자': pick('사업자'), '지출계': pick('지출 계'),
        '월급여': pick('월 급여'), '개인지출': pick('개인지출'), '새출발': pick('새출발기금'),
        '모으기+빌린돈': pick('모으기'), '빌린돈': pick('빌린돈 합계'), '기타수입': pick('기타 수입'), '남는돈': pick('남는돈'),
    }


def read_conn(folder):
    path = os.path.join(folder, '연결정보.txt')
    conf = {}
    with open(path, encoding='utf-8') as fh:
        for line in fh:
            if '=' in line:
                k, v = line.split('=', 1)
                conf[k.strip()] = v.strip()
    conf['URL'] = re.sub(r'/macros/u/\d+/s/', '/macros/s/', conf.get('URL', ''))
    if not conf.get('URL', '').startswith('https://script.google.com/') or not conf.get('CODE') or conf['CODE'].startswith('<'):
        sys.exit('연결정보.txt 의 URL / CODE 를 확인해 주세요.')
    return conf['URL'], conf['CODE']


def api(url, token, body):
    r = requests.post(url, data=json.dumps({**body, 'token': token}).encode('utf-8'),
                      headers={'Content-Type': 'text/plain;charset=utf-8'}, timeout=180)
    try:
        j = r.json()
    except ValueError:
        sys.exit(f'서버 응답을 읽지 못했어요 (HTTP {r.status_code}). 웹 앱 배포 권한을 확인해 주세요.')
    if not j.get('ok'):
        sys.exit(f"서버 오류: {j.get('error')} {j.get('message', '')}")
    return j


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx')
    ap.add_argument('--push', action='store_true', help='구글 시트에 실제로 넣기')
    _t = dt.date.today()
    ap.add_argument('--until', default=f'{_t.year}-{_t.month:02d}-{calendar.monthrange(_t.year, _t.month)[1]:02d}', help='이 날짜까지만 가져오기 (기본: 이번 달 말일)')
    ap.add_argument('--months', default='', help='가져올 달 (예: 1-9). 비우면 --until 까지 전부')
    ap.add_argument('--skip', action='append', default=[], help='빼고 넣을 것 "월:분류" (예: "7월:빌린 돈 갚기" = 7월 시트의 빌린 돈 갚기 줄 빼기)')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    year = None
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=40, max_col=1, values_only=True):
            if is_date(row[0]):
                year = row[0].year
                break
        if year:
            break
    months = range(1, 13)
    if args.months:
        lo, _, hi = args.months.partition('-')
        months = range(int(lo), int(hi or lo) + 1)

    entries, visits, checks = [], {}, {}
    for m in months:
        name = f'{m}월'
        if name not in wb.sheetnames:
            continue
        e, v, c = parse_month(wb[name], year, m)
        entries += e
        visits.update(v)
        checks[m] = c
    entries = [e for e in entries if e['d'] <= args.until]
    for spec in args.skip:
        sheet, _, cat = spec.partition(':')
        before = len(entries)
        entries = [e for e in entries if not (e['src'].split('!')[0] == sheet and e['cat'] == cat)]
        print(f'빼고 넣기: {sheet} {cat} {before - len(entries)}건')
    visits = {d: n for d, n in visits.items() if d <= args.until}
    assign_ids(entries)
    tot = total_sheet_values(wb, year)

    last_month = max((int(e['d'][5:7]) for e in entries), default=0)
    print(f'\n{year}년 · {args.until} 까지 · 기록 {len(entries)}건 · 방문 기록 {len(visits)}일\n')
    keys = ['수입', '방문', '고정비', '소모품', '사업자', '지출계', '월급여', '개인지출', '새출발', '빌린돈', '모으기+빌린돈', '남는돈']
    head = f"{'월':>3} " + ' '.join(f'{k:>11}' for k in keys)
    print(head.replace('모으기+빌린돈', ' 모으기(빌린돈포함)'))
    diffs = []
    for m in months:
        if m > last_month:
            break
        st = month_stats(entries, visits, f'{year}-{m:02d}')
        extra = tot['기타수입'][m - 1]
        st_left = st['남는돈'] + extra
        vals = dict(st, 남는돈=st_left)
        print(f"{m:>2}월 " + ' '.join(f'{int(vals[k]):>11,}' for k in keys))
        xl = {k: tot[k][m - 1] for k in keys}
        print(f"{'엑셀':>3} " + ' '.join(f'{int(xl[k]):>11,}' for k in keys))
        for k in keys:
            if abs(vals[k] - xl[k]) >= 1:
                diffs.append((m, k, vals[k], xl[k]))
        c = checks.get(m, {})
        sec = {'고정': st['고정비'] + st['새출발'], '사업자': st['소모품'] + st['사업자'], '개인': st['개인지출'], '모으기': st['모으기+빌린돈']}
        for k, v in sec.items():
            if k in c and abs(c[k] - v) >= 1:
                print(f'     ! {m}월 시트의 [{k}] 소계 {int(c[k]):,} ↔ 읽은 합계 {int(v):,}')
        if '매출' in c and abs(c['매출'] - st['수입']) >= 1:
            print(f"     ! {m}월 시트 매출 합계 {int(c['매출']):,} ↔ 읽은 합계 {int(st['수입']):,}")
        if '방문' in c and abs(c['방문'] - st['방문']) >= 1:
            print(f"     ! {m}월 시트 방문 합계 {int(c['방문']):,} ↔ 읽은 합계 {int(st['방문']):,}")
    print('\n다른 칸:', len(diffs))
    for m, k, a, b in diffs:
        print(f'  {m}월 {k}: 가져올 값 {int(a):,} / 엑셀 토탈 {int(b):,} (차이 {int(a - b):+,})')

    if not args.push:
        print('\n(확인만 했어요. 시트에 넣으려면 --push)')
        return

    url, token = read_conn(os.path.dirname(os.path.abspath(args.xlsx)))
    existing = api(url, token, {'action': 'list'})
    have = {e['id'] for e in existing['entries']}
    # 앱에서 이미 입력한 같은 기록(날짜·항목·결제·금액이 같음)은 건너뛰어요
    same = defaultdict(int)
    for e in existing['entries']:
        if not str(e['id']).startswith('xl'):
            same[(e['d'], e['t'], bool(e['biz']), e['cat'], e.get('pay', ''), int(e['amt']))] += 1
    skipped = []
    keep = []
    for e in entries:
        k = (e['d'], e['t'], e['biz'], e['cat'], e['pay'], e['amt'])
        if e['id'] not in have and same[k] > 0:
            same[k] -= 1
            skipped.append(e)
        else:
            keep.append(e)
    entries = keep
    if skipped:
        print(f'앱에서 이미 입력한 기록과 같아서 건너뛴 것: {len(skipped)}건 ' + ', '.join(f"{e['d'][5:]} {e['cat']} {e['amt']:,}" for e in skipped))
    ops = [{'op': 'upsert', 'entry': {k: e[k] for k in ('id', 'd', 't', 'biz', 'cat', 'pay', 'sup', 'vat', 'amt', 'memo', 'ts')}} for e in entries]
    ops += [{'op': 'visits', 'd': d, 'n': n} for d, n in sorted(visits.items())]
    extra_ops = []
    for m in months:
        v = tot['기타수입'][m - 1]
        if v and f'{year}-{m:02d}-01' <= args.until:
            e = dict(id=f'xl-etc-{year}{m:02d}', d=f'{year}-{m:02d}-01', t='in', biz=False, cat='기타 수입', pay='',
                     sup=0, vat=0, amt=int(round(v)), memo='엑셀 토탈 기타 수입', ts=0)
            extra_ops.append({'op': 'upsert', 'entry': e})
    ops += extra_ops
    new = sum(1 for e in entries if e['id'] not in have)
    print(f'\n시트에 넣는 중: 새 기록 {new}건, 이미 있던 기록 {len(entries) - new}건 갱신, 방문 {len(visits)}일, 기타 수입 {len(extra_ops)}건')
    CH = 250
    for i in range(0, len(ops), CH):
        api(url, token, {'action': 'batch', 'ops': ops[i:i + CH]})
        print(f'  {min(i + CH, len(ops))}/{len(ops)}')
    after = api(url, token, {'action': 'list'})
    print(f"완료. 시트의 기록 {len(after['entries'])}건 · 방문 {len(after['visits'])}일")


if __name__ == '__main__':
    main()
