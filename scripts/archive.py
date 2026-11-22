# -*- coding: utf-8 -*-
# chem99-report 주차별 아카이브 관리 스크립트
#
# 설치 (썸네일 생성용 — 없어도 허브는 만들어지고 썸네일만 빠집니다)
#   pip install playwright pillow
#   playwright install chromium
#
# 사용법 (레포 루트에서 실행)
#   python3 scripts/archive.py add <파일이 있는 폴더>   # 새 주 파일을 archive/YYYY-Wnn/ 에 복사 + 썸네일 생성 + 허브 갱신
#   python3 scripts/archive.py thumbs [--force]         # 없는 썸네일만(--force 면 전부) 다시 생성한 뒤 허브 갱신
#   python3 scripts/archive.py build                    # archive/ 를 다시 훑어 허브(index.html)와 archive/index.json 재생성
#
# 인식하는 파일 이름 (기준 금요일 YYYYMMDD가 이름에 들어 있어야 함)
#   chem99레포트_SM_20261120.html             → 회사양식 HTML 보고서
#   chem99레포트_SM_20261120_전체_S01-S21.pptx → PPT 전체 저장본
#   chem99_SM_번역_20261120.md                → 번역본 (0절 '가격 데이터' 표에서 카드의 가격·등락률을 읽음)
# 주차 폴더 이름은 기준 금요일의 ISO 주차입니다 (예: 2026-11-20 → 2026-W47). 원문 회차(제NN기)와 다릅니다.
# 썸네일은 HTML 보고서의 첫 쪽(핵심지표)을 캡처해 archive/<주차>/thumbs/<품목>.jpg 로 저장합니다.

import datetime
import html
import json
import re
import shutil
import sys
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent.parent          # 레포 루트
ARCHIVE = ROOT / 'archive'                              # 주차별 보관 폴더
PRODUCTS = ['SM', 'AN', 'BD', 'BPA', 'MMA']             # 허브에 표시할 품목 순서
PRODUCT_NAME = {                                        # 품목 한글 이름
    'SM': '스티렌모노머', 'AN': '아크릴로니트릴', 'BD': '부타디엔',
    'BPA': '비스페놀A', 'MMA': '메틸메타크릴레이트'
}
THUMB_W = 640                                           # 썸네일 가로 픽셀
# 파일 이름 규칙 (정규식)
RE_HTML = re.compile(r'^chem99레포트_([A-Za-z0-9]+)_(\d{8})\.html$')
RE_PPTX = re.compile(r'^chem99레포트_([A-Za-z0-9]+)_(\d{8})_전체_S\d+-S\d+\.pptx$')
RE_MD = re.compile(r'^chem99_([A-Za-z0-9]+)_번역_(\d{8})\.md$')


def fail(msg):
    """한글 오류 메시지를 출력하고 종료합니다."""
    print('오류: ' + msg)
    sys.exit(1)


def classify(name):
    """파일 이름 → (품목, 기준일 8자리, 종류) 또는 None. 종류는 html / pptx / md."""
    for kind, rx in (('html', RE_HTML), ('pptx', RE_PPTX), ('md', RE_MD)):
        m = rx.match(name)
        if m:
            return m.group(1).upper(), m.group(2), kind
    return None


def iso_week_dir(yyyymmdd):
    """기준 금요일 YYYYMMDD → 'YYYY-Wnn' (ISO 주차)."""
    try:
        d = datetime.date(int(yyyymmdd[:4]), int(yyyymmdd[4:6]), int(yyyymmdd[6:8]))
    except ValueError:
        fail('날짜 형식이 올바르지 않습니다: ' + yyyymmdd)
    y, w, _ = d.isocalendar()
    return '%d-W%02d' % (y, w), d


# ------------------------------------------------------------------ 새 주 추가
def add_week(src):
    """폴더 안의 보고서 파일을 읽어 archive/YYYY-Wnn/ 으로 복사한 뒤 썸네일·허브를 갱신합니다."""
    src = Path(src)
    if not src.is_dir():
        fail('폴더를 찾을 수 없습니다: ' + str(src))
    found = []
    for f in sorted(src.iterdir()):
        c = classify(f.name) if f.is_file() else None
        if c:
            found.append((f, c))
    if not found:
        fail('인식할 수 있는 보고서 파일이 없습니다. 파일 이름 규칙은 이 스크립트 맨 위 설명을 확인하세요.')
    weeks = {}
    for f, (prod, ymd, kind) in found:
        wk, _ = iso_week_dir(ymd)
        weeks.setdefault(wk, []).append(f)
    for wk, files in sorted(weeks.items()):
        dest = ARCHIVE / wk
        dest.mkdir(parents=True, exist_ok=True)
        for f in files:
            shutil.copy2(f, dest / f.name)              # 같은 이름이면 덮어씀(재작성 반영)
            print('복사: %s → archive/%s/' % (f.name, wk))
    make_thumbs(force=False, only_weeks=set(weeks))      # 새로 들어온 주차의 썸네일
    build_index()


# ------------------------------------------------------------------ 썸네일
def make_thumbs(force=False, only_weeks=None):
    """HTML 보고서 첫 쪽을 캡처해 thumbs/<품목>.jpg 로 저장합니다. playwright·pillow 가 없으면 건너뜁니다."""
    try:
        from playwright.sync_api import sync_playwright
        from PIL import Image
    except ImportError:
        print('안내: playwright/pillow 가 설치되어 있지 않아 썸네일 생성을 건너뜁니다. (pip install playwright pillow)')
        return
    jobs = []
    for w in scan():
        if only_weeks and w['week'] not in only_weeks:
            continue
        for prod, p in w['products'].items():
            if 'html' not in p:
                continue
            out = ARCHIVE / w['week'] / 'thumbs' / (prod + '.jpg')
            if force or not out.exists():
                jobs.append((ARCHIVE / w['week'] / p['html'], out))
    if not jobs:
        print('썸네일: 새로 만들 항목이 없습니다.')
        return
    tmp = ROOT / '.thumb_tmp.png'
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={'width': 1700, 'height': 1250})
        for src, out in jobs:
            try:
                page.goto(src.resolve().as_uri())
                page.wait_for_timeout(1500)
                page.evaluate('() => { if (typeof gotoTab === "function") gotoTab(0); }')   # 첫 쪽(핵심지표)
                page.wait_for_timeout(400)
                el = page.query_selector('#slidewrap') or page.query_selector('#stage')
                if el is None:
                    print('경고: 슬라이드 영역을 찾지 못했습니다 → ' + src.name)
                    continue
                el.screenshot(path=str(tmp))
                img = Image.open(tmp).convert('RGB')
                h = int(img.height * THUMB_W / img.width)
                out.parent.mkdir(parents=True, exist_ok=True)
                img.resize((THUMB_W, h), Image.LANCZOS).save(out, 'JPEG', quality=82, optimize=True)
                print('썸네일: ' + str(out.relative_to(ROOT)))
            except Exception as e:                       # 한 건 실패해도 나머지는 계속
                print('경고: 썸네일 생성 실패 (%s): %s' % (src.name, e))
        browser.close()
    if tmp.exists():
        tmp.unlink()


# ------------------------------------------------------------------ 번역본에서 카드용 가격 읽기
def md_tables(section):
    """마크다운 절 문자열 → 표 목록(표 = 행 목록, 행 = 칸 문자열 목록)."""
    tabs, cur = [], None
    for line in section.split('\n'):
        t = line.strip()
        if t.startswith('|'):
            if cur is None:
                cur = []
                tabs.append(cur)
            if not re.match(r'^\|\s*-{3}', t):
                cur.append([c.strip() for c in t.strip('|').split('|')])
        else:
            cur = None
    return tabs


def clean_cell(s):
    """굵게 표시·중문 병기 제거."""
    s = s.replace('**', '')
    s = re.sub(r'\s*\([^)]*[一-鿿][^)]*\)', '', s)
    return s.strip()


def read_kpi(md_path):
    """번역본 0절 첫 번째 표(가격 데이터)에서 앞의 두 행 → [{'label','price','pct'}]. 읽지 못하면 빈 목록."""
    try:
        md = Path(md_path).read_text(encoding='utf-8')
        i = md.index('### 0.')
        m = re.search(r'\n###? ', md[i + 5:])
        sec = md[i: i + 5 + m.start()] if m else md[i:]
        tabs = md_tables(sec)
        if not tabs or len(tabs[0]) < 2:
            return []
        head = [clean_cell(c) for c in tabs[0][0]]
        pc = next((k for k, c in enumerate(head) if '평균가' in c), None)       # '평균가' / '본주 평균가'
        if pc is None:
            pc = next((k for k, c in enumerate(head) if c == '본주'), None)      # BD: '본주'
        cc = next((k for k, c in enumerate(head) if '전주 대비' in c), None)
        if pc is None or cc is None:
            return []
        out = []
        for row in tabs[0][1:3]:
            if len(row) <= max(pc, cc):
                continue
            out.append({'label': clean_cell(row[0]), 'price': clean_cell(row[pc]), 'pct': clean_cell(row[cc])})
        return out
    except (OSError, ValueError):
        return []


# ------------------------------------------------------------------ 목록 스캔
def scan():
    """archive/ 아래를 훑어 주차별 목록을 만듭니다. 반환: [{week, friday, products:{품목:{html,pptx,md,thumb,kpi}}}] (최신 주 먼저)."""
    out = []
    if not ARCHIVE.is_dir():
        return out
    for wd in sorted(ARCHIVE.iterdir(), reverse=True):
        if not (wd.is_dir() and re.match(r'^\d{4}-W\d{2}$', wd.name)):
            continue
        prods, friday = {}, None
        for f in sorted(wd.iterdir()):
            c = classify(f.name) if f.is_file() else None
            if not c:
                continue
            prod, ymd, kind = c
            friday = friday or ymd
            prods.setdefault(prod, {})[kind] = f.name
        for prod, p in prods.items():
            if (wd / 'thumbs' / (prod + '.jpg')).exists():
                p['thumb'] = 'thumbs/' + prod + '.jpg'
            if 'md' in p:
                p['kpi'] = read_kpi(wd / p['md'])
        if prods:
            out.append({'week': wd.name, 'friday': friday, 'products': prods})
    return out


# ------------------------------------------------------------------ 허브 렌더링
def pct_class(pct):
    """등락률 문자열 → 색 클래스 (상승 빨강 / 하락 파랑 / 그 외 없음)."""
    s = pct.strip()
    if re.match(r'^\+\d', s):
        return 'up'
    if re.match(r'^-\d', s):
        return 'dn'
    return ''


def href(week, rel):
    """상대 경로 링크(한글 파일 이름은 퍼센트 인코딩)."""
    return 'archive/%s/%s' % (week, '/'.join(quote(x) for x in rel.split('/')))


def render_card(week, prod, p):
    """품목 카드 한 장의 HTML."""
    name = html.escape(PRODUCT_NAME.get(prod, prod))
    if 'thumb' in p and 'html' in p:
        thumb = '<a class="th" href="%s"><img loading="lazy" src="%s" alt="%s 보고서 첫 쪽"></a>' % (
            href(week, p['html']), href(week, p['thumb']), html.escape(prod))
    elif 'html' in p:
        thumb = '<a class="th noimg" href="%s"><span>%s</span></a>' % (href(week, p['html']), html.escape(prod))
    else:
        thumb = '<div class="th noimg"><span>%s</span></div>' % html.escape(prod)
    kpi = ''
    for k in p.get('kpi', []):
        kpi += '<div class="kr"><span class="kl">%s</span><span class="kv">%s</span><span class="kp %s">%s</span></div>' % (
            html.escape(k['label']), html.escape(k['price']), pct_class(k['pct']), html.escape(k['pct']))
    btn = []
    if 'html' in p:
        btn.append('<a class="b main" href="%s">보고서 열기</a>' % href(week, p['html']))
    if 'pptx' in p:
        btn.append('<a class="b" href="%s">PPT</a>' % href(week, p['pptx']))
    if 'md' in p:
        btn.append('<a class="b" href="%s">번역본</a>' % href(week, p['md']))
    return ('<article class="card" data-prod="%s">%s<div class="cb"><div class="ct"><b>%s</b><span>%s</span></div>'
            '<div class="kpi">%s</div><div class="bt">%s</div></div></article>') % (
        html.escape(prod), thumb, html.escape(prod), name, kpi, ''.join(btn))


def render_hub(weeks):
    """허브 페이지 HTML 문자열. 외부 라이브러리·CDN 없이 단독으로 열립니다."""
    order = lambda p: PRODUCTS.index(p) if p in PRODUCTS else 99
    n_reports = sum(len(w['products']) for w in weeks)
    secs = []
    for idx, w in enumerate(weeks):
        fr = w['friday']
        fr_txt = '%s-%s-%s' % (fr[:4], fr[4:6], fr[6:8])
        cards = ''.join(render_card(w['week'], prod, w['products'][prod]) for prod in sorted(w['products'], key=order))
        label = html.escape(w['week'])
        if idx == 0:
            secs.append(
                '<section class="wk latest" data-week="%s %s"><div class="wh"><h2>%s <em>최신</em></h2>'
                '<span class="fr">기준 금요일 %s · %d개 품목</span></div><div class="grid">%s</div></section>'
                % (label, fr_txt, label, fr_txt, len(w['products']), cards))
        else:
            secs.append(
                '<details class="wk" data-week="%s %s"><summary><b>%s</b><span class="fr">기준 금요일 %s · %d개 품목</span></summary>'
                '<div class="grid">%s</div></details>' % (label, fr_txt, label, fr_txt, len(w['products']), cards))
    chips = '<button class="chip on" data-f="">전체</button>' + ''.join(
        '<button class="chip" data-f="%s">%s</button>' % (p, p) for p in PRODUCTS)
    body = ''.join(secs) if secs else '<p class="empty">아직 보관된 주차가 없습니다.</p>'
    latest = weeks[0]['week'] if weeks else '-'
    page = '''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CHEM99 주간 시황 보고서 아카이브</title>
<style>
  :root{--navy:#0C4270;--line:#d9d9d9;--mute:#7f7f7f;--up:#ff0000;--dn:#0000ff}
  *{box-sizing:border-box}
  body{margin:0;background:#f3f5f8;color:#1a1a1a;font-family:"Malgun Gothic","Apple SD Gothic Neo",Arial,sans-serif}
  .top{background:var(--navy);color:#fff;border-bottom:4px solid #000}
  .top .in{max-width:1180px;margin:0 auto;padding:26px 20px 22px;display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:14px}
  .top h1{margin:0;font-size:27px;letter-spacing:-.3px}
  .top p{margin:6px 0 0;font-size:13px;color:#cfe0ee;line-height:1.6}
  .stats{display:flex;gap:10px}
  .stat{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.25);border-radius:10px;padding:8px 14px;text-align:center;min-width:84px}
  .stat b{display:block;font-size:20px}.stat span{font-size:11px;color:#cfe0ee}
  .wrap{max-width:1180px;margin:0 auto;padding:18px 20px 60px}
  .tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between;margin-bottom:16px}
  .chips{display:flex;gap:6px;flex-wrap:wrap}
  .chip{border:1px solid var(--line);background:#fff;border-radius:999px;padding:6px 14px;font:inherit;font-size:13px;font-weight:700;cursor:pointer;color:#333}
  .chip.on{background:var(--navy);border-color:var(--navy);color:#fff}
  #q{border:1px solid var(--line);border-radius:8px;padding:7px 12px;font:inherit;font-size:13px;width:200px;background:#fff}
  .wk{margin-bottom:14px}
  .wh{display:flex;align-items:baseline;gap:12px;margin:6px 2px 10px}
  .wh h2{margin:0;font-size:21px}
  .wh em{font-style:normal;font-size:11px;background:#ff0000;color:#fff;border-radius:4px;padding:2px 7px;vertical-align:3px;margin-left:6px}
  .fr{color:var(--mute);font-size:12px}
  details.wk{background:#fff;border:1px solid var(--line);border-radius:10px;padding:0 14px}
  details.wk summary{cursor:pointer;padding:13px 0;display:flex;align-items:baseline;gap:12px;font-size:16px}
  details.wk[open] summary{border-bottom:1px solid var(--line);margin-bottom:14px}
  details.wk .grid{padding-bottom:14px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:14px}
  .card{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 1px 3px rgba(0,0,0,.05)}
  .latest .card:hover{box-shadow:0 6px 18px rgba(12,66,112,.18);transform:translateY(-2px);transition:.15s}
  .th{display:block;aspect-ratio:4/3;background:#eef1f5;border-bottom:1px solid var(--line);overflow:hidden}
  .th img{width:100%;height:100%;object-fit:cover;object-position:top left;display:block}
  .th.noimg{display:flex;align-items:center;justify-content:center;color:#a6adb8;font-size:30px;font-weight:700;text-decoration:none}
  .cb{padding:12px 13px 13px;display:flex;flex-direction:column;gap:9px;flex:1}
  .ct{display:flex;align-items:baseline;gap:8px}.ct b{font-size:19px;color:var(--navy)}.ct span{font-size:12px;color:var(--mute)}
  .kpi{min-height:18px}
  .kr{display:flex;align-items:baseline;gap:6px;font-size:12.5px;padding:3px 0;border-bottom:1px dashed #e5e5e5}
  .kl{flex:1;color:#555;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .kv{font-weight:700}.kp{font-weight:700;min-width:54px;text-align:right}
  .up{color:var(--up)}.dn{color:var(--dn)}
  .bt{display:flex;gap:6px;margin-top:auto}
  .b{flex:1;text-align:center;border:1px solid var(--line);border-radius:7px;padding:7px 4px;font-size:12.5px;font-weight:700;color:var(--navy);text-decoration:none;background:#fff}
  .b.main{flex:2;background:var(--navy);border-color:var(--navy);color:#fff}
  .b:hover{filter:brightness(.93)}
  .empty,.none{color:var(--mute);text-align:center;padding:40px 0}
  .foot{margin-top:28px;font-size:12px;color:var(--mute);line-height:1.8;text-align:center}
  .foot a{color:var(--navy);font-weight:700}
  @media(max-width:560px){.top h1{font-size:22px}.grid{grid-template-columns:repeat(2,1fr);gap:10px}.ct b{font-size:16px}.b{font-size:11.5px}}
</style>
</head>
<body>
<header class="top"><div class="in">
  <div><h1>CHEM99 주간 시황 보고서 아카이브</h1>
  <p>卓创资讯(SCI99) 주보 번역 + 회사양식 HTML·PPT 보고서를 주차별(ISO 주차)로 보관합니다.<br>주차는 기준 금요일 기준이며 원문 회차(제NN기)와 다릅니다. 가격은 위안/톤, 등락률은 전주 대비입니다.</p></div>
  <div class="stats"><div class="stat"><b>@@latest@@</b><span>최신 주차</span></div><div class="stat"><b>@@nw@@</b><span>보관 주차</span></div><div class="stat"><b>@@nr@@</b><span>보고서</span></div></div>
</div></header>
<main class="wrap">
  <div class="tools"><div class="chips">@@chips@@</div><input id="q" type="search" placeholder="주차·날짜 검색 (예: W47, 2026-11)"></div>
  @@body@@
  <p class="none" id="none" hidden>조건에 맞는 항목이 없습니다.</p>
  <div class="foot">이 페이지는 <code>python3 scripts/archive.py build</code> 로 자동 생성됩니다 (직접 수정하지 마세요).<br>기존 War Room 대시보드는 <a href="warroom/index.html">warroom/</a> 에 있습니다.</div>
</main>
<script>
  // 품목 필터·주차 검색 (화면 표시만 바꾸며 데이터는 그대로)
  var chips = document.querySelectorAll('.chip'), q = document.getElementById('q'), none = document.getElementById('none');
  var curF = '';
  function apply() {
    var kw = q.value.trim().toLowerCase(), any = false;
    document.querySelectorAll('.wk').forEach(function (w) {
      var okW = !kw || (w.dataset.week || '').toLowerCase().indexOf(kw) >= 0, shown = 0;
      w.querySelectorAll('.card').forEach(function (c) {
        var ok = okW && (!curF || c.dataset.prod === curF);
        c.style.display = ok ? '' : 'none';
        if (ok) shown++;
      });
      w.style.display = shown ? '' : 'none';
      if (shown) any = true;
      if (shown && w.tagName === 'DETAILS' && (kw || curF)) w.open = true;   // 필터 중에는 접힌 주차도 펼침
    });
    none.hidden = any;
  }
  chips.forEach(function (b) {
    b.addEventListener('click', function () {
      chips.forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on'); curF = b.dataset.f; apply();
    });
  });
  q.addEventListener('input', apply);
</script>
</body>
</html>
'''
    for k, v in (('latest', html.escape(latest)), ('nw', str(len(weeks))), ('nr', str(n_reports)), ('chips', chips), ('body', body)):
        page = page.replace('@@%s@@' % k, v)                     # CSS 의 % 문자와 충돌하지 않도록 치환 방식 사용
    return page


def build_index():
    """허브(index.html)와 archive/index.json 을 다시 만듭니다."""
    weeks = scan()
    ARCHIVE.mkdir(exist_ok=True)
    (ARCHIVE / 'index.json').write_text(json.dumps(weeks, ensure_ascii=False, indent=2), encoding='utf-8')
    (ROOT / 'index.html').write_text(render_hub(weeks), encoding='utf-8')
    print('허브 갱신 완료: %d개 주차' % len(weeks))


def main(argv):
    if len(argv) >= 2 and argv[1] == 'build':
        build_index()
    elif len(argv) >= 2 and argv[1] == 'thumbs':
        make_thumbs(force='--force' in argv)
        build_index()
    elif len(argv) >= 3 and argv[1] == 'add':
        add_week(argv[2])
    else:
        print(__doc__)


if __name__ == '__main__':
    main(sys.argv)
