# -*- coding: utf-8 -*-
# chem99-report 주차별 아카이브 관리 스크립트 (허브 = 대표 보고서 뷰어 + 주차별 보관함)
#
# 설치: 추가 라이브러리 없이 파이썬 3 표준 라이브러리만 사용합니다. (pip install 불필요)
#
# 사용법 (레포 루트에서 실행)
#   python3 scripts/archive.py add <파일이 있는 폴더>   # 새 주 파일을 archive/YYYY-Wnn/ 에 복사한 뒤 허브 갱신
#   python3 scripts/archive.py build                    # archive/ 를 다시 훑어 허브(index.html)와 archive/index.json 재생성
#
# 인식하는 파일 이름 (기준 금요일 YYYYMMDD가 이름에 들어 있어야 함)
#   chem99레포트_SM_20261120.html             → 회사양식 HTML 보고서 (허브 뷰어에 표시되는 파일)
#   chem99레포트_SM_20261120_전체_S01-S21.pptx → PPT 전체 저장본
#   chem99_SM_번역_20261120.md                → 번역본 (0절 '가격 데이터' 표에서 허브의 가격·등락률을 읽음)
# 주차 폴더 이름은 기준 금요일의 ISO 주차입니다 (예: 2026-11-20 → 2026-W47). 원문 회차(제NN기)와 다릅니다.
#
# 허브 동작 요약
#   - 맨 위 뷰어가 선택한 품목(기본 SM)·주차의 HTML 보고서를 iframe 으로 직접 보여줍니다.
#   - [▶ 바로 발표] 는 보고서 안의 발표 모드(startPresent)를, [▣ 전체보기] 는 showOverview 를 그대로 호출합니다.
#   - 보고서 파일은 수정하지 않습니다. 같은 사이트(GitHub Pages) 안에서만 연동되며,
#     파일을 직접 더블클릭(file://)해서 연 경우에는 새 창 열기로 대신합니다.

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
PRODUCTS = ['SM', 'AN', 'BD', 'BPA', 'MMA']             # 허브에 표시할 품목 순서 (맨 앞 품목이 기본 대표 보고서)
PRODUCT_NAME = {                                        # 품목 한글 이름
    'SM': '스티렌모노머', 'AN': '아크릴로니트릴', 'BD': '부타디엔',
    'BPA': '비스페놀A', 'MMA': '메틸메타크릴레이트'
}
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
    """폴더 안의 보고서 파일을 읽어 archive/YYYY-Wnn/ 으로 복사한 뒤 허브를 갱신합니다."""
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
    build_index()


# ------------------------------------------------------------------ 번역본에서 허브용 가격 읽기
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
    """archive/ 아래를 훑어 주차별 목록을 만듭니다. 반환: [{week, friday, products:{품목:{html,pptx,md,kpi}}}] (최신 주 먼저)."""
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


def render_row(week, prod, p):
    """주차별 보관함의 품목 한 줄 HTML."""
    name = html.escape(PRODUCT_NAME.get(prod, prod))
    kpi = ''
    for k in p.get('kpi', []):                          # 주요 시장 2곳의 가격·등락률
        kpi += '<span class="kx"><i>%s</i> <b>%s</b> <em class="%s">%s</em></span>' % (
            html.escape(k['label']), html.escape(k['price']), pct_class(k['pct']), html.escape(k['pct']))
    acts = []
    if 'html' in p:
        acts.append('<button class="b main view" data-w="%s" data-p="%s">화면에서 보기</button>' % (html.escape(week), html.escape(prod)))
        acts.append('<a class="b" target="_blank" rel="noopener" href="%s">새 창</a>' % href(week, p['html']))
    if 'pptx' in p:
        acts.append('<a class="b" href="%s">PPT</a>' % href(week, p['pptx']))
    else:
        acts.append('<span class="b off" title="이 주차에는 보관된 PPT 파일이 없습니다 (보고서 안의 [PPT 전체] 버튼으로 저장할 수 있습니다)">PPT 없음</span>')
    if 'md' in p:
        acts.append('<a class="b" href="%s">번역본</a>' % href(week, p['md']))
    return ('<div class="row" data-prod="%s" data-w="%s"><span class="rp"><b>%s</b><small>%s</small></span>'
            '<span class="rk">%s</span><span class="ra">%s</span></div>') % (
        html.escape(prod), html.escape(week), html.escape(prod), name, kpi, ''.join(acts))


def render_hub(weeks):
    """허브 페이지 HTML 문자열. 외부 라이브러리·CDN 없이 단독으로 열립니다."""
    order = lambda p: PRODUCTS.index(p) if p in PRODUCTS else 99
    n_reports = sum(len(w['products']) for w in weeks)
    secs = []
    for idx, w in enumerate(weeks):                     # 주차별 보관함 (최신 주는 펼침)
        fr = w['friday']
        fr_txt = '%s-%s-%s' % (fr[:4], fr[4:6], fr[6:8])
        rows = ''.join(render_row(w['week'], prod, w['products'][prod]) for prod in sorted(w['products'], key=order))
        secs.append(
            '<details class="wk"%s data-week="%s %s"><summary><b>%s</b>%s<span class="fr">기준 금요일 %s · %d개 품목</span></summary>'
            '<div class="rows">%s</div></details>' % (
                ' open' if idx == 0 else '', html.escape(w['week']), fr_txt, html.escape(w['week']),
                '<em class="new">최신</em>' if idx == 0 else '', fr_txt, len(w['products']), rows))
    body = ''.join(secs) if secs else '<p class="empty">아직 보관된 주차가 없습니다.</p>'
    vchips = ''.join('<button class="vchip" data-p="%s">%s</button>' % (p, p) for p in PRODUCTS)
    filt = '<option value="">전체 품목</option>' + ''.join('<option value="%s">%s</option>' % (p, p) for p in PRODUCTS)
    latest = weeks[0]['week'] if weeks else '-'
    data = json.dumps({'weeks': weeks, 'order': PRODUCTS, 'names': PRODUCT_NAME}, ensure_ascii=False).replace('</', '<\\/')
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
  .top .in{max-width:1180px;margin:0 auto;padding:22px 20px 18px;display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:14px}
  .top h1{margin:0;font-size:26px;letter-spacing:-.3px}
  .top p{margin:6px 0 0;font-size:12.5px;color:#cfe0ee;line-height:1.6}
  .stats{display:flex;gap:10px}
  .stat{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.25);border-radius:10px;padding:7px 14px;text-align:center;min-width:84px}
  .stat b{display:block;font-size:19px}.stat span{font-size:11px;color:#cfe0ee}
  .wrap{max-width:1180px;margin:0 auto;padding:16px 20px 60px}
  /* ---- 대표 보고서 뷰어 ---- */
  .viewer{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(12,66,112,.10)}
  .vh{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:10px 12px;background:#fff;border-bottom:1px solid var(--line)}
  .vlab{font-size:12px;color:var(--mute);margin-right:2px}
  .vchip{border:1px solid var(--line);background:#fff;border-radius:999px;padding:5px 14px;font:inherit;font-size:13px;font-weight:700;cursor:pointer;color:#333}
  .vchip.on{background:var(--navy);border-color:var(--navy);color:#fff}
  .vchip:disabled{opacity:.35;cursor:default}
  select{border:1px solid var(--line);border-radius:8px;padding:6px 8px;font:inherit;font-size:13px;background:#fff}
  .sp{flex:1}
  .va{display:flex;flex-wrap:wrap;gap:6px;padding:9px 12px;background:#f7f9fc;border-bottom:1px solid var(--line)}
  .b{display:inline-block;text-align:center;border:1px solid var(--line);border-radius:7px;padding:7px 14px;font:inherit;font-size:13px;font-weight:700;color:var(--navy);text-decoration:none;background:#fff;cursor:pointer}
  .b.main{background:var(--navy);border-color:var(--navy);color:#fff}
  .b.big{padding:8px 18px;font-size:14px}
  .b:hover{filter:brightness(.93)}
  .b.off,.b:disabled{color:#a6adb8;background:#f1f3f6;cursor:default;filter:none}
  .vk{display:flex;flex-wrap:wrap;gap:6px 18px;align-items:baseline;padding:7px 14px;font-size:13px;border-bottom:1px solid var(--line);background:#fff;min-height:34px}
  .vk .tt{font-weight:700;color:var(--navy)}
  .kx i{font-style:normal;color:#555}.kx b{margin:0 2px}.kx em{font-style:normal;font-weight:700}
  .up{color:var(--up)}.dn{color:var(--dn)}
  #frame{display:block;width:100%;height:clamp(620px,calc(100vh - 210px),880px);border:0;background:#fff}
  .vf{padding:7px 14px;font-size:11.5px;color:var(--mute);background:#f7f9fc;border-top:1px solid var(--line);line-height:1.6}
  /* ---- 주차별 보관함 ---- */
  .lh{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:22px 2px 10px}
  .lh h2{margin:0;font-size:19px}
  #q{border:1px solid var(--line);border-radius:8px;padding:6px 12px;font:inherit;font-size:13px;width:200px;background:#fff}
  details.wk{background:#fff;border:1px solid var(--line);border-radius:10px;padding:0 14px;margin-bottom:10px}
  details.wk summary{cursor:pointer;padding:12px 0;display:flex;align-items:baseline;gap:10px;font-size:16px}
  details.wk[open] summary{border-bottom:1px solid var(--line)}
  .new{font-style:normal;font-size:11px;background:#ff0000;color:#fff;border-radius:4px;padding:2px 7px}
  .fr{color:var(--mute);font-size:12px}
  .row{display:grid;grid-template-columns:150px 1fr auto;gap:10px;align-items:center;padding:9px 0;border-bottom:1px dashed #e5e5e5}
  .row:last-child{border-bottom:0}
  .row.cur{background:#eef4fb;margin:0 -14px;padding-left:14px;padding-right:14px}
  .rp b{font-size:17px;color:var(--navy);margin-right:6px}.rp small{font-size:11.5px;color:var(--mute)}
  .rk{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:12.5px}
  .ra{display:flex;gap:5px;flex-wrap:wrap;justify-content:flex-end}
  .ra .b{padding:5px 10px;font-size:12px}
  .empty,.none{color:var(--mute);text-align:center;padding:40px 0}
  .foot{margin-top:26px;font-size:12px;color:var(--mute);line-height:1.8;text-align:center}
  .foot a{color:var(--navy);font-weight:700}
  #toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%);background:#0C4270;color:#fff;padding:10px 18px;border-radius:999px;font-size:13px;box-shadow:0 4px 14px rgba(0,0,0,.25);display:none;z-index:9;max-width:92vw;text-align:center}
  @media(max-width:700px){
    .top h1{font-size:21px}
    #frame{height:78vh;min-height:520px}
    .row{grid-template-columns:1fr;gap:6px}.ra{justify-content:flex-start}
    .b.big{flex:1 1 40%}
  }
</style>
</head>
<body>
<header class="top"><div class="in">
  <div><h1>CHEM99 주간 시황 보고서 아카이브</h1>
  <p>卓创资讯(SCI99) 주보 번역 + 회사양식 HTML·PPT 보고서를 주차별(ISO 주차)로 보관합니다.<br>주차는 기준 금요일 기준이며 원문 회차(제NN기)와 다릅니다. 가격은 위안/톤, 등락률은 전주 대비입니다.</p></div>
  <div class="stats"><div class="stat"><b>@@latest@@</b><span>최신 주차</span></div><div class="stat"><b>@@nw@@</b><span>보관 주차</span></div><div class="stat"><b>@@nr@@</b><span>보고서</span></div></div>
</div></header>
<main class="wrap">
  <section class="viewer" id="viewer">
    <div class="vh"><span class="vlab">대표 보고서</span>@@vchips@@<span class="sp"></span><span class="vlab">주차</span><select id="vweek"></select></div>
    <div class="va">
      <button class="b main big" id="bPres" disabled>▶ 바로 발표</button>
      <button class="b main big" id="bOv" disabled>▣ 전체보기</button>
      <a class="b big" id="bNew" target="_blank" rel="noopener" href="#">↗ 새 창으로 열기</a>
      <span class="sp"></span>
      <a class="b big" id="bPpt" href="#">PPT</a>
      <a class="b big" id="bMd" href="#">번역본</a>
    </div>
    <div class="vk" id="vk"></div>
    <iframe id="frame" title="대표 보고서" allow="fullscreen" allowfullscreen></iframe>
    <div class="vf">발표 중 키: →/Space 다음 · ← 이전 · V 전체보기 · S 단계 등장 · T 전환 · B/W 화면 가림 · L 레이저 · 숫자+Enter 쪽 이동 · Esc 종료 &nbsp;|&nbsp; 전체화면은 브라우저 정책상 버튼을 눌러야 켜지며, 아이폰 Safari 는 창 안 발표로 표시됩니다.</div>
  </section>

  <div class="lh"><h2>주차별 보관함</h2><span class="sp"></span>
    <select id="fp">@@filt@@</select>
    <input id="q" type="search" placeholder="주차·날짜 검색 (예: W47, 2026-11)"></div>
  @@body@@
  <p class="none" id="none" hidden>조건에 맞는 항목이 없습니다.</p>
  <div class="foot">이 페이지는 <code>python3 scripts/archive.py build</code> 로 자동 생성됩니다 (직접 수정하지 마세요).<br>기존 War Room 대시보드는 <a href="warroom/index.html">warroom/</a> 에 있습니다.</div>
</main>
<div id="toast"></div>
<script type="application/json" id="data">@@data@@</script>
<script>
  // 허브 동작: 대표 보고서 뷰어(iframe) 전환·발표·전체보기, 보관함 필터·검색
  var DATA = JSON.parse(document.getElementById('data').textContent);       // 스크립트가 심어 둔 주차·품목 목록
  var frame = document.getElementById('frame'), toastEl = document.getElementById('toast');
  var bPres = document.getElementById('bPres'), bOv = document.getElementById('bOv');
  var cur = { p: '', w: '' };                                                // 현재 뷰어에 표시 중인 품목·주차

  function toast(msg, ms) {                                                  // 하단 안내 문구
    toastEl.textContent = msg; toastEl.style.display = 'block';
    clearTimeout(toast.t); toast.t = setTimeout(function () { toastEl.style.display = 'none'; }, ms || 3500);
  }
  function enc(rel) { return rel.split('/').map(encodeURIComponent).join('/'); }   // 한글 파일 이름 인코딩
  function url(w, file) { return 'archive/' + w + '/' + enc(file); }
  function findWeek(w) { for (var i = 0; i < DATA.weeks.length; i++) if (DATA.weeks[i].week === w) return DATA.weeks[i]; return null; }
  function weeksOf(p) {                                                      // 해당 품목 HTML 이 있는 주차들(최신 먼저)
    return DATA.weeks.filter(function (x) { return x.products[p] && x.products[p].html; });
  }
  function pctCls(s) { return /^\\+\\d/.test(s) ? 'up' : (/^-\\d/.test(s) ? 'dn' : ''); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  function setLink(a, href, ok) {                                            // 링크 버튼 활성/비활성
    if (ok) { a.href = href; a.classList.remove('off'); a.removeAttribute('title'); }
    else { a.removeAttribute('href'); a.classList.add('off'); }
  }
  function show(p, w, scroll) {                                              // 뷰어에 품목·주차 표시
    var list = weeksOf(p);
    if (!list.length) { toast(p + ' 보고서가 아직 없습니다.'); return; }
    var wk = findWeek(w) && findWeek(w).products[p] && findWeek(w).products[p].html ? w : list[0].week;
    var item = findWeek(wk).products[p];
    document.querySelectorAll('.vchip').forEach(function (b) { b.classList.toggle('on', b.dataset.p === p); });
    var sel = document.getElementById('vweek');
    sel.innerHTML = list.map(function (x) { return '<option value="' + x.week + '">' + x.week + ' (' + x.friday.slice(0, 4) + '-' + x.friday.slice(4, 6) + '-' + x.friday.slice(6) + ')</option>'; }).join('');
    sel.value = wk;
    var changed = cur.p !== p || cur.w !== wk;
    cur = { p: p, w: wk };
    var u = url(wk, item.html);
    document.getElementById('bNew').href = u;
    setLink(document.getElementById('bPpt'), item.pptx ? url(wk, item.pptx) : '', !!item.pptx);
    if (!item.pptx) document.getElementById('bPpt').title = '이 주차에는 보관된 PPT 파일이 없습니다 (보고서 안의 [PPT 전체] 버튼으로 저장)';
    setLink(document.getElementById('bMd'), item.md ? url(wk, item.md) : '', !!item.md);
    var k = (item.kpi || []).map(function (x) { return '<span class="kx"><i>' + esc(x.label) + '</i> <b>' + esc(x.price) + '</b> <em class="' + pctCls(x.pct) + '">' + esc(x.pct) + '</em></span>'; }).join('');
    document.getElementById('vk').innerHTML = '<span class="tt">' + p + ' · ' + DATA.names[p] + ' · ' + wk + '</span>' + k;
    document.querySelectorAll('.row').forEach(function (r) { r.classList.toggle('cur', r.dataset.prod === p && r.dataset.w === wk); });
    if (changed) {                                                           // 같은 보고서면 다시 불러오지 않음
      bPres.disabled = true; bOv.disabled = true; bOv.textContent = '▣ 전체보기';
      frame.src = u;
    }
    try { history.replaceState(null, '', '?p=' + p + '&w=' + wk); } catch (e) {}
    if (scroll) document.getElementById('viewer').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  frame.addEventListener('load', function () {                               // 보고서가 열리면 버튼 활성화, 쪽 수 표시
    if (!frame.getAttribute('src')) return;
    bPres.disabled = false; bOv.disabled = false;
    try {
      var n = frame.contentWindow.eval('SLIDES.length');                      // 보고서 안의 쪽 수
      if (n) bOv.textContent = '▣ 전체보기(' + n + '쪽)';
    } catch (e) {}
  });
  function fallbackOpen(msg) {                                               // 연동이 막힌 경우(file:// 등) 새 창으로 대신
    window.open(document.getElementById('bNew').href, '_blank');
    toast(msg, 5000);
  }
  function runInReport(fn, label) {                                          // 보고서 안의 함수 호출 (같은 사이트일 때만 가능)
    var w;
    try { w = frame.contentWindow; if (typeof w[fn] !== 'function') throw new Error('x'); }
    catch (e) { fallbackOpen('연동이 막혀 새 창으로 열었습니다. 새 창에서 ' + label + ' 버튼을 눌러 주세요.'); return; }
    frame.focus(); w.focus();                                                // 방향키·Esc 가 보고서로 전달되도록 포커스 이동
    w[fn]();
  }
  bPres.addEventListener('click', function () { runInReport('startPresent', '[▶ 발표(F5)]'); });   // 클릭 1회로 전체화면 발표
  bOv.addEventListener('click', function () { runInReport('showOverview', '[▣ 전체보기]'); });
  document.querySelectorAll('.vchip').forEach(function (b) {                  // 품목 전환 (같은 주차 우선)
    if (!weeksOf(b.dataset.p).length) b.disabled = true;
    b.addEventListener('click', function () { show(b.dataset.p, cur.w, false); });
  });
  document.getElementById('vweek').addEventListener('change', function (e) { show(cur.p, e.target.value, false); });
  document.querySelectorAll('.view').forEach(function (b) {                   // 보관함의 [화면에서 보기]
    b.addEventListener('click', function () { show(b.dataset.p, b.dataset.w, true); });
  });

  // 보관함 필터·검색 (화면 표시만 바꾸며 데이터는 그대로)
  var q = document.getElementById('q'), fp = document.getElementById('fp'), none = document.getElementById('none');
  function apply() {
    var kw = q.value.trim().toLowerCase(), f = fp.value, any = false;
    document.querySelectorAll('details.wk').forEach(function (w) {
      var okW = !kw || (w.dataset.week || '').toLowerCase().indexOf(kw) >= 0, shown = 0;
      w.querySelectorAll('.row').forEach(function (r) {
        var ok = okW && (!f || r.dataset.prod === f);
        r.style.display = ok ? '' : 'none';
        if (ok) shown++;
      });
      w.style.display = shown ? '' : 'none';
      if (shown) any = true;
      if (shown && (kw || f)) w.open = true;                                  // 필터 중에는 접힌 주차도 펼침
    });
    none.hidden = any;
  }
  q.addEventListener('input', apply); fp.addEventListener('change', apply);

  // 처음 표시: 주소의 ?p=SM&w=2026-W47 이 있으면 그것을, 없으면 첫 품목(SM)의 최신 주차
  var qs = new URLSearchParams(location.search), p0 = qs.get('p');
  if (!p0 || !weeksOf(p0).length) p0 = DATA.order.filter(function (x) { return weeksOf(x).length; })[0];
  if (p0) show(p0, qs.get('w') || '', false);
</script>
</body>
</html>
'''
    for k, v in (('latest', html.escape(latest)), ('nw', str(len(weeks))), ('nr', str(n_reports)),
                 ('vchips', vchips), ('filt', filt), ('body', body), ('data', data)):
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
    elif len(argv) >= 3 and argv[1] == 'add':
        add_week(argv[2])
    else:
        print('사용법:\n  python3 scripts/archive.py add <파일이 있는 폴더>   # 새 주 파일 복사 + 허브 갱신\n  python3 scripts/archive.py build                   # 허브(index.html) 재생성')


if __name__ == '__main__':
    main(sys.argv)
