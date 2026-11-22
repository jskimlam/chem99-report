# -*- coding: utf-8 -*-
# chem99-report 주차별 아카이브 관리 스크립트
# 사용법 (레포 루트에서 실행, 라이브러리 설치 불필요 — 파이썬 3.8 이상 표준 라이브러리만 사용)
#   python3 scripts/archive.py add <파일이 있는 폴더>   # 새 주 파일을 archive/YYYY-Wnn/ 에 복사하고 허브(index.html) 갱신
#   python3 scripts/archive.py build                    # archive/ 폴더를 다시 훑어 허브(index.html)와 archive/index.json 재생성
#
# 인식하는 파일 이름 (기준 금요일 YYYYMMDD가 이름에 들어 있어야 함)
#   chem99레포트_SM_20261120.html            → 회사양식 HTML 보고서
#   chem99레포트_SM_20261120_전체_S01-S21.pptx → PPT 전체 저장본
#   chem99_SM_번역_20261120.md               → 번역본
# 주차 폴더 이름은 기준 금요일의 ISO 주차입니다 (예: 2026-11-20 → 2026-W47). 원문 회차(제NN기)와 다릅니다.

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
    'SM': '스티렌모노머 (SM)', 'AN': '아크릴로니트릴 (AN)', 'BD': '부타디엔 (BD)',
    'BPA': '비스페놀A (BPA)', 'MMA': '메틸메타크릴레이트 (MMA)'
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


def add_week(src):
    """폴더 안의 보고서 파일을 읽어 archive/YYYY-Wnn/ 으로 복사합니다."""
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


def scan():
    """archive/ 아래를 훑어 주차별 목록을 만듭니다. 반환: [{week, friday, products:{품목:{html,pptx,md}}}] (최신 주 먼저)."""
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
        if prods:
            out.append({'week': wd.name, 'friday': friday, 'products': prods})
    return out


def link(week, name, label, dl=False):
    """상대 경로 링크(한글 파일 이름은 퍼센트 인코딩). dl=True 면 download 속성으로 열지 않고 바로 저장."""
    return '<a href="archive/%s/%s"%s>%s</a>' % (week, quote(name), ' download' if dl else '', label)


def render_hub(weeks):
    """허브 페이지 HTML 문자열. 외부 라이브러리·CDN 없이 단독으로 열립니다."""
    order = lambda p: PRODUCTS.index(p) if p in PRODUCTS else 99
    rows = []
    for w in weeks:
        fr = w['friday']
        fr_txt = '%s-%s-%s' % (fr[:4], fr[4:6], fr[6:8])
        trs = []
        for prod in sorted(w['products'], key=order):
            p = w['products'][prod]
            cells = [
                link(w['week'], p['html'], '보고서 열기') if 'html' in p else '<span class="na">-</span>',
                link(w['week'], p['pptx'], 'PPT 받기', True) if 'pptx' in p else '<span class="na">-</span>',
                link(w['week'], p['md'], '번역본') if 'md' in p else '<span class="na">-</span>',
            ]
            trs.append('<tr><td class="pn">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>' % (
                html.escape(PRODUCT_NAME.get(prod, prod)), cells[0], cells[1], cells[2]))
        rows.append(
            '<details class="wk"%s><summary><b>%s</b> <span class="fr">기준 금요일 %s · %d개 품목</span></summary>'
            '<table><thead><tr><th>품목</th><th>HTML</th><th>PPT</th><th>번역(MD)</th></tr></thead><tbody>%s</tbody></table></details>'
            % (' open' if not rows else '', html.escape(w['week']), fr_txt, len(w['products']), ''.join(trs)))
    body = ''.join(rows) if rows else '<p>아직 보관된 주차가 없습니다.</p>'
    return '''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CHEM99 주간 시황 보고서 아카이브</title>
<link rel="icon" type="image/png" sizes="192x192" href="assets/lam-sci-a-192.png">
<link rel="shortcut icon" href="assets/lam-sci-a.ico">
<link rel="apple-touch-icon" href="assets/lam-sci-a-180.png">
<link rel="manifest" href="manifest.webmanifest">
<meta name="theme-color" content="#0C4270">
<style>
  body{margin:0;background:#f5f7fa;color:#1a1a1a;font-family:"Malgun Gothic","Apple SD Gothic Neo",Arial,sans-serif}
  .wrap{max-width:860px;margin:0 auto;padding:28px 18px 60px}
  .logo{display:block;height:46px;width:auto;margin:0 0 12px}
  h1{margin:0 0 6px;font-size:26px}
  .sub{color:#595959;font-size:13px;margin-bottom:22px;line-height:1.6}
  details.wk{background:#fff;border:1px solid #d9d9d9;border-radius:10px;margin-bottom:10px;padding:0 16px}
  summary{cursor:pointer;padding:14px 0;font-size:16px}
  .fr{color:#7f7f7f;font-size:12px;margin-left:8px}
  table{width:100%%;border-collapse:collapse;margin:0 0 14px;font-size:14px}
  th{background:#000;color:#fff;font-weight:700;padding:7px 8px;text-align:left}
  td{padding:8px;border-bottom:1px solid #d9d9d9}
  td.pn{font-weight:700}
  a{color:#0c4270;font-weight:700;text-decoration:none}
  a:hover{text-decoration:underline}
  .na{color:#bfbfbf}
  .foot{margin-top:26px;font-size:12px;color:#7f7f7f;line-height:1.7}
</style>
</head>
<body>
<div class="wrap">
  <img class="logo" src="assets/lam-sci-wordmark-ui.png" alt="LAM SCI">
  <h1>CHEM99 주간 시황 보고서 아카이브</h1>
  <div class="sub">卓创资讯(SCI99) 주보 번역 + 회사양식 HTML 보고서를 주차별(ISO 주차)로 보관합니다. 주차는 기준 금요일 기준이며 원문 회차(제NN기)와 다릅니다.</div>
  %s
  <div class="foot">이 페이지는 <code>python3 scripts/archive.py build</code> 로 자동 생성됩니다 (직접 수정하지 마세요). 기존 War Room 대시보드는 <a href="warroom/index.html">warroom/</a> 에 있습니다.</div>
</div>
</body>
</html>
''' % body


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
        print(__doc__)


if __name__ == '__main__':
    main(sys.argv)
