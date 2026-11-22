# chem99-report

卓创资讯(SCI99) 주간 시황 보고서(SM · AN · BD · BPA · MMA)의 **한국어 번역본 + 회사양식 HTML/PPT**를 주차별로 보관하는 아카이브입니다.
(기존 이름: `petchem-warroom-dashboard`)

## 폴더 구조

```
chem99-report/
├─ index.html                  # 아카이브 허브 (자동 생성 — 직접 수정 금지)
├─ archive/
│  ├─ index.json               # 주차·품목·파일 목록 (자동 생성)
│  └─ 2026-W47/                # 기준 금요일의 ISO 주차 (원문 회차 제NN기와 다름)
│     ├─ chem99레포트_SM_20261120.html
│     ├─ chem99레포트_AN_20261120.html
│     ├─ chem99레포트_AN_20261120_전체_S01-S16.pptx
│     ├─ chem99_AN_번역_20261120.md
│     ├─ thumbs/SM.jpg …       # 허브 카드용 첫 쪽 썸네일 (자동 생성)
│     └─ …
├─ scripts/archive.py          # 새 주 추가 · 허브 재생성
├─ warroom/index.html          # 기존 Petchem War Room 대시보드 (그대로 이동)
└─ .nojekyll                   # GitHub Pages가 파일을 가공하지 않도록
```

## 파일 이름 규칙 (스크립트가 이 이름으로 품목·주차를 인식합니다)

| 종류 | 이름 |
|---|---|
| HTML 보고서 | `chem99레포트_{품목}_{YYYYMMDD}.html` |
| PPT 전체 저장본 | `chem99레포트_{품목}_{YYYYMMDD}_전체_S01-S{N}.pptx` |
| 번역본 | `chem99_{품목}_번역_{YYYYMMDD}.md` |

`{YYYYMMDD}`는 보고 기준 금요일이며, 주차 폴더 이름은 그 날짜의 ISO 주차(`YYYY-Wnn`)로 자동 결정됩니다.

## 매주 추가하는 방법

1. 이번 주 파일을 한 폴더에 모읍니다 (예: `~/chem99_0127/`).
2. 레포 루트에서 실행합니다.
   ```bash
   pip install playwright pillow && playwright install chromium   # 최초 1회 (썸네일 생성용, 없으면 썸네일만 빠짐)
   python3 scripts/archive.py add ~/chem99_0127
   ```
   파일이 `archive/YYYY-Wnn/`으로 복사되고, HTML 첫 쪽 썸네일이 만들어지며, `index.html`·`archive/index.json`이 갱신됩니다. (같은 이름 파일은 덮어씁니다.)
   카드의 가격·등락률은 번역본 0절 "가격 데이터" 표의 앞 두 행에서 자동으로 읽습니다.
3. 커밋·푸시합니다.
   ```bash
   git add -A && git commit -m "2027-W04 주보 추가" && git push
   ```

허브만 다시 만들고 싶으면 `python3 scripts/archive.py build`, 썸네일을 전부 다시 만들려면 `python3 scripts/archive.py thumbs --force`.

## 보기

- GitHub Pages를 켜면(Settings → Pages → Branch: `main` / root) `https://<계정>.github.io/chem99-report/` 에서 허브와 HTML 보고서를 바로 열 수 있습니다.
- Pages 없이 쓰려면 저장소를 내려받아 `index.html`을 브라우저로 열면 됩니다 (CDN·외부 라이브러리 없음).

## 참고

- 보관 파일에는 계정·IP 등 개인정보가 들어 있지 않습니다 (번역본 부록 D에 제외 항목 기록).
- 이번 주(2026-W47)에는 SM의 PPT 저장본이 없습니다 (SM은 번역본 + HTML만 작성).
