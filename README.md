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
   python3 scripts/archive.py add ~/chem99_0127      # 추가 설치 없음 (파이썬 3 표준 라이브러리만 사용)
   ```
   파일이 `archive/YYYY-Wnn/`으로 복사되고, `index.html`·`archive/index.json`이 갱신됩니다. (같은 이름 파일은 덮어씁니다.)
   허브에 나오는 가격·등락률은 번역본 0절 "가격 데이터" 표의 앞 두 행에서 자동으로 읽습니다.
3. 커밋·푸시합니다.
   ```bash
   git add -A && git commit -m "2027-W04 주보 추가" && git push
   ```

허브만 다시 만들고 싶으면 `python3 scripts/archive.py build`.

## 허브 사용법

- 맨 위 **대표 보고서 뷰어**에 SM 최신 주차 보고서가 바로 열립니다. 위쪽 칩(SM·AN·BD·BPA·MMA)과 주차 선택으로 다른 보고서로 바꿉니다.
- **▶ 바로 발표**: 버튼 한 번으로 전체화면 발표 모드 (→/Space 다음, ← 이전, V 전체보기, B/W 가림, L 레이저, Esc 종료).
- **▣ 전체보기**: 모든 쪽을 한 화면에서 훑어보기. **↗ 새 창으로 열기**: 보고서만 새 탭에서 열기. PPT·번역본 링크도 같은 줄에 있습니다.
- 아래 **주차별 보관함**에서 주차·품목별로 찾고 [화면에서 보기]로 뷰어에 불러옵니다. 주소에 `?p=SM&w=2026-W47` 을 붙이면 그 보고서가 열린 채로 시작합니다.
- 보고서 HTML 파일은 수정하지 않고 허브가 보고서 안의 발표·전체보기 기능을 호출하는 방식이라, 같은 사이트(GitHub Pages)에서 열었을 때만 연동됩니다. 파일을 직접 더블클릭해 연 경우에는 새 창으로 대신 열립니다.
- 전체화면은 브라우저 정책상 클릭이 있어야 켜집니다. 아이폰 Safari는 전체화면을 지원하지 않아 창 안 발표로 표시됩니다.

## 보기

- GitHub Pages를 켜면(Settings → Pages → Branch: `main` / root) `https://<계정>.github.io/chem99-report/` 에서 허브와 HTML 보고서를 바로 열 수 있습니다.
- Pages 없이 쓰려면 저장소를 내려받아 `index.html`을 브라우저로 열면 됩니다 (CDN·외부 라이브러리 없음).

## 참고

- 보관 파일에는 계정·IP 등 개인정보가 들어 있지 않습니다 (번역본 부록 D에 제외 항목 기록).
- 이번 주(2026-W47)에는 SM의 PPT 저장본이 없습니다 (SM은 번역본 + HTML만 작성).
