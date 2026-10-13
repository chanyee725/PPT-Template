# PPT-Template

코드로 만드는 PPT 템플릿. `pptxgenjs` 로 덱을 만들고, 저장한 뒤 테마 색과 한글 테마
글꼴을 파일에 직접 써 넣는다. 모든 색이 테마 색이라 PowerPoint 의
`디자인 → 변형 → 색` 으로 덱 전체(띠·박스·차트)가 한 번에 바뀐다.

지금 있는 종류는 **중간 보고** 하나다 — 흰 바탕, 하단 네이비 띠, 탭 라벨 박스,
번호 알약, 검은 Summary 박스, Figure/Table 캡션.

## 처음 한 번

```
npm install
npm run setup        # 발표자 정보 — 회사 · 부서 · 이름 · 직책 · 이메일 · 연락처
```

`setup` 은 `presenter.json` 에 저장하고, 표지와 마지막 장이 그걸 읽는다. 이 파일은
`.gitignore` 에 있어 올라가지 않는다. 다시 돌리면 고칠 수 있다(Enter 는 그대로,
`-` 는 비우기). 비운 칸은 슬라이드에 나오지 않고, 다 비어 있으면 `회사 · 부서` /
`이름 · 직책` 자리 표시가 나온다. 형식은 `presenter.example.json` 참고.

**글꼴은 Pretendard 다.** 발표하는 PC 에 설치돼 있어야 한다
([Pretendard](https://github.com/orioncactus/pretendard)). 없으면 PowerPoint 가
다른 글꼴로 바꿔 그리므로 줄바꿈이 달라진다. 설치가 어려우면 PowerPoint 에서
`파일 → 옵션 → 저장 → 파일의 글꼴 포함` 으로 저장한다.

## 만들기

```
npm run interim      # out/interim-report.pptx — 중간 보고 견본
```

견본(`examples/interim-report.js`)은 모든 슬라이드·컴포넌트를 한 번씩 쓴다. 새 보고는
이 파일을 복사해 내용만 바꾸면 된다.

```js
const { createDeck } = require("./template");

const d = createDeck({ title: "프로젝트 중간 보고" });   // title 이 상단 머리글
d.cover({ title: "프로젝트 이름\n중간 보고", subtitle: "한 줄 설명", date: "2026. 10. 13." });
d.contents([{ title: "개요", items: ["요약"] }, { title: "진행 현황", items: ["경과"] }]);
d.divider("01", "개요", ["요약"]);                        // 섹션 시작
const s = d.content("요약", { lead: "핵심 메시지 한두 문장" });
d.summary(s, "sum", d.MX, 4.8, d.CW, 1.4, "Summary", ["완료", "막힘", "다음"]);
d.closing({ subtitle: "프로젝트 중간 보고" });
await d.save("out/deck.pptx");
```

발표자를 덱마다 바꾸려면 `createDeck({ presenter: { name, position, ... } })` 로
`presenter.json` 위에 덮어쓴다. 한 장만 바꾸려면 `cover({ org, name })`.

## 슬라이드 종류

| 함수 | 레이아웃 | 용도 |
|---|---|---|
| `cover` | COVER | 주제 · 부제 · 날짜 · 회사/부서 · 이름/직책 |
| `contents` | CONTENTS | 목차, 최대 3열 |
| `divider` | DIVIDER | 섹션 구분 — 큰 번호, 이후 슬라이드가 이 섹션에 묶인다 |
| `content` | CONTENT | 본문 — 제목 + `lead` 한 문단, 나머지는 컴포넌트로 |
| `closing` | CLOSING | 감사합니다 + 발표자, 이메일·연락처가 있으면 그 아래 |

## 컴포넌트

위치·크기는 인치, 페이지는 13.33 × 7.5. 본문 영역은 `MX`(0.6) 부터 폭 `CW`.
`lead` 아래 y ≈ 2.0 부터 하단 띠 위 y ≈ 6.7 까지 쓴다.

| 함수 | 모양 |
|---|---|
| `tabBox(s, name, x, y, w, h, label, { plain, dark })` | 위 가장자리에 탭 라벨이 붙은 박스. 기본은 연한 채움 |
| `stat(s, name, x, y, w, label, value, sub)` | 탭 박스 안의 큰 숫자 (높이 1.6 고정) |
| `pill(s, name, n, x, y, w, text, tag)` | 네이비 번호 원 + 연한 알약, `tag` 는 회색 꼬리표 |
| `summary(s, name, x, y, w, h, label, items)` | 검은 라벨 블록 + 회색 박스, 2개 이상이면 ①②③ |
| `barLabel(s, name, label, x, y, w)` | `| 소제목` |
| `table(s, name, header, rows, x, y, w, colW, { rowH })` | 네이비 머리 줄 표. 셀은 문자열이나 `high()` / `strong()` |
| `barChart(s, name, series, x, y, w, h, extra)` | 묶은 세로 막대, 시리즈 2개면 회색/네이비 |
| `caption(s, name, text, x, y, w)` | `Figure 1. ...` / `Table 1. ...` |
| `downArrow` / `rightArrow` | 흐름 화살표 |
| `text` / `bullets` | 기본 글상자 |

`name` 은 PowerPoint 선택 창에 보이는 개체 이름이다. 한 슬라이드 안에서 겹치지 않게.

## 색과 글꼴 바꾸기

```js
createDeck({
  title: "...",
  theme: {
    headFontFace: "Pretendard",              // 기본값
    bodyFontFace: "Pretendard",
    colors: { dk2: "6B2737", accent1: "6B2737", lt2: "F6EEF0" },  // 바꿀 칸만
  },
});
```

| 칸 | 쓰이는 곳 |
|---|---|
| `dk2` · `accent1` | 메인 색 — 띠, 번호, 제목 강조, 차트 두 번째 시리즈 |
| `lt2` | 박스·알약 채움 |
| `accent2` | 테두리·구분선, 차트 첫 시리즈 |
| `accent3` | 보조 글자 |
| `accent4` | Summary 블록 |
| `accent5` | `high()` 강조 |
| `accent6` | 표 가는 선 |

## 확인

LibreOffice 가 있으면 PDF 로 바꿔 보는 게 제일 정확하다. 없을 때 쓰는 대략 검사:

```
python tools/fit.py out/interim-report.pptx       # 글상자보다 글이 긴 곳
python tools/preview.py out/interim-report.pptx   # out/preview/sheet-N.png — 겹침·빈 곳
```

둘 다 `python-pptx` 와 `Pillow` 가 필요하다. 글자 폭은 `fc-match` 가 찾은
Pretendard 로 재므로, 이 머신에 Pretendard 가 없으면 대체 글꼴 기준의 추정이다
(`fit.py` 두 번째 인자로 글꼴 파일을 직접 줄 수 있다). 미리보기는 정렬과 세로
가운데를 무시하므로 최종 모습이 아니다.

## 파일

```
template.js            createDeck — 레이아웃, 슬라이드, 컴포넌트, 발표자 정보 읽기
apply-theme.js         저장 뒤 테마 색 · 한글 테마 글꼴 · +mj-ea 참조를 써 넣는다
scripts/setup.js       발표자 정보를 묻고 presenter.json 에 쓴다
presenter.example.json presenter.json 형식
examples/              종류별 견본 — 지금은 interim-report.js (중간 보고)
tools/                 fit.py · preview.py
```

`apply-theme.js` 가 하는 일: pptxgenjs 는 테마 색을 못 쓰고, 테마의 한글(`Hang`)
글꼴을 Office 기본값(맑은 고딕)으로 남겨서 한글이 지정한 글꼴로 나오지 않는다. 또
`+mj-lt` 로 지정한 글자의 동아시아 글꼴 칸에도 `+mj-lt` 를 쓰는데, 그 칸의 올바른
값은 `+mj-ea` 다. 저장한 파일을 열어 세 가지를 고친다.

중간 보고 디자인은 PPTPLUG 「차분한 스타일 논문 발표 PPT 템플릿」 미리보기를 참고해
코드로 새로 그렸다. 그 템플릿 파일이나 이미지는 들어 있지 않다.
