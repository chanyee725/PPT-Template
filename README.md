# PPT-Template

코드로 만드는 PPT 템플릿. `pptxgenjs` 로 덱을 만들고, 저장한 뒤 테마 색과 한글 테마
글꼴을 파일에 직접 써 넣는다. 모든 색이 테마 색이라 PowerPoint 의
`디자인 → 변형 → 색` 으로 덱 전체(띠·박스·차트)가 한 번에 바뀐다.

지금 있는 종류는 **중간 보고** 하나다 — 표지 → Project FlowChart(단계 타임라인) → 단계별 상세
(사진 + 설명, 단계마다 한 장 이상) → 리스크 · 다음 단계 → 마무리. 흰 바탕에 하단
네이비 띠, Pretendard.

## 처음 한 번

```
npm install
npm run setup        # 발표자 정보 — 회사 · 부서 · 이름 · 직책 · 이메일 · 연락처
```

`setup` 은 `presenter.json` 에 저장하고, 표지와 마지막 장이 그걸 읽는다. 이 파일은
`.gitignore` 에 있어 올라가지 않는다. 다시 돌리면 고칠 수 있다(Enter 는 그대로,
`-` 는 비우기). 비운 칸은 슬라이드에 나오지 않고, 다 비어 있으면 `회사 · 부서` /
`직책 이름` 자리 표시가 나온다. 형식은 `presenter.example.json` 참고.

**글꼴은 Pretendard 다.** 발표하는 PC 에 설치돼 있어야 한다
([Pretendard](https://github.com/orioncactus/pretendard)). 없으면 PowerPoint 가
다른 글꼴로 바꿔 그리므로 줄바꿈이 달라진다. 설치가 어려우면 PowerPoint 에서
`파일 → 옵션 → 저장 → 파일의 글꼴 포함` 으로 저장한다.

## 만들기

```
npm run interim      # templates/interim-report.pptx — 중간 보고
```

`templates/` 의 pptx 는 저장소에 올라가는 템플릿이라 발표자 칸이 자리 표시 그대로다
(`examples/interim-report.js` 가 `usePresenterFile: false` 로 만든다). 실제 보고는
이 파일을 복사해 내용을 바꾸고 `usePresenterFile: false` 를 지우면 `presenter.json`
의 발표자 정보가 들어간다. 그런 덱은 `outputs/<회의 날짜>/` 에 둔다(매주 월요일 회의,
예: `outputs/2026-10-19/`) — git 이 무시하는 폴더라 이름·연락처·사진이 올라가지 않는다.

```js
const { createDeck } = require("./template");

const d = createDeck({ title: "프로젝트 중간 보고" });   // title 이 상단 머리글
d.cover({ title: "프로젝트 이름\n중간 보고", subtitle: "한 줄 설명", date: "2026. 10. 13." });
d.divider("01", "개요", ["요약"]);                        // 섹션 시작
const s = d.content("요약", { lead: "핵심 메시지 한두 문장" });
d.summary(s, "sum", d.MX, 4.8, d.CW, 1.4, "Summary", ["완료", "막힘", "다음"]);
d.closing({ subtitle: "프로젝트 중간 보고" });
await d.save("templates/deck.pptx");
```

단계별 상세 슬라이드:

```js
d.phaseOverview("Project FlowChart", [
  { title: "로봇 셋업", desc: "..." }, { title: "카메라 셋업", desc: "..." },
  { title: "대차 PLC 통합", desc: "..." }, { title: "전체 시퀀스 테스트", desc: "..." },
], { current: 3 });
d.phaseDetail(1, {
  lead: "이 단계에서 한 일 한 문장",
  photos: [{ path: "outputs/2026-10-19/photos/robot.jpg", caption: "로봇 설치" }],
  sections: [{ title: "진행 내용", items: ["..."] }, { title: "결과", items: ["..."] }],
});
d.phaseDetail(1, { title: "로봇 셋업 — 툴 오프셋", ... });   // 같은 단계 두 번째 장
```

발표자를 덱마다 바꾸려면 `createDeck({ presenter: { name, position, ... } })` 로
`presenter.json` 위에 덮어쓴다. 한 장만 바꾸려면 `cover({ org, name })`.

## 슬라이드 종류

| 함수 | 레이아웃 | 용도 |
|---|---|---|
| `cover` | COVER | 주제 · 부제 한 줄 · 왼쪽 아래 회사/부서와 「직책 이름」 · 오른쪽 아래 날짜 |
| `contents` | CONTENTS | 목차, 최대 3열 — 중간 보고에는 쓰지 않는다 |
| `divider` | DIVIDER | 섹션 구분 — 큰 번호, 이후 슬라이드가 이 섹션에 묶인다 |
| `content` | CONTENT | 본문 — 제목 + `lead` 한 문단, 나머지는 컴포넌트로. `section` 을 주면 구분 슬라이드 없이 새 섹션 |
| `closing` | CLOSING | 감사합니다 + 발표자, 이메일·연락처가 있으면 그 아래 |
| `phaseOverview(title, phases, { current, remember })` | FEATURE | 흐름도 한 장 — `title` 은 오른쪽 위 머리글 자리, 단계 타임라인은 가운데. 단계 목록을 기억해 `phaseDetail` 이 쓴다(`remember: false` 면 기억하지 않음 — 번역본 같은 사본용) |
| `phaseDetail(n, { title, lead, photos, sections })` | CONTENT | n 번째 단계 상세 — 제목과 `lead` 아래 왼쪽 사진 1~2장, 오른쪽 `\| 소제목` + 글머리 묶음. 같은 단계로 여러 번 부르면 여러 장 |

## 컴포넌트

위치·크기는 인치, 페이지는 13.33 × 7.5. 본문 영역은 `MX`(0.6) 부터 폭 `CW`.
`lead` 아래 y ≈ 2.0 부터 하단 띠 위 y ≈ 6.7 까지 쓴다.

| 함수 | 모양 |
|---|---|
| `phaseFlow(s, name, phases, { current, x, y, w })` | 가로 단계 타임라인 — Phase 라벨 · 후광 점 · 제목 · 설명. `current`(1부터)는 후광이 크고, 뒤 단계는 회색, 선은 현재까지 네이비. 최대 8단계 — 5단계부터 글자가 작아지고, 긴 제목은 두 줄, 설명은 모든 열이 같은 높이에서 시작 |
| `photo(s, name, x, y, w, h, { path, caption })` | 상자 안에 비율 유지로 맞춘 사진(PNG · JPEG). `path` 가 없으면 X 표시 자리 상자. 캡션은 `Figure n.` 자동 번호 |
| `tabBox(s, name, x, y, w, h, label, { plain, dark })` | 위 가장자리에 탭 라벨이 붙은 박스. 기본은 연한 채움 |
| `stat(s, name, x, y, w, label, value, sub)` | 탭 박스 안의 큰 숫자 (높이 1.6 고정) |
| `pill(s, name, n, x, y, w, text, tag)` | 네이비 번호 원 + 연한 알약, `tag` 는 회색 꼬리표 |
| `summary(s, name, x, y, w, h, label, items)` | 검은 라벨 블록 + 회색 박스, 2개 이상이면 ①②③ |
| `barLabel(s, name, label, x, y, w)` | `| 소제목` |
| `table(s, name, header, rows, x, y, w, colW, { rowH, headH, fontSize, align })` | 위쪽 네이비 선 · 연한 머리 줄(본문 줄보다 낮게) · 가는 줄 구분. `align` 은 열마다 정렬(머리 줄도 같이), 기본은 첫 열만 가운데. 셀은 문자열이나 `high()` / `strong()` |
| `barChart(s, name, series, x, y, w, h, extra)` | 묶은 세로 막대, 시리즈 2개면 회색/네이비 |
| `caption(s, name, text, x, y, w)` | `Figure 1. ...` / `Table 1. ...` |
| `downArrow` / `rightArrow` | 흐름 화살표 |
| `text` / `bullets` | 기본 글상자 |

`name` 은 PowerPoint 선택 창에 보이는 개체 이름이다. 한 슬라이드 안에서 겹치지 않게.

## 색과 글꼴 바꾸기

색 묶음(palette)이 넷 있고 기본은 `navy` 다. `createDeck({ palette: "softBlue" })` 나
`PALETTE=softBlue npm run interim` 으로 바꾼다.

| palette | 메인 색 | 느낌 |
|---|---|---|
| `navy` (기본) | `#012D6A` | 진한 네이비 |
| `softBlue` | `#4A6BD1` | 밝은 파랑 |
| `slate` | `#4F6F96` | 회색 섞인 파랑 |
| `teal` | `#2F8291` | 청록 |

한 칸만 바꾸려면 `theme.colors` 로 덮어쓴다.

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
| `dk2` | 메인 색 — 띠, 번호, 제목 강조, 진행 선, 차트 두 번째 시리즈 |
| `accent1` | 메인 색보다 한 단계 밝은 색 — 흐름도에서 진행 중인 단계 동그라미에만 |
| `lt2` | 박스·알약 채움 |
| `accent2` | 테두리·구분선, 차트 첫 시리즈 |
| `accent3` | 보조 글자 |
| `accent4` | Summary 블록 |
| `accent5` | `high()` 강조 |
| `accent6` | 표 가는 선 |

## 확인

LibreOffice 가 있으면 PDF 로 바꿔 보는 게 제일 정확하다. 없을 때 쓰는 대략 검사:

```
python tools/fit.py templates/interim-report.pptx       # 글상자보다 글이 긴 곳
python tools/preview.py templates/interim-report.pptx   # preview/sheet-N.png — 겹침·빈 곳
```

둘 다 `python-pptx` 와 `Pillow` 가 필요하다. 글자 폭은 Pretendard 로 재고, 이
머신에 없으면 나눔고딕으로 대신하므로 그때는 추정이다(`fit.py` 두 번째 인자로 글꼴
파일을 직접 줄 수 있다). 미리보기는 세로 가운데 정렬을 무시하므로 최종 모습이 아니다.

## 파일

```
template.js            createDeck — 레이아웃, 슬라이드, 컴포넌트, 발표자 정보 읽기
apply-theme.js         저장 뒤 테마 색 · 한글 테마 글꼴 · +mj-ea 참조를 써 넣는다
scripts/setup.js       발표자 정보를 묻고 presenter.json 에 쓴다
presenter.example.json presenter.json 형식
examples/              종류별 생성 코드 — 지금은 interim-report.js (중간 보고)
templates/             생성된 템플릿 pptx (발표자 칸은 자리 표시)
outputs/<날짜>/         회의별 실제 보고 덱 · 사진 (git 무시)
tools/                 fit.py · preview.py
```

`apply-theme.js` 가 하는 일: pptxgenjs 는 테마 색을 못 쓰고, 테마의 한글(`Hang`)
글꼴을 Office 기본값(맑은 고딕)으로 남겨서 한글이 지정한 글꼴로 나오지 않는다. 또
`+mj-lt` 로 지정한 글자의 동아시아 글꼴 칸에도 `+mj-lt` 를 쓰는데, 그 칸의 올바른
값은 `+mj-ea` 다. 저장한 파일을 열어 세 가지를 고친다.

중간 보고 디자인은 PPTPLUG 「차분한 스타일 논문 발표 PPT 템플릿」 미리보기를 참고해
코드로 새로 그렸다. 그 템플릿 파일이나 이미지는 들어 있지 않다.
