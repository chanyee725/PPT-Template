// Interim report, page by page as the weekly meeting deck runs -- placeholder
// content throughout. Copy it into outputs/<meeting date>/ for a real report,
// drop usePresenterFile, and fill in the text and photos.
//
//   cover · overview · components · flow chart · detail slides per phase
//   (one photo, a screenshot, three photos as steps, a calendar, text only)
//   · sequence diagram · before/after pair · key numbers · video · plan · end
//
// node examples/interim-report.js templates/interim-report.pptx
// PALETTE=softBlue node examples/interim-report.js preview/soft.pptx
const path = require("path");
const fs = require("fs");
const { createDeck } = require("../template");

const out = process.argv[2] || path.join(__dirname, "..", "templates", "interim-report.pptx");
fs.mkdirSync(path.dirname(out), { recursive: true });

// templates/ is committed, so this one keeps the presenter placeholders.
const d = createDeck({ title: "프로젝트 중간 보고", usePresenterFile: false, palette: process.env.PALETTE || "navy" });
const { W, MX, CW } = d;

d.cover({
  title: "프로젝트 이름\n중간 보고",
  subtitle: "2026년 10월 2주차 중간보고",
  date: "2026. 10. 19.",
});

// ---------- overview: the setup in one picture, and what must come out of it ----------
{
  const s = d.content("프로젝트 개요", { section: "프로젝트 개요", lead: "목표 — 한 문장으로" });
  const pw = 6.8;
  d.photo(s, "layout", MX, 2.25, pw, 3.76, { caption: "전체 배치" });
  const rx = MX + pw + 0.45;
  const rw = W - MX - rx;
  d.barLabel(s, "data-h", "얻어야 하는 데이터", rx, 2.3, rw);
  d.table(s, "data", ["데이터", "한 줄 설명"], [
    [d.strong("데이터 1"), "무엇을 어떻게 재는지 한 줄"],
    [d.strong("데이터 2"), "무엇을 어떻게 재는지 한 줄"],
    [d.strong("데이터 3"), "무엇을 어떻게 재는지 한 줄"],
    [d.strong("데이터 4"), "무엇을 어떻게 재는지 한 줄"],
  ], rx, 2.78, rw, [1.7, rw - 1.7], { rowH: 0.58, fontSize: 11, align: ["left", "left"] });
}

// ---------- components: one large picture, two small ones with notes ----------
{
  const s = d.content("구성 요소", { section: "구성 요소", lead: "구성을 한 문장으로" });
  d.photo(s, "main", MX, 2.0, 7.0, 4.3, { caption: "전체 구성" });
  const rx = MX + 7.45;
  const pw = 1.9;
  const tx = rx + pw + 0.25;
  const tw = W - MX - tx;
  [["part-a", 2.0, 2.05, "구성 요소 A"], ["part-b", 4.55, 1.75, "구성 요소 B"]].forEach(([name, y, h, label]) => {
    d.photo(s, name, rx, y, pw, h, { caption: label });
    d.barLabel(s, `${name}-h`, label, tx, y, tw);
    d.bullets(s, `${name}-list`, ["사양", "사양"], { x: tx + 0.12, y: y + 0.45, w: tw - 0.12, h: h - 0.45, fontSize: 12, paraSpaceAfter: 4 });
  });
}

// ---------- the flow chart; the detail slides below follow its phases ----------
d.phaseOverview("Project FlowChart", [
  { title: "1단계", desc: "단계 설명을\n두세 줄로" },
  { title: "2단계", desc: "단계 설명을\n두세 줄로" },
  { title: "3단계", desc: "단계 설명을\n두세 줄로" },
  { title: "4단계", desc: "단계 설명을\n두세 줄로" },
  { title: "5단계", desc: "단계 설명을\n두세 줄로" },
], { current: 3 });

// lead: one line ending in a noun -- "~ 확인", "~ 완료". Below it 진행 내용, then
// the same items in English for the readers who do not read Korean.
const progress = (...pairs) => [
  { title: "진행 내용", items: pairs.map(([ko]) => ko) },
  { title: "Progress", items: pairs.map(([, en]) => en) },
];

d.phaseDetail(1, {
  lead: "(한 줄 요약) ~ 확인",
  photos: [{ caption: "사진 설명" }],
  sections: progress(["내용 ~ 완료", "(item) done"], ["내용 ~ 확인", "(item) confirmed"]),
});

// a second slide on the same phase, e.g. a screenshot of a test
d.phaseDetail(1, {
  title: "1단계 — 테스트",
  lead: "(한 줄 요약) ~ 완료",
  photos: [{ caption: "화면 캡처" }],
  sections: progress(["내용 ~ 완료", "(item) done"], ["내용 ~ 확인", "(item) confirmed"]),
});

// three pictures read as steps, photos only
d.phaseDetail(2, {
  flow: true,
  lead: "(한 줄 요약) ~ 확인",
  photos: [{ caption: "입력" }, { caption: "중간 결과" }, { caption: "최종 결과" }],
});

// nothing to photograph yet: the month, with the week something arrives
{
  const s = d.phaseDetail(3, {
    lead: "(장비) 10월 3주차 입고 예정",
    photos: [],
    sections: progress(["(장비) 10월 3주차 입고 예정", "(equipment) arrives in week 3 of October"]),
  });
  d.calendar(s, "month", 2026, 10, MX, 2.0, 7.1, 4.3, { marks: [{ from: 19, to: 23, label: "입고 예정" }] });
}

// not started: text only
d.phaseDetail(4, {
  lead: "(협의 필요 사항) — 아직 미착수",
  textOnly: true,
  sections: progress(["(누구) 와 (무엇) 협의 필요", "(what) to be agreed with (whom)"], ["협의 전까지 미착수", "Not started until then"]),
});

// ---------- the whole cycle as a sequence diagram ----------
{
  const s = d.content("전체 시퀀스", { section: "전체 시퀀스", lead: "한 사이클 — 순서대로 진행하고 반복" });
  d.sequenceDiagram(s, "seq", [
    {
      title: "Pick",
      lanes: ["로봇", "카메라"],
      events: [
        [0, 0, "① 준비 위치로 이동", "act"],
        [0, 1, "② 촬영 요청", "call"],
        [1, 1, "③ 대상 찾기", "act"],
        [1, 0, "위치", "reply"],
        [0, 0, "④ 집기", "act"],
      ],
    },
    {
      title: "Place",
      lanes: ["로봇", "카메라", "설비"],
      events: [
        [0, 2, "⑤ 이동 요청", "call"],
        [2, 2, "⑥ 이동", "act"],
        [2, 0, "⑦ 도착 신호", "reply"],
        [0, 1, "⑧ 놓을 곳 찾기", "call"],
        [1, 0, "놓을 위치", "reply"],
        [0, 0, "⑨ 놓기", "act"],
      ],
      while: { rows: [3, 4], label: "while [미검출 · ≤3회]" },
    },
  ]);
}

// ---------- a before/after pair joined by an arrow ----------
{
  const s = d.content("가상환경 구성", { section: "가상환경 구성", lead: "실제 배치를 가상환경으로 구현" });
  const gap = 0.8;
  const boxW = (CW - gap) / 2;
  const boxH = 3.2;
  const top = 1.6 + (6.75 - 1.6 - boxH - 0.35) / 2;
  d.photo(s, "real", MX, top, boxW, boxH, { caption: "실제" });
  d.photo(s, "sim", MX + boxW + gap, top, boxW, boxH, { caption: "가상환경" });
  d.flowArrow(s, "flow", MX + boxW + gap / 2, top + boxH / 2);
}

// ---------- key numbers next to the picture they belong to ----------
{
  const s = d.content("허용 오차", { section: "허용 오차", lead: "(무엇) 의 허용 오차를 x 축 (값) · y 축 (값) 으로 설정" });
  d.photo(s, "subject", MX, 1.75, 4.4, 4.0, { caption: "대상" });
  const rx = MX + 4.85;
  const rw = W - MX - rx;
  d.photo(s, "frame", rx, 1.75, rw, 2.3, { caption: "좌표 — x: (방향), y: (방향)" });
  d.valueTiles(s, "tol", [
    { badge: "x", label: "(방향)", value: 50, unit: " mm" },
    { badge: "y", label: "(방향)", value: 30, unit: " mm" },
  ], rx, 4.6, rw);
}

// ---------- the run on video ----------
// d.video(s, "run", "outputs/<date>/photos/run.mp4", x, y, w, h, { cover: "....png" })
// embeds an mp4; the box below stands in for it here.
{
  const s = d.content("실행 영상", { section: "실행 영상" });
  const vh = 5.6;
  const vw = vh * 16 / 9;
  d.photo(s, "run", MX + (CW - vw) / 2, 1.1, vw, vh, {});
}

// ---------- what comes next ----------
{
  const s = d.content("향후 일정", { section: "향후 일정", lead: "다음 단계를 한 문장으로" });
  d.phaseFlow(s, "plan", [
    { label: "완료", title: "할 일", desc: "설명" },
    { label: "10월 3주차", title: "할 일", desc: "설명" },
    { label: "이후", title: "할 일", desc: "설명" },
    { label: "10월 4주차", title: "할 일", desc: "설명" },
    { label: "예정", title: "할 일", desc: "설명" },
  ], { current: 2 });
}

d.ending();

d.save(out).then((f) => console.log("wrote", f));
