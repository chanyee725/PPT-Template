// Interim report: cover, the phase timeline, then detail slides phase by phase
// (photos + notes), issues and next steps, closing -- placeholder content.
// node examples/interim-report.js templates/interim-report.pptx
// PALETTE=slate node examples/interim-report.js preview/slate.pptx
const path = require("path");
const fs = require("fs");
const { createDeck } = require("../template");

const out = process.argv[2] || path.join(__dirname, "..", "templates", "interim-report.pptx");
fs.mkdirSync(path.dirname(out), { recursive: true });

// templates/ is committed, so this one keeps the presenter placeholders.
// A copy for a real report drops usePresenterFile to pick up presenter.json.
// PALETTE=softBlue (or slate, teal) builds the same deck in another palette.
const d = createDeck({ title: "프로젝트 중간 보고", usePresenterFile: false, palette: process.env.PALETTE || "navy" });
const { W, MX, CW } = d;

d.cover({
  title: "프로젝트 이름\n중간 보고",
  subtitle: "2026년 10월 2주차 중간보고",
  date: "2026. 10. 13.",
});

// ---------- the whole project, then this stretch of it ----------
// Each phaseOverview remembers its phases for the phaseDetail calls after it,
// so the one the detail slides follow comes last.
d.phaseOverview("Project Roadmap", [
  { title: "요구사항 정의", desc: "작업 대상과 처리량,\n현장 조건을 정리" },
  { title: "시스템 설계", desc: "로봇 · 카메라 · 대차\n구성과 인터페이스 결정" },
  { title: "통합 개발", desc: "장비 셋업과 연동,\n픽업 시퀀스 구현" },
  { title: "현장 시운전", desc: "실제 라인에서\n반복 운전과 보완" },
  { title: "운영 이관", desc: "운영 매뉴얼 작성과\n담당자 교육" },
], { current: 3 });

// ---------- the timeline ----------
d.phaseOverview("Project FlowChart", [
  { title: "로봇 셋업", desc: "컨트롤러 연결과 관절 리밋 확인,\n티칭 포인트 등록" },
  { title: "카메라 셋업", desc: "카메라 장착과 hand-eye\n캘리브레이션, 대상 인식 확인" },
  { title: "대차 PLC 통합", desc: "Modbus 로 대차 도착 신호를\n받아 시퀀스와 연동" },
  { title: "전체 시퀀스 테스트", desc: "픽업부터 적재까지\n전체 흐름을 반복 시험" },
], { current: 3 });

// ---------- one or more slides per phase ----------
// lead: one line ending in a noun -- "~ 확인", "~ 완료". Below it 진행 내용, then
// the same items in English for the readers who do not read Korean.
// photos: { path: "photos/robot.jpg", caption } -- leave path out for a placeholder.
// Up to 8 phases fit the flow chart.
const progress = (...pairs) => [
  { title: "진행 내용", items: pairs.map(([ko]) => ko) },
  { title: "Progress", items: pairs.map(([, en]) => en) },
];

d.phaseDetail(1, {
  lead: "(한 줄 요약) ~ 확인",
  photos: [{ caption: "사진 설명" }],
  sections: progress(["내용 ~ 완료", "(item) done"], ["내용 ~ 확인", "(item) confirmed"]),
});

d.phaseDetail(2, {
  lead: "(한 줄 요약) ~ 완료 — 사진 두 장을 나란히 둘 수도 있다",
  photos: [{ caption: "사진 설명" }, { caption: "사진 설명" }],
  sections: progress(["내용 ~ 완료", "(item) done"], ["내용 ~ 확인", "(item) confirmed"]),
});

d.phaseDetail(3, {
  lead: "(한 줄 요약) ~ 진행 중",
  photos: [{ caption: "사진 설명" }],
  sections: progress(["내용 ~ 완료", "(item) done"], ["내용 ~ 진행 중", "(item) in progress"]),
});

d.phaseDetail(4, {
  lead: "(한 줄 요약) ~ 예정",
  photos: [{ caption: "구성도 또는 계획" }],
  sections: progress(["내용 ~ 예정", "(item) planned"]),
});

// ---------- issues and next steps ----------
{
  const s = d.content("리스크", { section: "이슈 및 향후 계획", lead: "등급 · 리스크 · 영향 · 대응을 한 표에." });
  d.caption(s, "tab", "Table 1. 주요 리스크와 대응", MX, 2.05, 8);
  d.table(s, "risks", ["등급", "리스크", "영향", "대응"], [
    [d.high("높음"), "리스크", "영향", "대응"],
    [d.strong("중간"), "리스크", "영향", "대응"],
  ], MX, 2.45, CW, [1.1, 2.7, 4.9, CW - 8.7], { rowH: 0.75 });
}

{
  const s = d.content("다음 단계", { lead: "순서대로 할 일과, 듣는 사람에게 요청할 것." });
  ["할 일", "할 일", "할 일", "할 일"].forEach((t, i) => d.pill(s, `next-${i}`, i + 1, MX, 2.2 + i * 0.82, 7.4, t));
  const ax = MX + 7.9;
  d.tabBox(s, "ask", ax, 2.35, W - MX - ax, 3.9, "요청 사항");
  d.bullets(s, "ask-list", ["요청", "요청"], { x: ax + 0.3, y: 2.75, w: W - MX - ax - 0.6, h: 3.2, paraSpaceAfter: 12 });
}

d.closing({ subtitle: "프로젝트 중간 보고" });

d.save(out).then((f) => console.log("wrote", f));
