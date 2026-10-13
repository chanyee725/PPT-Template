// Interim report: every slide type and component once, with placeholder content.
// node examples/interim-report.js out/interim-report.pptx
const path = require("path");
const fs = require("fs");
const { createDeck } = require("../template");

const out = process.argv[2] || path.join(__dirname, "..", "out", "interim-report.pptx");
fs.mkdirSync(path.dirname(out), { recursive: true });

const d = createDeck({ title: "프로젝트 중간 보고" });
const { W, MX, CW } = d;

d.cover({
  title: "프로젝트 이름\n중간 보고",
  subtitle: "한 줄로 쓰는 프로젝트 설명",
  date: "2026. 10. 13.",
});

d.contents([
  { title: "프로젝트 개요", items: ["요약", "시스템 구성"] },
  { title: "진행 현황", items: ["추진 경과", "성능 측정", "단계별 현황"] },
  { title: "이슈 및 향후 계획", items: ["리스크", "다음 단계"] },
]);

// ---------- 01 ----------
d.divider("01", "프로젝트 개요", ["요약", "시스템 구성"]);

{
  const s = d.content("요약", { lead: "핵심 메시지를 한두 문장으로 쓴다. 보고를 듣는 사람이 이 문장만 읽어도 지금 상태를 알 수 있어야 한다." });
  const gap = 0.35;
  const bw = (CW - gap * 3) / 4;
  [["지표 A", "94", "설명 한 줄"], ["지표 B", "580", "설명 한 줄"], ["지표 C", "3 / 7", "설명 한 줄"], ["지표 D", "717ms", "설명 한 줄"]]
    .forEach(([label, value, sub], i) => d.stat(s, `stat-${i}`, MX + i * (bw + gap), 2.35, bw, label, value, sub));
  d.downArrow(s, "sum-arrow", W / 2, 4.2);
  d.summary(s, "sum", MX + 0.8, 4.85, CW - 1.6, 1.45, "Summary", [
    "완료 — 지금까지 끝난 것",
    "막힘 — 진행을 막고 있는 것",
    "다음 — 바로 이어서 할 것",
  ]);
}

{
  const s = d.content("시스템 구성", { lead: "구성 요소와 그 사이의 접점을 보여 준다." });
  const bw = 4.3;
  d.tabBox(s, "left", MX, 2.3, bw, 1.35, "구성 요소 A", { plain: true });
  d.text(s, "left-b", "하는 일을 한 줄로", { x: MX + 0.3, y: 2.6, w: bw - 0.6, h: 0.8, valign: "middle" });
  d.tabBox(s, "right", W - MX - bw, 2.3, bw, 1.35, "구성 요소 B", { plain: true });
  d.text(s, "right-b", "하는 일을 한 줄로", { x: W - MX - bw + 0.3, y: 2.6, w: bw - 0.6, h: 0.8, valign: "middle" });
  d.rightArrow(s, "ab", W / 2, 2.97);
  d.downArrow(s, "down", W / 2, 3.85);
  d.tabBox(s, "whole", MX, 4.5, CW, 1.95, "전체 시스템");
  const pw = (CW - 0.6 - 0.3 * 3) / 4;
  ["화면", "처리", "제어", "설정"].forEach((h, i) => {
    const x = MX + 0.3 + i * (pw + 0.3);
    d.barLabel(s, `part-${i}`, h, x, 4.85, pw);
    d.text(s, `part-${i}-b`, "세부 내용 한두 줄", { x: x + 0.12, y: 5.3, w: pw - 0.12, h: 0.9, fontSize: 13 });
  });
  d.caption(s, "fig1", "Figure 1. 시스템 구성", MX, 6.6, 6);
}

// ---------- 02 ----------
d.divider("02", "진행 현황", ["추진 경과", "성능 측정", "단계별 현황"]);

{
  const s = d.content("추진 경과", { lead: "날짜 순으로 무엇을 했는지 정리한다." });
  d.caption(s, "tab1", "Table 1. 주요 추진 경과", MX, 2.05, 8);
  d.table(s, "timeline", ["일자", "구분", "주요 내용"], [
    ["09-21", "착수", "내용"],
    ["09-28", "설계", "내용"],
    ["10-05", "구현", "내용"],
    ["10-12", "시험", "내용"],
  ], MX, 2.45, CW, [1.8, 2.2, CW - 4.0], { rowH: 0.62 });
}

{
  const s = d.content("성능 측정", { lead: "측정값은 표와 차트로, 조건은 캡션에 적는다." });
  d.tabBox(s, "cond", MX, 2.35, 5.6, 3.9, "측정 조건");
  ["항목 1", "항목 2", "항목 3", "항목 4"].forEach((h, i) => {
    const y = 2.75 + i * 0.85;
    d.barLabel(s, `cond-${i}`, h, MX + 0.3, y, 5.0);
    d.text(s, `cond-${i}-b`, "설명", { x: MX + 0.42, y: y + 0.36, w: 4.9, h: 0.3, fontSize: 12, color: d.C.accent3 });
  });
  d.barChart(s, "chart", [
    { name: "이전", labels: ["중앙값", "p95"], values: [6.4, 20.5] },
    { name: "이후", labels: ["중앙값", "p95"], values: [0.1, 0.2] },
  ], 6.7, 2.2, 6.0, 3.9);
  d.caption(s, "fig2", "Figure 2. 측정 결과 (단위)", 6.7, 6.2, 6.0);
}

{
  const s = d.content("단계별 현황", { lead: "단계를 번호로 나열하고, 막힌 단계는 꼬리표를 단다." });
  const steps = [["단계 1", null], ["단계 2", null], ["단계 3", "대기 사유"], ["단계 4", null], ["단계 5", "대기 사유"], ["단계 6", null]];
  const colW = (CW - 0.5) / 2;
  steps.forEach(([t, tag], i) => {
    const col = i < 3 ? 0 : 1;
    d.pill(s, `step-${i}`, i + 1, MX + col * (colW + 0.5), 2.1 + (i % 3) * 0.68, colW, t, tag);
  });
  d.downArrow(s, "arrow", W / 2, 4.35);
  d.summary(s, "now", MX + 0.8, 4.95, CW - 1.6, 1.15, "현재 가능", ["지금 실제로 돌릴 수 있는 범위"]);
}

{
  const s = d.content("확인된 것과 막힌 것", { lead: "왼쪽은 확인된 것, 오른쪽은 막힌 것." });
  const bw = 5.3;
  d.tabBox(s, "ok", MX, 2.35, bw, 2.75, "확인됨", { plain: true });
  d.bullets(s, "ok-list", ["항목", "항목", "항목"], { x: MX + 0.3, y: 2.6, w: bw - 0.6, h: 2.3, valign: "middle" });
  d.rightArrow(s, "arrow", W / 2, 3.72);
  d.tabBox(s, "blocked", W - MX - bw, 2.35, bw, 2.75, "막힘", { dark: true });
  d.bullets(s, "blocked-list", ["항목", "항목", "항목"], { x: W - MX - bw + 0.3, y: 2.6, w: bw - 0.6, h: 2.3, valign: "middle", color: d.C.background1 });
  d.summary(s, "note", MX, 5.45, CW, 1.15, "주의", ["한 줄짜리 주의 사항"]);
}

// ---------- 03 ----------
d.divider("03", "이슈 및 향후 계획", ["리스크", "다음 단계"]);

{
  const s = d.content("리스크", { lead: "등급 · 리스크 · 영향 · 대응을 한 표에." });
  d.caption(s, "tab", "Table 2. 주요 리스크와 대응", MX, 2.05, 8);
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
