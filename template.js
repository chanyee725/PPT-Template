// Navy report deck: white pages, Pretendard, a navy band, tabbed boxes,
// numbered pills and a dark "Summary" block. Every colour is a theme colour,
// so PowerPoint's Design > Variants > Colors recolours the whole deck.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply-theme");

const DEFAULT_THEME = {
  name: "Navy Report",
  headFontFace: "Pretendard",
  bodyFontFace: "Pretendard",
  colors: {
    dk1: "222222", // body text
    lt1: "FFFFFF", // page
    dk2: "012D6A", // navy: bands, numbers, headings
    lt2: "EEF2F8", // pale tint: box fills, pills
    accent1: "012D6A",
    accent2: "8A9BB8", // borders and rules
    accent3: "6E6E6E", // muted text
    accent4: "262626", // summary block
    accent5: "C0392B", // "high" emphasis
    accent6: "D9DEE7", // hairlines
    hlink: "012D6A",
    folHlink: "6E6E6E",
  },
};

const W = 13.333;
const H = 7.5;
const MX = 0.6;
const CW = W - 2 * MX;
const HEAD = "+mj-lt";
const MARKS = ["①", "②", "③", "④", "⑤", "⑥"];

// Who is presenting. `npm run setup` writes presenter.json, which stays out of
// git; a field left empty is simply not shown.
const PRESENTER_FILE = path.join(__dirname, "presenter.json");
const PRESENTER_FIELDS = ["company", "department", "name", "position", "email", "phone"];

function loadPresenter() {
  try {
    return JSON.parse(fs.readFileSync(PRESENTER_FILE, "utf8"));
  } catch (err) {
    if (err.code === "ENOENT") return {};
    throw new Error(`${PRESENTER_FILE}: ${err.message}`);
  }
}

// "회사 · 부서" and "이름 · 직책" lines, or placeholders where nothing is set.
function presenterLines(p) {
  const join = (...parts) => parts.filter(Boolean).join(" · ");
  return {
    org: join(p.company, p.department) || "회사 · 부서",
    who: join(p.name, p.position) || "이름 · 직책",
    contact: join(p.email, p.phone),
  };
}

function createDeck({ title, header, theme, presenter } = {}) {
  const THEME = Object.assign({}, DEFAULT_THEME, theme, {
    colors: Object.assign({}, DEFAULT_THEME.colors, theme && theme.colors),
  });
  const HEX = THEME.colors;
  const runningHeader = header || title || "";
  const who = presenterLines(Object.assign({}, loadPresenter(), presenter));

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = title || "";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const NAVY = C.text2;
  const TINT = C.background2;
  let section = null;

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: C.background1 },
    objects: [
      { line: { x: 0.9, y: 1.05, w: 7.0, h: 0, line: { color: NAVY, width: 1.25 } } },
      { rect: { x: 0, y: 6.95, w: W, h: 0.55, fill: { color: NAVY }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 1.2, w: 10.5, h: 1.8, fontFace: HEAD, fontSize: 40, bold: true, color: NAVY, valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 3.1, w: 10.5, h: 0.9, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENTS",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 0, y: 0, w: 0.45, h: H, fill: { color: NAVY }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 0.4, w: 5, h: 0.6, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: runningHeader, options: { x: W - 6.6, y: 0.5, w: 6.0, h: 0.4, fontFace: HEAD, fontSize: 11, color: NAVY, align: "right", margin: 0 } } },
    ],
  });

  pres.defineSlideMaster({
    title: "DIVIDER",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 0.9, y: 0, w: 2.2, h: H, fill: { color: NAVY }, line: { type: "none" } } },
      { line: { x: 3.1, y: 3.95, w: W - 3.1, h: 0, line: { color: C.accent2, width: 0.75 } } },
      { placeholder: { options: { name: "title", type: "title", x: 3.6, y: 3.1, w: 8.5, h: 0.8, fontFace: HEAD, fontSize: 32, bold: true, color: NAVY, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "num", type: "body", x: 0.9, y: 2.6, w: 2.2, h: 1.5, fontSize: 72, color: C.background1, align: "center", valign: "middle", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { text: { text: runningHeader, options: { x: W - 6.6, y: 0.38, w: 6.0, h: 0.35, fontFace: HEAD, fontSize: 11, color: NAVY, align: "right", margin: 0 } } },
      { rect: { x: 0, y: 7.28, w: W, h: 0.22, fill: { color: NAVY }, line: { type: "none" } } },
      { line: { x: W - 0.95, y: 7.12, w: 0.35, h: 0, line: { color: C.accent2, width: 0.5 } } },
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.3, w: 6.5, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: C.text1, valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: W - 0.95, y: 6.82, w: 0.35, h: 0.28, fontSize: 10, color: C.accent3, align: "center" },
  });

  pres.defineSlideMaster({
    title: "CLOSING",
    background: { color: NAVY },
    objects: [
      { line: { x: 0.9, y: 1.6, w: 0, h: 1.5, line: { color: C.background1, width: 1 } } },
      { placeholder: { options: { name: "title", type: "title", x: 1.2, y: 1.6, w: 9, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: C.background1, valign: "top", margin: 0 }, text: "" } },
    ],
  });

  // ---------- primitives ----------
  function text(slide, name, str, opts) {
    slide.addText(str, Object.assign({ objectName: name, isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, opts));
  }

  function bullets(slide, name, items, opts) {
    slide.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })),
      Object.assign({ objectName: name, isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", paraSpaceAfter: 8 }, opts));
  }

  function startSection(name) {
    section = name;
    pres.addSection({ title: name });
  }

  // ---------- slides ----------
  // Presenter lines come from presenter.json unless given here.
  function cover({ title, subtitle, date, org = who.org, name = who.who }) {
    startSection("표지");
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    s.addText([subtitle, date].filter(Boolean).join("\n"), { placeholder: "body" });
    text(s, "presenter-org", org, { x: 0.9, y: 5.65, w: 8, h: 0.35, fontSize: 13 });
    text(s, "presenter-name", name, { x: 0.9, y: 6.05, w: 8, h: 0.45, fontSize: 18, bold: true });
    return s;
  }

  // sections: [{ title, items: [] }], up to three columns
  function contents(sections, heading = "목차") {
    startSection(heading);
    const s = pres.addSlide({ masterName: "CONTENTS", sectionTitle: section });
    s.addText(heading, { placeholder: "title" });
    const colW = 3.5;
    sections.forEach(({ title, items }, i) => {
      const x = 1.6 + i * (colW + 0.35);
      text(s, `toc-${i}-num`, String(i + 1).padStart(2, "0"), { x, y: 2.2, w: 0.9, h: 0.7, fontFace: HEAD, fontSize: 32, color: NAVY });
      s.addShape(pres.shapes.LINE, { objectName: `toc-${i}-rule`, x: x + 0.95, y: 2.25, w: 0, h: 2.9, line: { color: C.accent2, width: 0.75 } });
      text(s, `toc-${i}-h`, title, { x: x + 1.15, y: 2.25, w: colW - 1.2, h: 0.45, fontFace: HEAD, fontSize: 18, bold: true, color: NAVY });
      bullets(s, `toc-${i}-list`, items, { x: x + 1.15, y: 2.85, w: colW - 1.2, h: 2.3, color: C.accent3, paraSpaceAfter: 6 });
    });
    return s;
  }

  // Starts a section; content slides after it belong to it.
  function divider(num, title, items = []) {
    startSection(`${num} ${title}`);
    const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: section });
    s.addText(num, { placeholder: "num" });
    s.addText(title, { placeholder: "title" });
    if (items.length) bullets(s, "divider-items", items, { x: 3.8, y: 4.2, w: 6, h: 1.4, color: C.accent3, paraSpaceAfter: 4 });
    return s;
  }

  function content(title, { lead } = {}) {
    if (!section) startSection("본문");
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    if (lead) text(s, "lead", lead, { x: MX, y: 1.15, w: CW, h: 0.75, fontSize: 15, lineSpacingMultiple: 1.2 });
    return s;
  }

  function closing({ title = "감사합니다", subtitle, org = who.org, name = who.who, contact = who.contact } = {}) {
    startSection("마무리");
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    if (subtitle) text(s, "close-sub", subtitle, { x: 1.2, y: 2.65, w: 9, h: 0.45, fontSize: 16, color: C.background2 });
    text(s, "close-org", org, { x: 1.2, y: 5.15, w: 8, h: 0.35, fontSize: 13, color: C.background2 });
    text(s, "close-name", name, { x: 1.2, y: 5.55, w: 8, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    if (contact) text(s, "close-contact", contact, { x: 1.2, y: 6.05, w: 8, h: 0.35, fontSize: 12, color: C.background2 });
    return s;
  }

  // ---------- components (positions in inches on a 13.33 x 7.5 page) ----------

  // "| label" heading inside a content area
  function barLabel(slide, name, label, x, y, w) {
    slide.addShape(pres.shapes.LINE, { objectName: name + "-bar", x, y: y + 0.04, w: 0, h: 0.26, line: { color: NAVY, width: 2 } });
    text(slide, name, label, { x: x + 0.12, y, w: w - 0.12, h: 0.34, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
  }

  // Bordered box with a tab label on its top edge. opts.dark fills it navy.
  function tabBox(slide, name, x, y, w, h, label, opts = {}) {
    const fill = opts.dark ? NAVY : opts.plain ? C.background1 : TINT;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      objectName: name, x, y, w, h, rectRadius: 0.05,
      fill: { color: fill }, line: { color: opts.dark ? NAVY : C.accent2, width: 0.75 },
    });
    if (label) {
      const lw = Math.max(1.3, label.length * 0.17 + 0.5);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        objectName: name + "-tab", x: x + 0.2, y: y - 0.2, w: lw, h: 0.4, rectRadius: 0.04,
        fill: { color: C.background1 }, line: { color: NAVY, width: 0.75 },
      });
      text(slide, name + "-tabtext", label, { x: x + 0.2, y: y - 0.2, w: lw, h: 0.4, fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
    }
  }

  // Big number in a tabbed box.
  function stat(slide, name, x, y, w, label, value, sub) {
    tabBox(slide, name, x, y, w, 1.6, label);
    text(slide, name + "-num", value, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.8, fontFace: HEAD, fontSize: 36, bold: true, color: NAVY, align: "center", valign: "middle" });
    if (sub) text(slide, name + "-sub", sub, { x: x + 0.25, y: y + 1.03, w: w - 0.5, h: 0.4, fontSize: 12, color: C.accent3, align: "center" });
  }

  // Navy numbered circle on a pale pill; tag is a muted suffix.
  function pill(slide, name, n, x, y, w, str, tag) {
    const d = 0.56;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name + "-pill", x: x + d / 2, y: y + 0.05, w: w - d / 2, h: d - 0.1, rectRadius: 0.23, fill: { color: TINT }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { objectName: name + "-num", x, y, w: d, h: d, fill: { color: NAVY }, line: { type: "none" } });
    text(slide, name + "-numtext", String(n).padStart(2, "0"), { x, y, w: d, h: d, fontFace: HEAD, fontSize: 16, color: C.background1, align: "center", valign: "middle" });
    const runs = [{ text: str, options: { fontSize: 14, color: C.text1 } }];
    if (tag) runs.push({ text: "   " + tag, options: { fontSize: 12, color: C.accent3 } });
    text(slide, name + "-text", runs, { x: x + d + 0.2, y, w: w - d - 0.35, h: d, valign: "middle" });
  }

  // Dark label block beside a light box; more than one item gets ①②③.
  function summary(slide, name, x, y, w, h, label, items) {
    const lw = 1.5;
    slide.addShape(pres.shapes.RECTANGLE, { objectName: name + "-label", x, y, w: lw, h, fill: { color: C.accent4 }, line: { type: "none" } });
    text(slide, name + "-labeltext", label, { x, y, w: lw, h, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    slide.addShape(pres.shapes.RECTANGLE, { objectName: name + "-box", x: x + lw, y, w: w - lw, h, fill: { color: "F4F4F4" }, line: { color: C.accent4, width: 0.75 } });
    slide.addText(items.map((t, i) => ({ text: items.length > 1 ? `${MARKS[i]}  ${t}` : t, options: { breakLine: i < items.length - 1 } })), {
      objectName: name + "-items", isTextBox: true, x: x + lw + 0.3, y, w: w - lw - 0.5, h, fontSize: 14, color: C.text1, valign: "middle", paraSpaceAfter: 4, margin: 0,
    });
  }

  // "Figure 1. ..." / "Table 1. ..."
  function caption(slide, name, str, x, y, w) {
    text(slide, name, str, { x, y, w, h: 0.3, fontSize: 11, bold: true });
  }

  function downArrow(slide, name, cx, y) {
    slide.addShape(pres.shapes.DOWN_ARROW, { objectName: name, x: cx - 0.45, y, w: 0.9, h: 0.38, fill: { color: "C9CED6" }, line: { type: "none" } });
  }

  function rightArrow(slide, name, cx, cy) {
    slide.addShape(pres.shapes.RIGHT_ARROW, { objectName: name, x: cx - 0.35, y: cy - 0.28, w: 0.7, h: 0.55, fill: { color: "C9CED6" }, line: { type: "none" } });
  }

  // Clustered column chart; series: [{ name, labels, values }].
  function barChart(slide, name, series, x, y, w, h, extra) {
    slide.addChart(pres.charts.BAR, series, Object.assign({
      objectName: name, x, y, w, h,
      barDir: "col", barGrouping: "clustered",
      chartColors: [HEX.accent2, HEX.dk2],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.accent3, valAxisLabelColor: HEX.accent3, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
      valGridLine: { color: HEX.accent6, size: 0.5 }, catGridLine: { style: "none" },
      showLegend: series.length > 1, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.dk1,
    }, extra));
  }

  // Navy-ruled table. A cell is a string or { text, options }.
  // `high(text)` gives a cell in the emphasis colour.
  function table(slide, name, header, rows, x, y, w, colW, { rowH = 0.5 } = {}) {
    const hair = { type: "solid", pt: 0.75, color: HEX.accent6 };
    const none = { type: "none" };
    const head = header.map((h) => ({
      text: h,
      options: { bold: true, color: HEX.dk2, fill: { color: HEX.lt2 }, align: "center", border: [{ type: "solid", pt: 1.5, color: HEX.dk2 }, none, hair, none] },
    }));
    const body = rows.map((r) => r.map((c, j) => {
      const cell = typeof c === "string" ? { text: c, options: {} } : { text: c.text, options: Object.assign({}, c.options) };
      cell.options = Object.assign({ align: j === 0 ? "center" : "left", border: [none, none, hair, none] }, cell.options);
      return cell;
    }));
    slide.addTable([head].concat(body), { objectName: name, x, y, w, colW, fontSize: 13, color: HEX.dk1, valign: "middle", margin: [4, 8, 4, 8], rowH });
  }

  const high = (t) => ({ text: t, options: { bold: true, color: HEX.accent5 } });
  const strong = (t) => ({ text: t, options: { bold: true, color: HEX.dk2 } });

  async function save(fileName) {
    await pres.writeFile({ fileName });
    await applyTheme(fileName, THEME);
    return fileName;
  }

  return {
    pres, C, W, H, MX, CW,
    cover, contents, divider, content, closing,
    text, bullets, barLabel, tabBox, stat, pill, summary, caption, downArrow, rightArrow, barChart, table, high, strong,
    save,
  };
}

module.exports = { createDeck, DEFAULT_THEME, PRESENTER_FIELDS, PRESENTER_FILE };
