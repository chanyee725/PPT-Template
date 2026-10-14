// Report deck: white pages, Pretendard, a coloured band, tabbed boxes,
// numbered pills and a dark "Summary" block. Every colour is a theme colour,
// so PowerPoint's Design > Variants > Colors recolours the whole deck.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply-theme");

// Palettes. dk2 is the main colour (bands, headings, numbers, progress) and
// accent1 a shade brighter, used only for the phase in progress; lt2 fills boxes,
// accent2 draws borders, accent6 hairlines. createDeck({ palette }) picks one.
const PALETTES = {
  // Light blue after the reference workflow slide.
  softBlue: { dk2: "4A6BD1", accent1: "6E8AE6", lt2: "EEF2FC", accent2: "AEBDE8", accent6: "E1E7F6" },
  // Greyed blue, calmer than softBlue.
  slate: { dk2: "4F6F96", accent1: "6E8DB3", lt2: "EFF3F7", accent2: "AFC0D3", accent6: "E2E8EF" },
  // Blue-green.
  teal: { dk2: "2F8291", accent1: "4BA0AE", lt2: "EDF5F6", accent2: "A6CBD1", accent6: "DDEBED" },
  // Deep navy. The default.
  navy: { dk2: "012D6A", accent1: "1F4E94", lt2: "EEF2F8", accent2: "8A9BB8", accent6: "D9DEE7" },
};

const BASE_COLORS = {
  dk1: "222222", // body text
  lt1: "FFFFFF", // page
  accent3: "6E6E6E", // muted text
  accent4: "262626", // summary block
  accent5: "C0392B", // "high" emphasis
  folHlink: "6E6E6E",
};

function paletteColors(name) {
  const p = PALETTES[name];
  if (!p) throw new Error(`unknown palette "${name}"; one of ${Object.keys(PALETTES).join(", ")}`);
  return Object.assign({}, BASE_COLORS, p, { hlink: p.dk2 });
}

const DEFAULT_THEME = {
  name: "Report",
  headFontFace: "Pretendard",
  bodyFontFace: "Pretendard",
  colors: paletteColors("navy"),
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

// "회사 · 부서" and "직책 이름" lines, or placeholders where nothing is set.
function presenterLines(p) {
  const join = (...parts) => parts.filter(Boolean).join(" · ");
  return {
    org: join(p.company, p.department) || "회사 · 부서",
    who: [p.position, p.name].filter(Boolean).join(" ") || "직책 이름",
    contact: join(p.email, p.phone),
  };
}

// usePresenterFile: false leaves presenter.json out, for decks that are
// committed (templates/) and so must keep placeholders.
// palette: a PALETTES name; theme.colors still overrides single slots.
function createDeck({ title, header, theme, palette = "navy", presenter, usePresenterFile = true } = {}) {
  const THEME = Object.assign({}, DEFAULT_THEME, { name: `Report ${palette}` }, theme, {
    colors: Object.assign({}, paletteColors(palette), theme && theme.colors),
  });
  const HEX = THEME.colors;
  const runningHeader = header || title || "";
  const who = presenterLines(Object.assign({}, usePresenterFile ? loadPresenter() : {}, presenter));

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
      { rect: { x: 0, y: 6.95, w: W, h: 0.55, fill: { color: NAVY }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.9, y: 0.73, w: 10.5, h: 1.6, fontFace: HEAD, fontSize: 40, lineSpacingMultiple: 1.1, bold: true, color: NAVY, valign: "top", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", align: "left", x: 0.9, y: 2.53, w: 10.5, h: 0.45, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENTS",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 0, y: 0, w: 0.45, h: H, fill: { color: NAVY }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 0.9, y: 0.4, w: 5, h: 0.6, fontFace: HEAD, fontSize: 26, bold: true, color: NAVY, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: runningHeader, options: { x: W - 6.6, y: 0.5, w: 6.0, h: 0.4, fontFace: HEAD, fontSize: 11, bold: true, color: NAVY, align: "right", margin: 0 } } },
    ],
  });

  pres.defineSlideMaster({
    title: "DIVIDER",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 0.9, y: 0, w: 2.2, h: H, fill: { color: NAVY }, line: { type: "none" } } },
      { line: { x: 3.1, y: 3.95, w: W - 3.1, h: 0, line: { color: C.accent2, width: 0.75 } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 3.6, y: 3.1, w: 8.5, h: 0.8, fontFace: HEAD, fontSize: 32, bold: true, color: NAVY, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "num", type: "body", x: 0.9, y: 2.6, w: 2.2, h: 1.5, fontSize: 72, color: C.background1, align: "center", valign: "middle", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { text: { text: runningHeader, options: { x: W - 6.6, y: 0.38, w: 6.0, h: 0.35, fontFace: HEAD, fontSize: 11, bold: true, color: NAVY, align: "right", margin: 0 } } },
      { rect: { x: 0, y: 7.28, w: W, h: 0.22, fill: { color: NAVY }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: MX, y: 0.3, w: 6.5, h: 0.55, fontFace: HEAD, fontSize: 24, bold: true, color: C.text1, valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: W - 0.95, y: 6.82, w: 0.35, h: 0.28, fontSize: 10, color: C.accent3, align: "center" },
  });

  // Same frame as CONTENT with neither title nor running header: the slide
  // puts its own name where the header goes (the flow chart slide).
  pres.defineSlideMaster({
    title: "FEATURE",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 0, y: 7.28, w: W, h: 0.22, fill: { color: NAVY }, line: { type: "none" } } },
    ],
    slideNumber: { x: W - 0.95, y: 6.82, w: 0.35, h: 0.28, fontSize: 10, color: C.accent3, align: "center" },
  });

  pres.defineSlideMaster({
    title: "CLOSING",
    background: { color: NAVY },
    objects: [
      { rect: { x: 0.89, y: 1.6, w: 0.014, h: 1.5, fill: { color: C.background1 }, line: { type: "none" } } },
      { placeholder: { options: { name: "title", type: "title", align: "left", x: 1.2, y: 1.6, w: 9, h: 1.0, fontFace: HEAD, fontSize: 48, bold: true, color: C.background1, valign: "top", margin: 0 }, text: "" } },
    ],
  });

  // ---------- primitives ----------

  // A vertical line from (x, y) down by h. Drawn as a level line turned 90
  // degrees: a zero-width line converts to a slanted "\\" in some tools. A
  // beginArrowType lands at the top, an endArrowType at the bottom.
  function vline(slide, name, x, y, h, line) {
    slide.addShape(pres.shapes.LINE, { objectName: name, x: x - h / 2, y: y + h / 2, w: h, h: 0, rotate: 90, line });
  }
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
  // Presenter lines come from presenter.json unless given here. The date sits
  // bottom-right, on the same baseline as the name.
  function cover({ title, subtitle, date, org = who.org, name = who.who }) {
    startSection("표지");
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: section });
    // A little air between the two title lines.
    s.addText(title, { placeholder: "title", align: "left", lineSpacingMultiple: 1.1 });
    if (subtitle) s.addText(subtitle, { placeholder: "body", align: "left" });
    text(s, "presenter-org", org, { x: 0.9, y: 5.62, w: 8, h: 0.32, fontSize: 13, bold: true, valign: "bottom" });
    text(s, "presenter-name", name, { x: 0.9, y: 5.98, w: 8, h: 0.42, fontSize: 17, bold: true, valign: "bottom" });
    if (date) text(s, "date", date, { x: W - 0.9 - 3, y: 5.98, w: 3, h: 0.42, fontSize: 13, bold: true, align: "right", valign: "bottom" });
    return s;
  }

  // sections: [{ title, items: [] }], up to three columns
  function contents(sections, heading = "목차") {
    startSection(heading);
    const s = pres.addSlide({ masterName: "CONTENTS", sectionTitle: section });
    s.addText(heading, { placeholder: "title", align: "left" });
    const colW = 3.5;
    sections.forEach(({ title, items }, i) => {
      const x = 1.6 + i * (colW + 0.35);
      text(s, `toc-${i}-num`, String(i + 1).padStart(2, "0"), { x, y: 2.2, w: 0.9, h: 0.7, fontFace: HEAD, fontSize: 32, color: NAVY });
      vline(s, `toc-${i}-rule`, x + 0.95, 2.25, 2.9, { color: C.accent2, width: 0.75 });
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
    s.addText(title, { placeholder: "title", align: "left" });
    if (items.length) bullets(s, "divider-items", items, { x: 3.8, y: 4.2, w: 6, h: 1.4, color: C.accent3, paraSpaceAfter: 4 });
    return s;
  }

  // `section` starts a new section without a divider slide.
  function content(title, { lead, section: name } = {}) {
    if (name) startSection(name);
    else if (!section) startSection("본문");
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title, { placeholder: "title", align: "left" });
    if (lead) text(s, "lead", lead, { x: MX, y: 1.0, w: CW, h: 0.7, fontSize: 13.5, lineSpacingMultiple: 1.2 });
    return s;
  }

  function closing({ title = "감사합니다", subtitle, org = who.org, name = who.who, contact = who.contact } = {}) {
    startSection("마무리");
    const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: section });
    s.addText(title, { placeholder: "title", align: "left" });
    if (subtitle) text(s, "close-sub", subtitle, { x: 1.2, y: 2.65, w: 9, h: 0.45, fontSize: 16, color: C.background2 });
    text(s, "close-org", org, { x: 1.2, y: 5.15, w: 8, h: 0.35, fontSize: 13, color: C.background2 });
    text(s, "close-name", name, { x: 1.2, y: 5.55, w: 8, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    if (contact) text(s, "close-contact", contact, { x: 1.2, y: 6.05, w: 8, h: 0.35, fontSize: 12, color: C.background2 });
    return s;
  }

  // ---------- components (positions in inches on a 13.33 x 7.5 page) ----------

  // "| label" heading inside a content area
  function barLabel(slide, name, label, x, y, w) {
    slide.addShape(pres.shapes.RECTANGLE, { objectName: name + "-bar", x: x - 0.014, y: y + 0.04, w: 0.028, h: 0.26, fill: { color: NAVY }, line: { type: "none" } });
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

  // "Figure 1. ..." / "Table 1. ..." -- small and light, under the thing.
  function caption(slide, name, str, x, y, w) {
    text(slide, name, str, { x, y, w, h: 0.28, fontSize: 10, color: C.accent3 });
  }

  function downArrow(slide, name, cx, y) {
    slide.addShape(pres.shapes.DOWN_ARROW, { objectName: name, x: cx - 0.45, y, w: 0.9, h: 0.38, fill: { color: "C9CED6" }, line: { type: "none" } });
  }

  function rightArrow(slide, name, cx, cy) {
    slide.addShape(pres.shapes.RIGHT_ARROW, { objectName: name, x: cx - 0.35, y: cy - 0.28, w: 0.7, h: 0.55, fill: { color: "C9CED6" }, line: { type: "none" } });
  }

  // Type sizes for a row of n phases: smaller as the columns narrow.
  function phaseSizes(n) {
    if (n <= 4) return { label: 11, title: 14, desc: 11 };
    if (n <= 6) return { label: 11, title: 13, desc: 10.5 };
    return { label: 10, title: 12, desc: 10 };
  }

  // Rough rendered width in inches: Hangul and other wide glyphs are about one
  // em, Latin about half. Good enough to decide whether a title needs two lines.
  function textWidth(str, pt) {
    let em = 0;
    for (const ch of str) em += ch.charCodeAt(0) > 0x2e80 ? 1 : ch === " " ? 0.3 : 0.56;
    return (em * pt) / 72;
  }

  // Horizontal phase timeline: "Phase n" label, a dot with a soft halo on one
  // line, the phase title and a short description, all centred per column.
  // phases: [{ label, title, desc }]; `current` (1-based) is the phase
  // in progress -- drawn a shade brighter, later phases grey, and the line in the
  // main colour up to it. Up to 8 phases; type shrinks past 4, a long title takes
  // two lines, and every description starts at the same height.
  function phaseFlow(slide, name, phases, { current = 1, x = MX, y = 2.87, w = CW } = {}) {
    const n = phases.length;
    if (n > 8) throw new Error(`phaseFlow: ${n} phases is too many for one row (8 at most); split them over two slides`);
    const colW = w / n;
    const size = phaseSizes(n);
    const cx = (i) => x + colW * (i + 0.5);
    const lineY = y + 0.62;
    const state = (i) => (i + 1 < current ? "done" : i + 1 === current ? "current" : "planned");

    const titleW = colW - 0.2;
    const titleLines = phases.some((p) => textWidth(p.title, size.title) > titleW - 0.05) ? 2 : 1;
    const titleH = (titleLines * size.title * 1.25) / 72 + 0.06;
    const titleY = lineY + 0.4;
    const descY = titleY + titleH + 0.14;

    slide.addShape(pres.shapes.LINE, { objectName: name + "-line", x: cx(0), y: lineY, w: cx(n - 1) - cx(0), h: 0, line: { color: C.accent2, width: 1.5 } });
    if (current > 1) {
      const end = cx(Math.min(current, n) - 1);
      slide.addShape(pres.shapes.LINE, { objectName: name + "-progress", x: cx(0), y: lineY, w: end - cx(0), h: 0, line: { color: NAVY, width: 2 } });
    }

    const circle = (id, c, d, fill) => slide.addShape(pres.shapes.OVAL, {
      objectName: id, x: c - d / 2, y: lineY - d / 2, w: d, h: d, fill, line: { type: "none" },
    });

    phases.forEach((p, i) => {
      const c = cx(i);
      const st = state(i);
      const id = `${name}-${i}`;
      const planned = st === "planned";
      const label = p.label || `Phase ${i + 1}`;

      text(slide, id + "-label", label, { x: c - colW / 2, y, w: colW, h: 0.32, fontSize: size.label, bold: true, color: planned ? C.accent3 : C.text1, align: "center", valign: "middle" });

      // Same circles for every phase; the one in progress is a shade brighter.
      const ink = st === "current" ? C.accent1 : NAVY;
      circle(id + "-halo", c, 0.34, planned ? { color: C.accent6 } : { color: ink, transparency: 60 });
      circle(id + "-dot", c, 0.22, { color: planned ? C.accent2 : ink });

      text(slide, id + "-title", p.title, { x: c - titleW / 2, y: titleY, w: titleW, h: titleH, fontSize: size.title, bold: true, color: planned ? C.accent3 : C.text1, align: "center", valign: "top" });
      if (p.desc) {
        text(slide, id + "-desc", p.desc, { x: c - colW / 2 + 0.12, y: descY, w: colW - 0.24, h: 6.6 - descY, fontSize: size.desc, color: C.accent3, align: "center", lineSpacingMultiple: 1.3 });
      }
    });
  }

  // ---------- phases: one overview slide, then detail slides per phase ----------
  let plan = null;
  let figure = 0;

  // The flow chart slide: `title` top-right where the running header goes
  // (placed exactly like it), the phase timeline centred below. Remembers the
  // phases for phaseDetail unless `remember: false` -- e.g. a translated copy
  // that the detail slides should not take their titles from.
  function phaseOverview(title, phases, { current = 1, remember = true } = {}) {
    if (remember || !plan) plan = { phases, current };
    startSection(title);
    const s = pres.addSlide({ masterName: "FEATURE", sectionTitle: section });
    text(s, "header", title, { x: W - 6.6, y: 0.38, w: 6.0, h: 0.35, fontFace: HEAD, fontSize: 11, bold: true, color: NAVY, align: "right", valign: "middle" });
    phaseFlow(s, "phases", phases, { current });
    return s;
  }

  // Width and height in pixels of a PNG or JPEG (EXIF rotation is ignored).
  function imageSize(file) {
    const b = fs.readFileSync(file);
    if (b.readUInt32BE(0) === 0x89504e47) return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
    if (b[0] === 0xff && b[1] === 0xd8) {
      for (let i = 2; i + 9 < b.length && b[i] === 0xff; i += 2 + b.readUInt16BE(i + 2)) {
        const m = b[i + 1];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { w: b.readUInt16BE(i + 7), h: b.readUInt16BE(i + 5) };
      }
    }
    throw new Error(`${file}: only PNG and JPEG are supported`);
  }

  // A photo fitted inside the box, centred across and top-aligned (no frame:
  // what the photo does not fill stays page white), or a crossed placeholder
  // box when there is no path yet. A caption becomes "Figure n. caption" below the box.
  function photo(slide, name, x, y, w, h, { path: file, caption: cap } = {}) {
    // The caption sits right under the picture itself, not the box around it.
    let capX = x;
    let capY = y + h + 0.08;
    let capW = w;
    if (file) {
      const { w: pw, h: ph } = imageSize(file);
      const scale = Math.min(w / pw, h / ph);
      const iw = pw * scale;
      const ih = ph * scale;
      // Centred across, top-aligned: the picture starts level with the text
      // beside it, and any spare height ends up below it.
      slide.addImage({ objectName: name, path: file, x: x + (w - iw) / 2, y, w: iw, h: ih });
      capX = x + (w - iw) / 2;
      capY = y + ih + 0.08;
      capW = Math.max(iw, Math.min(w, 3));
    } else {
      slide.addShape(pres.shapes.RECTANGLE, { objectName: name + "-frame", x, y, w, h, fill: { color: "F4F5F7" }, line: { color: C.accent6, width: 1 } });
      slide.addShape(pres.shapes.LINE, { objectName: name + "-x1", x, y, w, h, line: { color: C.accent6, width: 1 } });
      slide.addShape(pres.shapes.LINE, { objectName: name + "-x2", x, y, w, h, flipV: true, line: { color: C.accent6, width: 1 } });
      text(slide, name + "-hint", "사진", { x: x + w / 2 - 0.6, y: y + h / 2 - 0.2, w: 1.2, h: 0.4, fontSize: 13, color: C.accent3, align: "center", valign: "middle", fill: { color: "F4F5F7" } });
    }
    if (cap) caption(slide, name + "-caption", `Figure ${++figure}. ${cap}`, capX, capY, capW);
  }

  // A slim arrow between pictures that read as steps, centred on (cx, cy):
  // one hairline with an open head. Drawn as a plain line -- no flips, which
  // some editors ignore.
  function flowArrow(slide, name, cx, cy, len = 0.36) {
    slide.addShape(pres.shapes.LINE, {
      objectName: name, x: cx - len / 2, y: cy, w: len, h: 0,
      line: { color: NAVY, width: 1.5, endArrowType: "arrow" },
    });
  }

  // One slide about phase `n` (1-based): a lead, then photos and labelled
  // bullet groups -- one or two photos on the left with the groups beside them,
  // or three photos in a row with the groups side by side underneath. Call it again for a second
  // slide on the same phase, with a `title` to tell them apart.
  //   photos:   [{ path, caption }]  -- one or two; a missing path draws a placeholder
  //   sections: [{ title, items: [] }]
  //   flow:     true puts a ">" between three photos, read as steps
  //   photos: [] leaves the left area (MX, 2.0, 7.1 x 4.2) empty for the
  //   caller -- e.g. a calendar.
  //   textOnly: true -- no photos at all; the groups sit side by side across
  //   the full width.
  function phaseDetail(n, { title, lead, photos = [{}], sections = [], flow = false, textOnly = false } = {}) {
    if (!plan) throw new Error("phaseDetail needs phaseOverview first");
    const p = plan.phases[n - 1];
    const name = `Phase ${n} ${p.title}`;
    if (section !== name) startSection(name);
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title || p.title, { placeholder: "title", align: "left" });
    if (lead) text(s, "lead", lead, { x: MX, y: 1.0, w: CW, h: 0.55, fontSize: 13.5 });

    const size = 13;
    const listHeight = (items, w) => {
      const lines = items.reduce((k, item) => k + Math.max(1, Math.ceil(textWidth(item, size) / (w - 0.45))), 0);
      return (lines * size * 1.25 + items.length * 4) / 72 + 0.1;
    };
    const group = (sec, i, x, y, w) => {
      const h = listHeight(sec.items, w);
      barLabel(s, `sec-${i}`, sec.title, x, y, w);
      bullets(s, `sec-${i}-list`, sec.items, { x: x + 0.12, y: y + 0.45, w: w - 0.12, h, fontSize: size, paraSpaceAfter: 4 });
      return 0.45 + h;
    };
    const gap = 0.25;

    if (textOnly) {
      const colW = (CW - 0.5 * (sections.length - 1)) / Math.max(1, sections.length);
      sections.forEach((sec, i) => group(sec, i, MX + i * (colW + 0.5), 2.0, colW));
      return s;
    }

    // Three photos: a row across the full width, the groups side by side below.
    // With `flow`, wider gaps hold a ">" between them, level with the pictures.
    // Without sections the row is a little narrower and sits in the middle of
    // the page, photos only.
    if (photos.length >= 3) {
      const shown = photos.slice(0, 3);
      const alone = sections.length === 0;
      const rowGap = flow ? 0.6 : gap;
      const rowW = alone ? CW - 0.8 : CW;
      const each = (rowW - rowGap * 2) / 3;
      const first = shown[0].path && imageSize(shown[0].path);
      const boxH = first ? (each * first.h) / first.w : 2.9;
      const x0 = MX + (CW - rowW) / 2;
      const top = alone ? 1.6 + (6.75 - 1.6 - boxH - 0.35) / 2 : 1.7;
      shown.forEach((ph, i) => photo(s, `photo-${i}`, x0 + i * (each + rowGap), top, each, boxH, ph));
      if (flow) {
        for (let i = 0; i < 2; i++) {
          const gx = x0 + (i + 1) * each + i * rowGap;
          flowArrow(s, `flow-${i}`, gx + rowGap / 2, top + boxH / 2);
        }
      }
      const colW = (CW - 0.5 * (sections.length - 1)) / Math.max(1, sections.length);
      sections.forEach((sec, i) => group(sec, i, MX + i * (colW + 0.5), 5.0, colW));
      return s;
    }

    // One or two photos on the left, groups stacked on the right -- each as
    // tall as its (wrapped) items, so a second group (e.g. the same items in
    // English) sits right under the first.
    const top = 2.0;
    const picW = 7.1;
    const picH = 4.2;
    if (photos.length) {
      const each = (picW - gap * (photos.length - 1)) / photos.length;
      photos.forEach((ph, i) => photo(s, `photo-${i}`, MX + i * (each + gap), top, each, picH, ph));
    }

    const rx = MX + picW + 0.45;
    const rw = W - MX - rx;
    let y = top;
    sections.forEach((sec, i) => { y += group(sec, i, rx, y, rw) + 0.35; });
    return s;
  }

  // A month on a white outlined card (no shadow), weeks starting on Sunday: a big month name,
  // weekday initials, centred day numbers with no grid (the tail of the
  // previous month in light grey). `marks` lay a soft rounded bar under a run
  // of days inside one week, colour those dates and write a note under them:
  //   [{ from: 19, to: 23, label: "카메라 · 그리퍼 입고 예정" }]
  // Marks use the emphasis colour (accent5).
  function calendar(slide, name, year, month, x, y, w, h, { marks = [] } = {}) {
    const days = new Date(year, month, 0).getDate();
    const prevDays = new Date(year, month - 1, 0).getDate();
    const lead0 = new Date(year, month - 1, 1).getDay();
    const rows = Math.ceil((lead0 + days) / 7);
    const pad = 0.3;
    const headH = 0.7;
    const dowH = 0.34;
    const gx = x + pad;
    const gw = w - pad * 2;
    const cellW = gw / 7;
    const gridY = y + pad * 0.6 + headH + dowH;
    const cellH = (y + h - pad * 0.6 - gridY) / rows;
    const marked = (d) => marks.find((m) => d >= m.from && d <= m.to);
    const at = (k) => ({ cx: gx + (k % 7) * cellW, cy: gridY + Math.floor(k / 7) * cellH, col: k % 7 });

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      objectName: name + "-card", x, y, w, h, rectRadius: 0.12,
      fill: { color: C.background1 }, line: { color: C.accent6, width: 0.75 },
    });
    text(slide, name + "-month", [
      { text: `${month}월`, options: { fontSize: 24, bold: true, color: NAVY } },
      { text: `   ${year}`, options: { fontSize: 13, color: C.accent3 } },
    ], { x: gx, y: y + pad * 0.6, w: gw, h: headH, valign: "middle" });

    ["일", "월", "화", "수", "목", "금", "토"].forEach((dn, i) => {
      text(slide, `${name}-dow-${i}`, dn, { x: gx + i * cellW, y: gridY - dowH, w: cellW, h: dowH, fontSize: 10.5, bold: true, color: i === 0 ? C.accent5 : C.accent3, align: "center", valign: "middle" });
    });

    const numH = 0.3;
    const barDrop = 0.05; // the bar sits a little below the numbers, like a highlighter
    marks.forEach((m, i) => {
      const a = at(lead0 + m.from - 1);
      const b = at(lead0 + m.to - 1);
      const bx = a.cx + 0.06;
      const bw = b.cx + cellW - a.cx - 0.12;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        objectName: `${name}-mark-${i}`, x: bx, y: a.cy + 0.05 + barDrop, w: bw, h: numH, rectRadius: numH / 2,
        fill: { color: C.accent5, transparency: 85 }, line: { type: "none" },
      });
      if (m.label) {
        text(slide, `${name}-mark-${i}-label`, m.label, { x: bx, y: a.cy + 0.05 + numH + barDrop, w: bw, h: cellH - numH - 0.04 - barDrop, fontSize: 10.5, bold: true, color: C.accent5, align: "center", valign: "top" });
      }
    });

    for (let k = 0; k < rows * 7; k++) {
      const d = k - lead0 + 1;
      if (d > days) break;
      const { cx, cy, col } = at(k);
      const inMonth = d >= 1;
      const mk = inMonth && marked(d);
      text(slide, `${name}-d${k}`, String(inMonth ? d : prevDays + d), {
        x: cx, y: cy + 0.05, w: cellW, h: numH, fontSize: 12, bold: Boolean(mk), align: "center", valign: "middle",
        color: !inMonth ? C.accent6 : mk ? C.accent5 : col === 0 ? C.accent5 : C.text1,
      });
    }
  }

  // A box with its text inside the same shape, centred both ways. Safer than a
  // text box laid over a shape: editors that add their own line spacing push
  // an overlaid text box down and out of the shape, but always centre text
  // that belongs to the shape.
  function labelBox(slide, name, str, x, y, w, h, { shape = "roundRect", fill = C.background1, line = { color: NAVY, width: 0.75 }, radius = 0.05, ...textOpts } = {}) {
    const kind = shape === "rect" ? pres.shapes.RECTANGLE : shape === "ellipse" ? pres.shapes.OVAL : pres.shapes.ROUNDED_RECTANGLE;
    slide.addText(str, Object.assign({
      objectName: name, shape: kind, x, y, w, h, rectRadius: shape === "roundRect" ? radius : undefined,
      fill: fill ? { color: fill } : { type: "none" }, line: line || { type: "none" },
      // pptxgenjs reads a margin array as [left, right, bottom, top] (points).
      // 8 pt at the bottom lifts the text: editors that set Korean with a
      // taller line than PowerPoint otherwise show it low in the box.
      margin: [0, 0, 8, 0], fontSize: 12, color: C.text1, align: "center", valign: "middle",
    }, textOpts));
  }

  // An embedded video. `cover` is a PNG shown before it plays (the tool stores
  // the poster as PNG, so it must be one); a caption sits under it.
  function video(slide, name, file, x, y, w, h, { cover, caption: cap } = {}) {
    const poster = cover ? "data:image/png;base64," + fs.readFileSync(cover).toString("base64") : undefined;
    slide.addMedia({ objectName: name, type: "video", path: file, x, y, w, h, cover: poster });
    if (cap) caption(slide, name + "-caption", cap, x, y + h + 0.08, w);
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

  // Navy-ruled table: a navy rule on top, a pale header row shorter than the
  // body rows, thin rules between rows. A cell is a string or { text, options };
  // `high(text)` / `strong(text)` colour one. `align` gives each column's
  // alignment (header and body alike); by default the first column is centred.
  function table(slide, name, header, rows, x, y, w, colW, { rowH = 0.5, headH = 0.4, fontSize = 13, align } = {}) {
    const hair = { type: "solid", pt: 0.75, color: HEX.accent6 };
    const none = { type: "none" };
    const colAlign = (j) => (align ? align[j] : j === 0 ? "center" : "left");
    const head = header.map((h, j) => ({
      text: h,
      options: { bold: true, color: HEX.dk2, fill: { color: HEX.lt2 }, align: colAlign(j), border: [{ type: "solid", pt: 1.5, color: HEX.dk2 }, none, hair, none] },
    }));
    const body = rows.map((r) => r.map((c, j) => {
      const cell = typeof c === "string" ? { text: c, options: {} } : { text: c.text, options: Object.assign({}, c.options) };
      cell.options = Object.assign({ align: colAlign(j), border: [none, none, hair, none] }, cell.options);
      return cell;
    }));
    slide.addTable([head].concat(body), {
      objectName: name, x, y, w, colW, fontSize, color: HEX.dk1, valign: "middle", margin: [4, 8, 4, 8],
      rowH: [headH].concat(rows.map(() => rowH)),
    });
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
    cover, contents, divider, content, closing, phaseOverview, phaseDetail,
    text, bullets, vline, barLabel, tabBox, stat, pill, summary, caption, downArrow, rightArrow, phaseFlow, photo, video, flowArrow, labelBox, calendar, barChart, table, high, strong,
    save,
  };
}

module.exports = { createDeck, DEFAULT_THEME, PALETTES, PRESENTER_FIELDS, PRESENTER_FILE };
