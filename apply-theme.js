// pptxgenjs cannot write theme colours, and it leaves the East Asian theme
// fonts to Office's defaults, so Korean headings come out in 맑은 고딕 rather
// than the heading font. This rewrites the saved file's theme and font refs.
const fs = require("fs");
const JSZip = require("jszip");

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

const escape = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function themeXml(xml, theme) {
  const scheme = SLOTS.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k]}"/></a:${k}>`).join("");
  xml = xml.replace(/<a:clrScheme [^>]*>[\s\S]*?<\/a:clrScheme>/, `<a:clrScheme name="${escape(theme.name)}">${scheme}</a:clrScheme>`);
  const fonts = (tag, face) => {
    xml = xml.replace(new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`), (m, inner) => {
      inner = inner
        .replace(/<a:latin typeface="[^"]*"/, `<a:latin typeface="${escape(face)}"`)
        .replace(/<a:ea typeface="[^"]*"/, `<a:ea typeface="${escape(face)}"`)
        .replace(/<a:font script="Hang" typeface="[^"]*"/, `<a:font script="Hang" typeface="${escape(face)}"`);
      return `<a:${tag}>${inner}</a:${tag}>`;
    });
  };
  fonts("majorFont", theme.headFontFace);
  fonts("minorFont", theme.bodyFontFace);
  return xml;
}

// A run set to "+mj-lt" is written with the same token on <a:ea>, which is
// the Latin slot; the East Asian theme font is "+mj-ea".
function fontRefs(xml) {
  return xml
    .replace(/<a:ea typeface="\+mj-lt"/g, '<a:ea typeface="+mj-ea"')
    .replace(/<a:ea typeface="\+mn-lt"/g, '<a:ea typeface="+mn-ea"');
}

async function applyTheme(file, theme) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".xml")) continue;
    let xml = await zip.file(path).async("string");
    const before = xml;
    if (/^ppt\/theme\/theme\d+\.xml$/.test(path)) xml = themeXml(xml, theme);
    xml = fontRefs(xml);
    if (xml !== before) zip.file(path, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { applyTheme };
