/**
 * CTRL 180 Export Service · fill-template (Observer overlay)
 * Path: /pages/api/fill-template.js  (ctrl-180-service)
 *
 * Template:
 *   CTRL_Perspective_Assessment_Profile_template_slim_180.pdf
 */

export const config = { runtime: "nodejs" };

/* ───────────── imports ───────────── */
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── tiny utils ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0)   => (Number.isFinite(+v) ? +v : fb);

const norm = (v, fb = "") =>
  String(v ?? fb)
    .normalize("NFKC")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2010-\u2014]/g, "-")
    .replace(/\u2026/g, "...")
    .replace(/\u00A0/g, " ")
    .replace(/[•·]/g, "-")
    // arrows → WinAnsi-safe
    .replace(/\u2194/g, "<->").replace(/\u2192/g, "->").replace(/\u2190/g, "<-")
    .replace(/\u2191/g, "^").replace(/\u2193/g, "v").replace(/[\u2196-\u2199]/g, "->")
    .replace(/\u21A9/g, "<-").replace(/\u21AA/g, "->")
    .replace(/\u00D7/g, "x")
    // zero-width, emoji/PUA
    .replace(/[\u200B-\u200D\u2060]/g, "")
    .replace(/[\uD800-\uDFFF]/g, "")
    .replace(/[\uE000-\uF8FF]/g, "")
    // tidy
    .replace(/\t/g, " ").replace(/\r\n?/g, "\n")
    .replace(/[ \f\v]+/g, " ").replace(/[ \t]+\n/g, "\n").trim();

/* ───────────── payload helpers ───────────── */

function parseDataParam(b64ish) {
  if (!b64ish) return {};
  let s = String(b64ish);
  try { s = decodeURIComponent(s); } catch {}
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  try { return JSON.parse(Buffer.from(s, "base64").toString("utf8")); }
  catch { return {}; }
}

/* GET/POST payload reader (supports ?data= and JSON body) */
async function readPayload(req) {
  const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
  if (q.data) return parseDataParam(q.data);
  if (req.method === "POST" && !q.data) {
    try {
      return typeof req.json === "function" ? await req.json() : (req.body || {});
    } catch { /* fallthrough */ }
  }
  return {};
}

/* TL → simple textbox (used for headers, cover, etc.) */
function drawTextBox(page, font, text, spec = {}, opts = {}) {
  if (!page) return;
  const {
    x = 40,
    y = 40,
    w = 540,
    size = 12,
    lineGap = 3,
    color = rgb(0, 0, 0),
    align = "left",
    h,
  } = spec;

  const lineHeight = Math.max(1, size) + lineGap;
  const maxLines =
    opts.maxLines ??
    spec.maxLines ??
    (h ? Math.max(1, Math.floor(h / lineHeight)) : 40);

  const hard = norm(text || "");
  if (!hard) return;

  const lines = hard.split(/\n/).map((s) => s.trim());
  const wrapped = [];
  const widthOf = (s) => font.widthOfTextAtSize(s, Math.max(1, size));

  const wrapLine = (ln) => {
    const words = ln.split(/\s+/);
    let cur = "";
    for (let i = 0; i < words.length; i++) {
      const nxt = cur ? `${cur} ${words[i]}` : words[i];
      if (widthOf(nxt) <= w || !cur) cur = nxt;
      else {
        wrapped.push(cur);
        cur = words[i];
      }
    }
    wrapped.push(cur);
  };
  for (const ln of lines) wrapLine(ln);

  const out = wrapped.slice(0, maxLines);
  const pageH = page.getHeight();
  const baselineY = pageH - y;

  let yCursor = baselineY;
  for (const ln of out) {
    let xDraw = x;
    const wLn = widthOf(ln);
    if (align === "center") xDraw = x + (w - wLn) / 2;
    else if (align === "right") xDraw = x + (w - wLn);
    page.drawText(ln, {
      x: xDraw,
      y: yCursor - size,
      size: Math.max(1, size),
      font,
      color,
    });
    yCursor -= lineHeight;
  }
}

/* New: rich overlay drawer (bold headings like 'What this difference may mean:') */
function drawOverlayBox(page, fonts, text, spec = {}) {
  if (!page) return;
  const fontReg  = fonts.reg;
  const fontBold = fonts.bold;

  const {
    x = 40,
    y = 40,
    w = 540,
    size = 15,        // base body size (larger than default)
    lineGap = 6,
    color = rgb(0, 0, 0),
    align = "left",
    maxLines = 120
  } = spec;

  const hard = norm(text || "");
  if (!hard) return;

  const pageH      = page.getHeight();
  let   yCursor    = pageH - y;
  let   usedLines  = 0;

  // base line height used for spacing and "fake blank lines"
  const baseLineHeight = size + lineGap;

  const pushLine = (ln, { font, fSize, indent = 0 }) => {
    if (usedLines >= maxLines) return;
    const lineHeight = fSize + lineGap;
    const widthOf    = (s) => font.widthOfTextAtSize(s, fSize);

    // basic wrapping
    const words = ln.split(/\s+/);
    let current = "";
    const segments = [];

    for (let i = 0; i < words.length; i++) {
      const next = current ? current + " " + words[i] : words[i];
      if (widthOf(next) <= (w - indent) || !current) {
        current = next;
      } else {
        segments.push(current);
        current = words[i];
      }
    }
    if (current) segments.push(current);

    for (let s of segments) {
      if (usedLines >= maxLines) break;
      let xDraw = x + indent;
      const wLn = widthOf(s);
      if (align === "center") xDraw = x + (w - wLn) / 2;
      else if (align === "right") xDraw = x + (w - wLn);

      page.drawText(s, {
        x: xDraw,
        y: yCursor - fSize,
        size: fSize,
        font,
        color,
      });
      yCursor  -= lineHeight;
      usedLines++;
    }
  };

  const rawLines = hard.split(/\n+/);

  for (let raw of rawLines) {
    const line = raw.trim();
    if (!line) {
      // blank line = extra spacing
      yCursor -= size + lineGap;
      continue;
    }

    const isHeading = /^what\b.*:\s*$/i.test(line);
    const isBullet  = /^-\s+/.test(line);

    if (isHeading) {
      // BIG extra gap before any heading (≈ two blank lines)
      yCursor -= baseLineHeight * 2;

      pushLine(line.replace(/:\s*$/, ":"), {
        font: fontBold,
        fSize: size + 1,
        indent: 0
      });

      // small extra gap after the heading
      yCursor -= baseLineHeight * 0.5;
      continue;
    }

    if (isBullet) {
      const content = line.replace(/^-\s+/, "• ");
      pushLine(content, {
        font: fontReg,
        fSize: size,
        indent: 10    // small indent for bullets
      });
      continue;
    }

    // normal paragraph text
    pushLine(line, {
      font: fontReg,
      fSize: size,
      indent: 0
    });
  }
}

/* robust /public template loader */
async function loadTemplateBytesLocal(filename) {
  const fname = String(filename || "").trim();
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir  = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(__dir, "public", fname),
    path.join(process.cwd(), "public", fname),
    path.join(__dir, fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try { return await fs.readFile(pth); }
    catch (err) { lastErr = err; }
  }
  throw new Error(`Template not found for /public: ${fname} (${lastErr?.message || "no detail"})`);
}

/* safe page getter */
const pageOrNull = (pages, idx0) => (pages[idx0] ?? null);

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  try {
    const q = req.method === "POST" ? (req.body || {}) : (req.query || {});

    const defaultTpl = "CTRL_Perspective_Assessment_Profile_template_slim_180.pdf";
    const tpl        = S(q.tpl || defaultTpl).replace(/[^A-Za-z0-9._-]/g, "");
    const src        = await readPayload(req);

    const P = {
      name:      norm(src?.person?.fullName || src?.["p1:n"] || src?.fullName || "Perspective Overlay"),
      dateLbl:   norm(src?.dateLbl || src?.["p1:d"] || ""),

      summary:   norm(src?.summary   || src?.["p3:summary"] || ""),
      frequency: norm(src?.frequency || src?.["p4:freq"]    || ""),
      sequence:  norm(src?.sequence  || src?.["p5:seq"]     || ""),
      themepair: norm(src?.themepair || src?.["p6:theme"]   || ""),
      tips:      norm(src?.tips      || src?.["p7:tips"]    || "")
    };

    if (!P.dateLbl && src?.["p1:d"]) {
      const human = norm(src["p1:d"]);
      P.dateLbl = human.replace(/\s+/g, "_").toUpperCase();
    }

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc   = await PDFDocument.load(pdfBytes);
    const font     = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const fonts = { reg: font, bold: fontBold };

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);

    /* ───────── layout anchors ───────── */
    const L = {
      header: {
        x: 380,
        y: 51,
        w: 400,
        size: 13,
        align: "left",
        maxLines: 1
      },
      p1: {
        name: { x: 7,   y: 473, w: 500, size: 30, align: "center" },
        date: { x: 210, y: 600, w: 500, size: 25, align: "left" }
      },
      // p3: summary
      p3: {
        summary: { x: 25, y: 150, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      },
      // p4: frequency
      p4: {
        frequency: { x: 25, y: 150, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      },
      // p5: sequence
      p5: {
        sequence: { x: 25, y: 150, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      },
      // p6: themepair / theme lens
      p6: {
        themepair: { x: 25, y: 250, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      },
      // p7: tips / actions
      p7: {
        tips: { x: 25, y: 150, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      }
    };

    const overrideBox = (box, key) => {
      if (!box) return;
      if (q[`${key}x`]   != null) box.x        = N(q[`${key}x`],   box.x);
      if (q[`${key}y`]   != null) box.y        = N(q[`${key}y`],   box.y);
      if (q[`${key}w`]   != null) box.w        = N(q[`${key}w`],   box.w);
      if (q[`${key}s`]   != null) box.size     = N(q[`${key}s`],   box.size);
      if (q[`${key}max`] != null) box.maxLines = N(q[`${key}max`], box.maxLines);
      if (q[`${key}align`])       box.align    = String(q[`${key}align`]);
    };

    overrideBox(L.p3.summary,   "sum");
    overrideBox(L.p4.frequency, "freq");
    overrideBox(L.p5.sequence,  "seq");
    overrideBox(L.p6.themepair, "tp");
    overrideBox(L.p7.tips,      "tips");

    /* ───────── p1: full name & date ───────── */
    if (p1 && P.name)    drawTextBox(p1, font, P.name,    L.p1.name);
    if (p1 && P.dateLbl) drawTextBox(p1, font, P.dateLbl, L.p1.date);

    /* ───────── page headers ───────── */
    const putHeader = (page) => {
      if (!page || !P.name) return;
      drawTextBox(page, font, P.name, L.header, { maxLines: 1 });
    };
    [p2, p3, p4, p5, p6, p7, p8].forEach(putHeader);

    /* ───────── p3–p7: rich overlay content ───────── */
    if (p3 && P.summary)   drawOverlayBox(p3, fonts, P.summary,   L.p3.summary);
    if (p4 && P.frequency) drawOverlayBox(p4, fonts, P.frequency, L.p4.frequency);
    if (p5 && P.sequence)  drawOverlayBox(p5, fonts, P.sequence,  L.p5.sequence);
    if (p6 && P.themepair) drawOverlayBox(p6, fonts, P.themepair, L.p6.themepair);
    if (p7 && P.tips)      drawOverlayBox(p7, fonts, P.tips,      L.p7.tips);

    /* ───────── output ───────── */
    const bytes   = await pdfDoc.save();
    const outName = S(
      q.out || `CTRL_180_${P.name || "Perspective"}_${P.dateLbl || ""}.pdf`
    ).replace(/[^\w.-]+/g, "_");

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.end(Buffer.from(bytes));
  } catch (err) {
    console.error("fill-template-180 error", err);
    res
      .status(400)
      .json({ ok:false, error:`fill-template-180 error: ${err?.message || String(err)}` });
  }
}
