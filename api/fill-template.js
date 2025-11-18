// /api/fill-template.js  — CTRL 180 overlay PDF service
// Template: /public/CTRL_Perspective_Assessment_Profile_template_slim_180.pdf
//
// Layout (0-indexed pages):
// p1: cover (name + date)
// p2: header only / intro (optional)
// p3: summary (dominant + similarities/differences)
// p4: frequency story
// p5: sequence story
// p6: themepair lens
// p7: tips & actions

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
    .replace(/\u200B-\u200D\u2060/g, "")
    .replace(/[\uD800-\uDFFF]/g, "")
    .replace(/[\uE000-\uF8FF]/g, "")
    .replace(/\t/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \f\v]+/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

/* data param as in your other services (?data=base64json) */
function parseDataParam(b64ish) {
  if (!b64ish) return {};
  let s = String(b64ish);
  try { s = decodeURIComponent(s); } catch {}
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  try { return JSON.parse(Buffer.from(s, "base64").toString("utf8")); }
  catch { return {}; }
}

async function readPayload(req) {
  const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
  if (q.data) return parseDataParam(q.data);
  if (req.method === "POST" && !q.data) {
    try {
      return typeof req.json === "function" ? await req.json() : (req.body || {});
    } catch { /* ignore */ }
  }
  return {};
}

/* Simple TL text box */
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
    (h ? Math.max(1, Math.floor(h / lineHeight)) : 20);

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

/* robust /public loader */
async function loadTemplateBytesLocal(filename) {
  const fname = String(filename || "").trim();
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir  = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
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
    const q   = req.method === "POST" ? (req.body || {}) : (req.query || {});
    const defaultTpl = "CTRL_Perspective_Assessment_Profile_template_slim_180.pdf";
    const tpl  = S(q.tpl || defaultTpl).replace(/[^A-Za-z0-9._-]/g, "");
    const src  = await readPayload(req);

    // expected payload from Make/Botpress (you can tweak names later)
    const P = {
      name:       norm(src.fullName || src.person?.fullName || "Perspective Overlay"),
      dateLbl:    norm(src.dateLbl || src.dateLabel || ""),
      summary:    norm(src.summary || src.summary180 || src.overview || ""),
      frequency:  norm(src.frequency || src.frequency180 || ""),
      sequence:   norm(src.sequence || src.sequence180 || ""),
      themepair:  norm(src.themepair || src.themeLens || ""),
      tips:       norm(src.tipsActions || src.tips || src.actions || ""),
    };

    const pdfBytes = await loadTemplateBytesLocal(tpl);
    const pdfDoc   = await PDFDocument.load(pdfBytes);
    const font     = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);

    /* ───────────── layout anchors ───────────── */
    const L = {
      header: { x: 380, y: 51, w: 400, size: 13, align: "left", maxLines: 1 },

      p1: {
        name: { x: 7,   y: 473, w: 500, size: 30, align: "center" },
        date: { x: 210, y: 600, w: 500, size: 25, align: "left" },
      },

      // p3 summary
      p3: {
        summary: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 40 },
      },

      // p4 frequency
      p4: {
        freq: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 40 },
      },

      // p5 sequence
      p5: {
        seq: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 40 },
      },

      // p6 themepair
      p6: {
        theme: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 40 },
      },

      // p7 tips + actions
      p7: {
        tips: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 40 },
      },
    };

    /* optional overrides via query: sum*, freq*, seq*, tp*, ta* */
    const overrideBox = (box, key) => {
      if (!box) return;
      if (q[`${key}x`]   != null) box.x        = N(q[`${key}x`],   box.x);
      if (q[`${key}y`]   != null) box.y        = N(q[`${key}y`],   box.y);
      if (q[`${key}w`]   != null) box.w        = N(q[`${key}w`],   box.w);
      if (q[`${key}s`]   != null) box.size     = N(q[`${key}s`],   box.size);
      if (q[`${key}max`] != null) box.maxLines = N(q[`${key}max`], box.maxLines);
      if (q[`${key}align`])       box.align    = String(q[`${key}align`]);
    };

    overrideBox(L.p3.summary, "sum");
    overrideBox(L.p4.freq,    "freq");
    overrideBox(L.p5.seq,     "seq");
    overrideBox(L.p6.theme,   "tp");
    overrideBox(L.p7.tips,    "ta");

    /* ───────────── p1: name + date ───────────── */
    if (p1 && P.name)    drawTextBox(p1, font, P.name,    L.p1.name);
    if (p1 && P.dateLbl) drawTextBox(p1, font, P.dateLbl, L.p1.date);

    /* headers on p2–p7 */
    const putHeader = (page) => {
      if (!page || !P.name) return;
      drawTextBox(page, font, P.name, L.header, { maxLines: 1 });
    };
    [p2, p3, p4, p5, p6, p7].forEach(putHeader);

    /* content pages */
    if (p3 && P.summary)   drawTextBox(p3, font, P.summary,   L.p3.summary);
    if (p4 && P.frequency) drawTextBox(p4, font, P.frequency, L.p4.freq);
    if (p5 && P.sequence)  drawTextBox(p5, font, P.sequence,  L.p5.seq);
    if (p6 && P.themepair) drawTextBox(p6, font, P.themepair, L.p6.theme);
    if (p7 && P.tips)      drawTextBox(p7, font, P.tips,      L.p7.tips);

    /* ───────── output ───────── */
    const bytes   = await pdfDoc.save();
    const outName = S(
      q.out || `CTRL_180_${P.name || "Profile"}_${P.dateLbl || ""}.pdf`
    ).replace(/[^\w.-]+/g, "_");

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.end(Buffer.from(bytes));
  } catch (err) {
    console.error("fill-template 180 error", err);
    res.statusCode = 400;
    res.setHeader("Content-Type", "application/json");
    res.end(
      JSON.stringify({
        ok: false,
        error: `fill-template 180 error: ${err?.message || String(err)}`
      })
    );
  }
}
