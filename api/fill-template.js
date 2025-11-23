/**
 * CTRL 180 Export Service · fill-template (Observer overlay)
 * Path: /pages/api/fill-template.js  (ctrl-180-service)
 *
 * Template:
 *   CTRL_Perspective_Assessment_Profile_template_slim_180.pdf
 *
 * Layout:
 *   p1: name/date (cover)
 *   p2: name header only
 *   p3: summary          (overlay summary)
 *   p4: frequency        (overlay frequency story)
 *   p5: sequence         (overlay sequence story)
 *   p6: themepair        (overlay theme lens)
 *   p7: tips/actions     (overlay tips / actions)
 *   p8: name header only
 *
 * Header: full name appears on pages 2–8.
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

/* TL → simple textbox (does internal TL->BL conversion) */
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
    h,               // optional height for auto maxLines
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

    // default to 180 template; allow ?tpl= override if needed
    const defaultTpl = "CTRL_Perspective_Assessment_Profile_template_slim_180.pdf";
    const tpl        = S(q.tpl || defaultTpl).replace(/[^A-Za-z0-9._-]/g, "");
    const src        = await readPayload(req);

    // Expected payload from Botpress/Make:
    // {
    //   person: { fullName },
    //   dateLbl,          // 23_NOV_2025
    //   "p1:n", "p1:d",   // optional, for legacy
    //   "p3:summary",
    //   "p4:freq",
    //   "p5:seq",
    //   "p6:theme",
    //   "p7:tips"
    // }
    const P = {
      name:      norm(src?.person?.fullName || src?.["p1:n"] || src?.fullName || "Perspective Overlay"),
      dateLbl:   norm(src?.dateLbl || src?.["p1:d"] || ""),

      summary:   norm(src?.summary   || src?.["p3:summary"] || ""),
      frequency: norm(src?.frequency || src?.["p4:freq"]    || ""),
      sequence:  norm(src?.sequence  || src?.["p5:seq"]     || ""),
      themepair: norm(src?.themepair || src?.["p6:theme"]   || ""),
      tips:      norm(src?.tips      || src?.["p7:tips"]    || "")
    };

    // Fallback: if dateLbl missing but p1:d has a human string, convert to label-like:
    if (!P.dateLbl && src?.["p1:d"]) {
      const human = norm(src["p1:d"]);
      // crude: replace spaces with underscores, keep caps
      P.dateLbl = human.replace(/\s+/g, "_").toUpperCase(); // "23_NOV_2025"
    }

    // load template
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
    const p8 = pageOrNull(pages, 7);

    /* ───────────── layout anchors (defaults) ───────────── */
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
        summary: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 100 }
      },
      // p4: frequency
      p4: {
        frequency: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 100 }
      },
      // p5: sequence
      p5: {
        sequence: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 100 }
      },
      // p6: themepair / theme lens
      p6: {
        themepair: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 100 }
      },
      // p7: tips / actions
      p7: {
        tips: { x: 25, y: 150, w: 550, size: 13, align: "left", maxLines: 100 }
      }
    };

    /* ───────────── optional URL overrides (if you ever tweak) ───────────── */
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

    /* ───────────── p1: full name & date ───────────── */
    if (p1 && P.name)    drawTextBox(p1, font, P.name,    L.p1.name);
    if (p1 && P.dateLbl) drawTextBox(p1, font, P.dateLbl, L.p1.date);

    /* ───────────── page headers (p2..p8) ───────────── */
    const putHeader = (page) => {
      if (!page || !P.name) return;
      drawTextBox(page, font, P.name, L.header, { maxLines: 1 });
    };
    [p2, p3, p4, p5, p6, p7, p8].forEach(putHeader);

    /* ───────────── p3: summary ───────────── */
    if (p3 && P.summary) {
      drawTextBox(p3, font, P.summary, L.p3.summary);
    }

    /* ───────────── p4: frequency ───────────── */
    if (p4 && P.frequency) {
      drawTextBox(p4, font, P.frequency, L.p4.frequency);
    }

    /* ───────────── p5: sequence ───────────── */
    if (p5 && P.sequence) {
      drawTextBox(p5, font, P.sequence, L.p5.sequence);
    }

    /* ───────────── p6: themepair ───────────── */
    if (p6 && P.themepair) {
      drawTextBox(p6, font, P.themepair, L.p6.themepair);
    }

    /* ───────────── p7: tips/actions ───────────── */
    if (p7 && P.tips) {
      drawTextBox(p7, font, P.tips, L.p7.tips);
    }

    /* ───────── output ───────── */
    const bytes = await pdfDoc.save();
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
