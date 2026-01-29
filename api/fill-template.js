/**
 * CTRL 180 Export Service · fill-template (V8)
 *
 * NOTE (V8 change only):
 * - Added filename/date helpers (clampStrForFilename / parseDateLabelToYYYYMMDD / makeOutputFilename)
 * - Expanded dateLbl sourcing (supports identity.dateLabel/dateLbl etc.)
 * - Added P.fullName (used only for output filename)
 * - Output filename now uses makeOutputFilename(...) unless overridden by ?out=
 *
 * Everything else is unchanged from V7.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────── runtime ───────── */
export const config = { runtime: "nodejs" };

/* ───────── tiny utils ───────── */
const S = (v) => (v == null ? "" : String(v));
const norm = (s) =>
  S(s)
    .replace(/\u00A0/g, " ")
    .replace(/\t/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \f\v]+/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);


/* ───────── filename helpers ───────── */
function clampStrForFilename(s) {
  return S(s)
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_\-]/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function parseDateLabelToYYYYMMDD(dateLbl) {
  const s = S(dateLbl).trim();
  const m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const monRaw = m[2].toLowerCase();
    const yyyy = m[3];
    const map = {
      jan: "01", january: "01",
      feb: "02", february: "02",
      mar: "03", march: "03",
      apr: "04", april: "04",
      may: "05",
      jun: "06", june: "06",
      jul: "07", july: "07",
      aug: "08", august: "08",
      sep: "09", sept: "09", september: "09",
      oct: "10", october: "10",
      nov: "11", november: "11",
      dec: "12", december: "12",
    };
    const mm = map[monRaw] || map[monRaw.slice(0, 3)];
    if (mm) return `${yyyy}-${mm}-${dd}`;
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  return clampStrForFilename(s || "date");
}

function makeOutputFilename(fullName, dateLbl) {
  const parts = S(fullName).trim().split(/\s+/).filter(Boolean);
  const first = parts[0] || "First";
  const last = parts.length > 1 ? parts[parts.length - 1] : "Surname";
  const datePart = parseDateLabelToYYYYMMDD(dateLbl);
  const fn = clampStrForFilename(first);
  const ln = clampStrForFilename(last);
  return `PoC_Profile_${fn}_${ln}_${datePart}.pdf`;
}

const bullets = (arr = []) =>
  arr.map((s) => norm(s || "")).filter(Boolean).map((s) => `- ${s}`).join("\n");

const pageOrNull = (pages, i) => (Array.isArray(pages) && pages[i] ? pages[i] : null);

function okNum(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function cap(n, lo, hi) {
  const x = okNum(n);
  if (x == null) return null;
  return Math.max(lo, Math.min(hi, x));
}

/* ───────── drawing primitives (unchanged) ───────── */
function drawTextBox(page, font, text, box, opts = {}) {
  const t = norm(text || "");
  if (!t) return;

  const x = okNum(box?.x) ?? 0;
  const y = okNum(box?.y) ?? 0;
  const w = okNum(box?.w) ?? 100;
  const h = okNum(box?.h) ?? 20;

  const size = okNum(opts.size ?? box?.size) ?? 12;
  const lineHeight = okNum(opts.lineHeight ?? box?.lineHeight) ?? Math.round(size * 1.25);

  const color = opts.color || rgb(0, 0, 0);
  const align = (opts.align ?? box?.align ?? "left").toLowerCase();
  const maxLines = okNum(opts.maxLines ?? box?.maxLines) ?? null;

  const words = t.split(/\s+/);
  const lines = [];
  let cur = "";

  for (const word of words) {
    const test = cur ? `${cur} ${word}` : word;
    const testW = font.widthOfTextAtSize(test, size);
    if (testW <= w || !cur) {
      cur = test;
    } else {
      lines.push(cur);
      cur = word;
    }
  }
  if (cur) lines.push(cur);

  const finalLines = maxLines ? lines.slice(0, maxLines) : lines;
  const totalH = finalLines.length * lineHeight;
  let yy = y + h - lineHeight;

  // If the content is taller than the box, start from top
  if (totalH > h) yy = y + h - lineHeight;

  for (const line of finalLines) {
    let xx = x;
    const lw = font.widthOfTextAtSize(line, size);
    if (align === "center") xx = x + (w - lw) / 2;
    if (align === "right") xx = x + (w - lw);

    page.drawText(line, { x: xx, y: yy, size, font, color });
    yy -= lineHeight;
    if (yy < y) break;
  }
}

function drawOverlayBox(page, fonts, text, box, opts = {}) {
  const t = norm(text || "");
  if (!t) return;

  const x = okNum(box?.x) ?? 0;
  const y = okNum(box?.y) ?? 0;
  const w = okNum(box?.w) ?? 100;
  const h = okNum(box?.h) ?? 20;

  const bg = box?.bg || opts.bg || rgb(1, 1, 1);
  const border = box?.border || opts.border || null;

  page.drawRectangle({ x, y, width: w, height: h, color: bg });
  if (border) {
    page.drawRectangle({
      x,
      y,
      width: w,
      height: h,
      borderColor: border,
      borderWidth: okNum(box?.borderWidth) ?? 1,
    });
  }

  drawTextBox(page, fonts.regular, t, box, opts);
}

/* ───────── main handler ───────── */
export default async function handler(req, res) {
  try {
    // Load template
    const tpl = S(req.query?.tpl || req.query?.template || "").trim();
    const baseDir = process.cwd();
    const defaultTplPath = path.join(baseDir, "public", "CTRL_PoC_180_Assessment_Report_template_fallback.pdf");
    const templatePath = tpl
      ? path.join(baseDir, "public", tpl)
      : defaultTplPath;

    const tplBytes = await readFile(templatePath);
    const pdfDoc = await PDFDocument.load(tplBytes);

    // Fonts
    const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const fonts = {
      regular: fontRegular,
      bold: fontBold,
    };

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);

    // Parse payload
    const rawBody = req.body;
    const src =
      (isObj(rawBody) ? rawBody : null) ||
      (isObj(req.query) ? req.query : null) ||
      {};

    const T = isObj(src?.text) ? src.text : null;

    const L = isObj(src?.layout) ? src.layout : null;

    const P = {
      fullName: norm(
        src?.person?.fullName ||
        src?.identity?.fullName ||
        src?.fullName ||
        ""
      ),
      name:    norm(src?.person?.fullName || src?.fullName || "Perspective Overlay"),
      dateLbl: norm(
        src?.dateLbl ||
        src?.identity?.dateLabel ||
        src?.identity?.dateLbl ||
        src?.person?.dateLabel ||
        src?.person?.dateLbl ||
        src?.date ||
        src?.Date ||
        ""
      ),

      // Legacy blocks (still supported)
      summary:   norm(T?.summary   || src?.summary   || ""),
      sequence:  norm(T?.sequence  || src?.sequence  || ""),
      themes:    norm(T?.themes    || src?.themes    || ""),

      // Prefer split fields if provided
      overview:  norm(T?.overview  || src?.overview  || ""),
      deepdive:  norm(T?.deepdive  || src?.deepdive  || ""),

      // Work-with blocks
      workWith_Colleagues: norm(T?.workWith_Colleagues || src?.workWith_Colleagues || ""),
      workWith_Leaders:    norm(T?.workWith_Leaders    || src?.workWith_Leaders    || ""),

      // Page 7
      tip1: norm(T?.tip1 || src?.tip1 || ""),
      tip2: norm(T?.tip2 || src?.tip2 || ""),
      act1: norm(T?.act1 || src?.act1 || ""),
      act2: norm(T?.act2 || src?.act2 || ""),
      act3: norm(T?.act3 || src?.act3 || ""),
    };

    // Default layout (unchanged)
    const DEFAULT_LAYOUT = {
      p1: {
        name: { x: 60, y: 735, w: 520, h: 40, size: 26, align: "left", maxLines: 1 },
        date: { x: 60, y: 705, w: 520, h: 20, size: 14, align: "left", maxLines: 1 },
      },
      p2: {
        summary: { x: 60, y: 530, w: 520, h: 240, size: 12, lineHeight: 16, align: "left", maxLines: 14 },
      },
      p3: {
        sequence: { x: 60, y: 420, w: 520, h: 350, size: 12, lineHeight: 16, align: "left", maxLines: 20 },
      },
      p4: {
        themes: { x: 60, y: 420, w: 520, h: 350, size: 12, lineHeight: 16, align: "left", maxLines: 20 },
      },
      p5: {
        overview: { x: 60, y: 420, w: 520, h: 350, size: 12, lineHeight: 16, align: "left", maxLines: 20 },
      },
      p6: {
        deepdive: { x: 60, y: 420, w: 520, h: 350, size: 12, lineHeight: 16, align: "left", maxLines: 20 },
      },
      p7: {
        workWith: {
          colleagues: { x: 60, y: 520, w: 520, h: 180, size: 12, lineHeight: 16, align: "left", maxLines: 10 },
          leaders:    { x: 60, y: 320, w: 520, h: 180, size: 12, lineHeight: 16, align: "left", maxLines: 10 },
        },
        tips: {
          tip1: { x: 60, y: 190, w: 520, h: 40, size: 12, lineHeight: 16, align: "left", maxLines: 2 },
          tip2: { x: 60, y: 140, w: 520, h: 40, size: 12, lineHeight: 16, align: "left", maxLines: 2 },
        },
        actions: {
          act1: { x: 60, y: 90, w: 520, h: 30, size: 12, lineHeight: 16, align: "left", maxLines: 1 },
          act2: { x: 60, y: 60, w: 520, h: 30, size: 12, lineHeight: 16, align: "left", maxLines: 1 },
          act3: { x: 60, y: 30, w: 520, h: 30, size: 12, lineHeight: 16, align: "left", maxLines: 1 },
        },
      },
    };

    const layout = isObj(L) ? L : {};
    const Lx = {
      ...DEFAULT_LAYOUT,
      ...layout,
      p1: { ...DEFAULT_LAYOUT.p1, ...(layout?.p1 || {}) },
      p2: { ...DEFAULT_LAYOUT.p2, ...(layout?.p2 || {}) },
      p3: { ...DEFAULT_LAYOUT.p3, ...(layout?.p3 || {}) },
      p4: { ...DEFAULT_LAYOUT.p4, ...(layout?.p4 || {}) },
      p5: { ...DEFAULT_LAYOUT.p5, ...(layout?.p5 || {}) },
      p6: { ...DEFAULT_LAYOUT.p6, ...(layout?.p6 || {}) },
      p7: {
        ...DEFAULT_LAYOUT.p7,
        ...(layout?.p7 || {}),
        workWith: {
          ...DEFAULT_LAYOUT.p7.workWith,
          ...(layout?.p7?.workWith || {}),
          colleagues: { ...DEFAULT_LAYOUT.p7.workWith.colleagues, ...(layout?.p7?.workWith?.colleagues || {}) },
          leaders:    { ...DEFAULT_LAYOUT.p7.workWith.leaders,    ...(layout?.p7?.workWith?.leaders || {}) },
        },
        tips: {
          ...DEFAULT_LAYOUT.p7.tips,
          ...(layout?.p7?.tips || {}),
          tip1: { ...DEFAULT_LAYOUT.p7.tips.tip1, ...(layout?.p7?.tips?.tip1 || {}) },
          tip2: { ...DEFAULT_LAYOUT.p7.tips.tip2, ...(layout?.p7?.tips?.tip2 || {}) },
        },
        actions: {
          ...DEFAULT_LAYOUT.p7.actions,
          ...(layout?.p7?.actions || {}),
          act1: { ...DEFAULT_LAYOUT.p7.actions.act1, ...(layout?.p7?.actions?.act1 || {}) },
          act2: { ...DEFAULT_LAYOUT.p7.actions.act2, ...(layout?.p7?.actions?.act2 || {}) },
          act3: { ...DEFAULT_LAYOUT.p7.actions.act3, ...(layout?.p7?.actions?.act3 || {}) },
        },
      },
    };

    // Draw page 1 (unchanged)
    if (p1) {
      if (P.name) drawTextBox(p1, fontBold, P.name, Lx.p1.name, { maxLines: Lx.p1.name.maxLines ?? 1 });
      if (P.dateLbl) drawTextBox(p1, fontRegular, P.dateLbl, Lx.p1.date, { maxLines: Lx.p1.date.maxLines ?? 1 });
    }

    // Page 2 summary
    if (p2 && P.summary) {
      drawTextBox(p2, fontRegular, P.summary, Lx.p2.summary);
    }

    // Page 3 sequence
    if (p3 && P.sequence) {
      drawTextBox(p3, fontRegular, P.sequence, Lx.p3.sequence);
    }

    // Page 4 themes
    if (p4 && P.themes) {
      drawTextBox(p4, fontRegular, P.themes, Lx.p4.themes);
    }

    // Page 5 overview
    if (p5 && P.overview) {
      drawTextBox(p5, fontRegular, P.overview, Lx.p5.overview);
    }

    // Page 6 deepdive
    if (p6 && P.deepdive) {
      drawTextBox(p6, fontRegular, P.deepdive, Lx.p6.deepdive);
    }

    // Page 7 work-with + tips/actions
    if (p7) {
      const wwC = P.workWith_Colleagues;
      const wwL = P.workWith_Leaders;

      if (wwC) drawOverlayBox(p7, fonts, wwC, Lx.p7.workWith.colleagues);
      if (wwL) drawOverlayBox(p7, fonts, wwL, Lx.p7.workWith.leaders);

      const tip1 = P.tip1;
      const tip2 = P.tip2;

      if (tip1) drawOverlayBox(p7, fonts, tip1, Lx.p7.tips.tip1);
      if (tip2) drawOverlayBox(p7, fonts, tip2, Lx.p7.tips.tip2);

      const act1 = P.act1;
      const act2 = P.act2;
      const act3 = P.act3;

      if (act1) drawOverlayBox(p7, fonts, act1, Lx.p7Actions?.act1 || Lx.p7.actions.act1);
      if (act2) drawOverlayBox(p7, fonts, act2, Lx.p7Actions?.act2 || Lx.p7.actions.act2);
      if (act3) drawOverlayBox(p7, fonts, act3, Lx.p7Actions?.act3 || Lx.p7.actions.act3);
    }

    const bytes = await pdfDoc.save();

    const outName = S(
      req.query?.out ||
      makeOutputFilename(P.fullName || P.name || "Perspective", P.dateLbl || "")
    ).replace(/[^\w.-]+/g, "_");

    // Response headers
    res.setHeader("X-CTRL-TPL", tpl);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.status(200).send(Buffer.from(bytes));
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: S(err?.message || err),
    });
  }
}
