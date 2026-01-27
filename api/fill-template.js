/**
 * CTRL 180 Export Service · fill-template (V2 — COORDS UPDATE ONLY)
 * Path: /pages/api/fill-template.js  (ctrl-180-service)
 *
 * Templates live in:
 *   ctrl-180-service/public/CTRL_PoC_180_Assessment_Report_template_XX.pdf
 *
 * Required:
 *   /api/fill-template?tpl=<filename.pdf>&data=<base64json>
 *
 * Only fallback allowed (when tpl missing/blank):
 *   CTRL_PoC_180_Assessment_Report_template_fallback.pdf
 *
 * V2 adds:
 *  - payload.text / payload.textV2 inspection in debug mode
 *  - ?debug=1 returns JSON diagnostics (no PDF)
 *  - headers for chart/template/payload tracing
 *
 * THIS EDIT:
 *  - ONLY layout coordinates changed to match "fill-template user v12.3"
 *  - plus page 6 uses provided p6WorkWith coords
 */

export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── utils ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);

const norm = (v, fb = "") =>
  String(v ?? fb)
    .normalize("NFKC")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2010-\u2014]/g, "-")
    .replace(/\u2026/g, "...")
    .replace(/\u00A0/g, " ")
    .replace(/[•·]/g, "-")
    .replace(/\u2194/g, "<->").replace(/\u2192/g, "->").replace(/\u2190/g, "<-")
    .replace(/\u2191/g, "^").replace(/\u2193/g, "v").replace(/[\u2196-\u2199]/g, "->")
    .replace(/\u21A9/g, "<-").replace(/\u21AA/g, "->")
    .replace(/\u00D7/g, "x")
    .replace(/[\u200B-\u200D\u2060]/g, "")
    .replace(/[\uD800-\uDFFF]/g, "")
    .replace(/[\uE000-\uF8FF]/g, "")
    .replace(/\t/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \f\v]+/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);

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

async function readPayload(req) {
  const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
  if (q.data) return parseDataParam(q.data);
  if (req.method === "POST" && !q.data) {
    try {
      return typeof req.json === "function" ? await req.json() : (req.body || {});
    } catch {}
  }
  return {};
}

/* ───────────── drawing helpers ───────────── */
function drawTextBox(page, font, text, spec = {}, opts = {}) {
  if (!page) return;
  const {
    x = 40, y = 40, w = 540,
    size = 12, lineGap = 3,
    color = rgb(0, 0, 0),
    align = "left", h,
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
      else { wrapped.push(cur); cur = words[i]; }
    }
    wrapped.push(cur);
  };
  for (const ln of lines) wrapLine(ln);

  const out = wrapped.slice(0, maxLines);
  const pageH = page.getHeight();
  let yCursor = pageH - y;

  for (const ln of out) {
    let xDraw = x;
    const wLn = widthOf(ln);
    if (align === "center") xDraw = x + (w - wLn) / 2;
    else if (align === "right") xDraw = x + (w - wLn);

    page.drawText(ln, { x: xDraw, y: yCursor - size, size: Math.max(1, size), font, color });
    yCursor -= lineHeight;
  }
}

function drawOverlayBox(page, fonts, text, spec = {}) {
  if (!page) return;
  const fontReg = fonts.reg;
  const fontBold = fonts.bold;

  const {
    x = 40, y = 40, w = 540,
    size = 15, lineGap = 6,
    color = rgb(0, 0, 0),
    align = "left",
    maxLines = 120
  } = spec;

  const hard = norm(text || "");
  if (!hard) return;

  const pageH = page.getHeight();
  let yCursor = pageH - y;
  let usedLines = 0;

  const pushLine = (ln, { font, fSize, indent = 0 }) => {
    if (usedLines >= maxLines) return;
    const widthOf = (s) => font.widthOfTextAtSize(s, fSize);

    const words = ln.split(/\s+/);
    let current = "";
    const segments = [];

    for (let i = 0; i < words.length; i++) {
      const next = current ? current + " " + words[i] : words[i];
      if (widthOf(next) <= (w - indent) || !current) current = next;
      else { segments.push(current); current = words[i]; }
    }
    if (current) segments.push(current);

    for (let s of segments) {
      if (usedLines >= maxLines) break;
      let xDraw = x + indent;
      const wLn = widthOf(s);
      if (align === "center") xDraw = x + (w - wLn) / 2;
      else if (align === "right") xDraw = x + (w - wLn);

      page.drawText(s, { x: xDraw, y: yCursor - fSize, size: fSize, font, color });
      yCursor -= (fSize + lineGap);
      usedLines++;
    }
  };

  const rawLines = hard.split(/\n+/);
  for (let raw of rawLines) {
    const line = raw.trim();
    if (!line) { yCursor -= (size + lineGap); continue; }

    const isHeading = /^what\b.*:\s*$/i.test(line);
    const isBullet  = /^-\s+/.test(line);

    if (isHeading) {
      yCursor -= (size + lineGap) * 2;
      pushLine(line.replace(/:\s*$/, ":"), { font: fontBold, fSize: size + 1, indent: 0 });
      yCursor -= (size + lineGap) * 0.5;
      continue;
    }

    if (isBullet) {
      pushLine(line.replace(/^-\s+/, "• "), { font: fontReg, fSize: size, indent: 10 });
      continue;
    }

    pushLine(line, { font: fontReg, fSize: size, indent: 0 });
  }
}

/* ───────────── template loader (NO silent fallback) ───────────── */
async function loadTemplateBytesLocal(filename) {
  const fname = String(filename || "").trim();
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir  = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(process.cwd(), "public", fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try { return await fs.readFile(pth); }
    catch (err) { lastErr = err; }
  }
  throw new Error(`Template not found in /public: ${fname} (${lastErr?.message || "no detail"})`);
}

const pageOrNull = (pages, idx0) => (pages[idx0] ?? null);

/* ───────────── chart fetch/embed ───────────── */
async function fetchBytes(url, timeoutMs = 9000) {
  if (!url) return { ok:false, reason:"no_url", url:"", contentType:"", bytes:null };
  const u = String(url).trim();
  if (!/^https?:\/\//i.test(u)) return { ok:false, reason:"not_http", url:u, contentType:"", bytes:null };

  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);

  try {
    const r = await fetch(u, { signal: ctrl.signal });
    if (!r.ok) return { ok:false, reason:`http_${r.status}`, url:u, contentType:"", bytes:null };
    const ct = (r.headers.get("content-type") || "").toLowerCase();
    const ab = await r.arrayBuffer();
    return { ok:true, reason:"ok", url:u, contentType:ct, bytes:new Uint8Array(ab) };
  } catch (e) {
    return { ok:false, reason:String(e?.name || e?.message || "fetch_error"), url:u, contentType:"", bytes:null };
  } finally {
    clearTimeout(t);
  }
}

function looksPng(u, ct = "") {
  const s = String(u || "").toLowerCase();
  return ct.includes("png") || s.endsWith(".png") || s.includes(".png?");
}
function looksJpg(u, ct = "") {
  const s = String(u || "").toLowerCase();
  return ct.includes("jpeg") || ct.includes("jpg") || s.endsWith(".jpg") || s.endsWith(".jpeg") || s.includes(".jpg?") || s.includes(".jpeg?");
}

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  let tpl = "";
  let usingFallback = false;
  let chartFetch = { ok:false, reason:"not_attempted", url:"", contentType:"", bytes:null };
  let payloadSize = 0;

  try {
    const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
    const debugMode = String(q.debug || "").trim() === "1";

    const FALLBACK_TPL = "CTRL_PoC_180_Assessment_Report_template_fallback.pdf";

    tpl = S(q.tpl || "").trim();
    tpl = tpl.replace(/[^A-Za-z0-9._-]/g, "");

    usingFallback = !tpl;
    if (usingFallback) tpl = FALLBACK_TPL;

    const pdfBytes = await loadTemplateBytesLocal(tpl);

    const src = await readPayload(req);
    payloadSize = Buffer.byteLength(JSON.stringify(src || {}), "utf8");

    const P = {
      name:      norm(src?.person?.fullName || src?.fullName || "Perspective Overlay"),
      dateLbl:   norm(src?.dateLbl || ""),

      summary:   norm(src?.summary   || ""),
      frequency: norm(src?.frequency || ""),
      sequence:  norm(src?.sequence  || ""),
      themepair: norm(src?.themepair || ""),
      tips:      norm(src?.tips      || ""),

      // Optional (does not change behaviour unless payload provides it)
      workWith: isObj(src?.workWith) ? src.workWith : null
    };

    const TextV2 = isObj(src?.textV2) ? src.textV2 : null;
    const Text   = isObj(src?.text)   ? src.text   : null;

    const chartUrl = norm(
      src?.chartUrl ||
      src?.spiderChartUrl ||
      src?.chart?.spiderUrl ||
      ""
    );

    if (debugMode) {
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.setHeader("X-CTRL-TPL", tpl);
      res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
      res.setHeader("X-CTRL-DEBUG", "1");
      res.status(200).end(JSON.stringify({
        ok: true,
        debug: true,
        tpl,
        usingFallback,
        received: {
          person: src?.person || null,
          dateLbl: src?.dateLbl || null,
          chartUrl: chartUrl || null,
          hasText: !!Text,
          hasTextV2: !!TextV2,
          fields: {
            summary_len: P.summary.length,
            frequency_len: P.frequency.length,
            sequence_len: P.sequence.length,
            themepair_len: P.themepair.length,
            tips_len: P.tips.length
          },
          hasWorkWith: !!P.workWith,
          workWithKeys: P.workWith ? Object.keys(P.workWith).slice(0, 60) : []
        }
      }, null, 2));
      return;
    }

    const pdfDoc   = await PDFDocument.load(pdfBytes);
    const font     = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const fonts    = { reg: font, bold: fontBold };

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);

    // ─────────────────────────────────────────────────────────────
    // Layout — COORDINATES UPDATED per request (from user v12.3)
    // ─────────────────────────────────────────────────────────────
    const L = {
      header: { x: 380, y: 51, w: 400, size: 13, align: "left", maxLines: 1 },

      // a) Page 1 (180) = Page 1 on attached (name/date)
      p1: {
        name: { x: 60, y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
        date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left",   maxLines: 1 }
      },

      // b) Page 2 chart stays the same
      p2: {
        chart: { x: 250, y: 160, w: 320, h: 320 }
      },

      // c) Page 3 (180) = Page 4 on attached (exec summary area)
      // Using the primary big text region (exec2)
      p3: {
        summary: { x: 25, y: 390, w: 550, size: 16, lineGap: 6, align: "left", maxLines: 22 }
      },

      // d) Page 4 (180) = Page 5 on attached (overview main text region ov2)
      p4: {
        frequency: { x: 25, y: 590, w: 550, size: 16, lineGap: 6, align: "left", maxLines: 23 }
      },

      // e) Page 5 (180) = Page 6 on attached (deepdive main region dd2)
      p5: {
        sequence: { x: 25, y: 270, w: 550, size: 16, lineGap: 6, align: "left", maxLines: 17 }
      },

      // f) Page 6 (180) = provided p6WorkWith coords (two-column)
      // NOTE: only draws if payload.workWith.* exists (does not affect legacy behaviour)
      p6: {
        p6WorkWith: {
          collabC_text: { x: 30,  y: 300, w: 270, h: 110, size: 16, align: "left", maxLines: 7 },
          collabC_q:    { x: 30,  y: 410, w: 270, h: 310, size: 16, align: "left", maxLines: 24 },
          collabT_text: { x: 320, y: 300, w: 260, h: 110, size: 16, align: "left", maxLines: 7 },
          collabT_q:    { x: 320, y: 410, w: 260, h: 310, size: 16, align: "left", maxLines: 24 }
        },
        // legacy single-block fallback (kept, but not used if workWith is present)
        themepair: { x: 25, y: 280, w: 550, size: 15, lineGap: 6, align: "left", maxLines: 110 }
      },

      // g) Page 7 (180) = Page 8 on attached (actions area act1)
      // Using the primary actions text region
      p7: {
        tips: { x: 50, y: 380, w: 440, size: 17, lineGap: 6, align: "left", maxLines: 18 }
      }
      // h) Page 8 stays the same (no changes required; page exists for template)
    };

    // Optional query overrides for chart placement (kept unchanged)
    if (q.cx != null) L.p2.chart.x = N(q.cx, L.p2.chart.x);
    if (q.cy != null) L.p2.chart.y = N(q.cy, L.p2.chart.y);
    if (q.cw != null) L.p2.chart.w = N(q.cw, L.p2.chart.w);
    if (q.ch != null) L.p2.chart.h = N(q.ch, L.p2.chart.h);

    // p1: name & date
    if (p1 && P.name)    drawTextBox(p1, fontBold, P.name,   L.p1.name, { maxLines: 1 });
    if (p1 && P.dateLbl) drawTextBox(p1, font,     P.dateLbl, L.p1.date, { maxLines: 1 });

    // headers (kept unchanged)
    const putHeader = (page) => {
      if (!page || !P.name) return;
      drawTextBox(page, font, P.name, L.header, { maxLines: 1 });
    };
    [p2, p3, p4, p5, p6, p7, p8].forEach(putHeader);

    // p2: chart image (optional) (kept unchanged)
    if (p2 && chartUrl) {
      chartFetch = await fetchBytes(chartUrl, 9000);
      if (chartFetch.ok && chartFetch.bytes && chartFetch.bytes.length) {
        let img = null;
        if (looksPng(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedPng(chartFetch.bytes);
        else if (looksJpg(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedJpg(chartFetch.bytes);

        if (img) {
          const { x, y, w, h } = L.p2.chart;
          const pageH = p2.getHeight();
          const yBottom = pageH - y - h;
          p2.drawImage(img, { x, y: yBottom, width: w, height: h });
        }
      }
    }

    // p3–p5, p7 overlay blocks (kept same behaviour; coords updated)
    if (p3 && P.summary)   drawOverlayBox(p3, fonts, P.summary,   L.p3.summary);
    if (p4 && P.frequency) drawOverlayBox(p4, fonts, P.frequency, L.p4.frequency);
    if (p5 && P.sequence)  drawOverlayBox(p5, fonts, P.sequence,  L.p5.sequence);

    // p6: workWith if present, else legacy themepair (behaviour remains backward compatible)
    if (p6) {
      const WW = P.workWith;

      if (WW && (WW.collabC_text || WW.collabC_q || WW.collabT_text || WW.collabT_q)) {
        const B = L.p6.p6WorkWith;

        if (WW.collabC_text) drawOverlayBox(p6, fonts, norm(WW.collabC_text), B.collabC_text);
        if (WW.collabC_q)    drawOverlayBox(p6, fonts, norm(WW.collabC_q),    B.collabC_q);

        if (WW.collabT_text) drawOverlayBox(p6, fonts, norm(WW.collabT_text), B.collabT_text);
        if (WW.collabT_q)    drawOverlayBox(p6, fonts, norm(WW.collabT_q),    B.collabT_q);
      } else {
        if (P.themepair) drawOverlayBox(p6, fonts, P.themepair, L.p6.themepair);
      }
    }

    if (p7 && P.tips)      drawOverlayBox(p7, fonts, P.tips,      L.p7.tips);

    const bytes = await pdfDoc.save();

    const outName = S(
      q.out || `CTRL_180_${P.name || "Perspective"}_${P.dateLbl || ""}.pdf`
    ).replace(/[^\w.-]+/g, "_");

    // Response headers (kept)
    res.setHeader("X-CTRL-TPL", tpl);
    res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
    res.setHeader("X-CTRL-PAYLOAD-SIZE", String(payloadSize || 0));
    res.setHeader("X-CTRL-CHART", chartUrl ? "1" : "0");
    res.setHeader("X-CTRL-CHART-FETCH", chartFetch?.ok ? "ok" : (chartFetch?.reason || "no"));
    res.setHeader("X-CTRL-DEBUG", "0");

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.end(Buffer.from(bytes));
  } catch (err) {
    console.error("fill-template-180 error", err);

    try {
      res.setHeader("X-CTRL-TPL", tpl || "");
      res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
      res.setHeader("X-CTRL-PAYLOAD-SIZE", String(payloadSize || 0));
      res.setHeader("X-CTRL-CHART-FETCH", chartFetch?.reason || "error");
      res.setHeader("X-CTRL-DEBUG", "0");
    } catch {}

    res.status(400).json({
      ok: false,
      error: `fill-template-180 error: ${err?.message || String(err)}`
    });
  }
}
