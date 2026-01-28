/**
 * CTRL 180 Export Service · fill-template (V5 — ctrl_overview + ctrl_overviewQ as 2 clean blobs)
 * Path: /pages/api/fill-template.js  (ctrl-180-service)
 *
 * Required:
 *   /api/fill-template?tpl=<filename.pdf>&data=<base64json>
 *
 * Only fallback allowed (when tpl missing/blank):
 *   CTRL_PoC_180_Assessment_Report_template_fallback.pdf
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

/* ───────── DEFAULT LAYOUT (LOCKED IN) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 60, y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p5: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p7: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p8: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    p3Text: {
      summary: { x: 25, y: 380, w: 550, h: 250, size: 16, align: "left", maxLines: 13 },
    },

    // ✅ V5: replace old single-field flow with two explicit blobs
    p4Text: {
      ctrl_overview:  { x: 25, y: 80,  w: 160, h: 240, size: 15, align: "left", maxLines: 30 },
      ctrl_overviewQ: { x: 25, y: 390, w: 550, h: 420, size: 16, align: "left", maxLines: 23 },
      chart: { x: 250, y: 160, w: 320, h: 320 }, // chart target (page 4)
    },

    p5Text: {
      sequence: { x: 25, y: 140, w: 550, h: 240, size: 16, align: "left", maxLines: 13 },
      theme:    { x: 25, y: 540, w: 550, h: 160, size: 16, align: "left", maxLines: 9 },
    },

    p6WorkWith: {
      collabCol: { x: 30,  y: 300, w: 270, h: 420, size: 15, align: "left", maxLines: 14 },
      collabLe:  { x: 320, y: 300, w: 260, h: 420, size: 15, align: "left", maxLines: 14 },
    },

    p7Actions: {
      act1: { x: 50,  y: 380, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act2: { x: 100, y: 530, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act3: { x: 50,  y: 670, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
    },
  },
};

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
  const { x = 40, y = 40, w = 540, size = 12, lineGap = 3, color = rgb(0, 0, 0), align = "left", h } = spec;
  const lineHeight = Math.max(1, size) + lineGap;
  const maxLines = opts.maxLines ?? spec.maxLines ?? (h ? Math.max(1, Math.floor(h / lineHeight)) : 40);

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

  const { x = 40, y = 40, w = 540, size = 15, lineGap = 6, color = rgb(0, 0, 0), align = "left", maxLines = 120 } = spec;
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

/* ───────────── tiny helpers for page 6/7 splitting ───────────── */
function combineWorkWithSide(WW, side) {
  if (!isObj(WW)) return "";

  // New shape
  if (side === "C" && WW.collabCol) return norm(WW.collabCol);
  if (side === "T" && WW.collabLe)  return norm(WW.collabLe);

  // Legacy shape
  if (side === "C") {
    const a = norm(WW.collabC_text || "");
    const b = norm(WW.collabC_q || "");
    return norm([a, b].filter(Boolean).join("\n\n"));
  }
  if (side === "T") {
    const a = norm(WW.collabT_text || "");
    const b = norm(WW.collabT_q || "");
    return norm([a, b].filter(Boolean).join("\n\n"));
  }

  return "";
}

function splitTipsInto3(tips) {
  const t = norm(tips || "");
  if (!t) return ["", "", ""];
  const parts = t.split(/\n{2,}/).map(s => s.trim()).filter(Boolean);
  const out = [parts[0] || "", parts[1] || "", parts[2] || ""];
  if (parts.length <= 3) return out;
  out[2] = norm([out[2], ...parts.slice(3)].filter(Boolean).join("\n\n"));
  return out;
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

    // Prefer textV2, then text, then top-level
    const T = isObj(src?.textV2) ? src.textV2 : (isObj(src?.text) ? src.text : (isObj(src?.fields) ? src.fields : {}));

    const P = {
      name:      norm(src?.person?.fullName || src?.fullName || "Perspective Overlay"),
      dateLbl:   norm(src?.dateLbl || ""),

      summary:        norm(T?.summary || src?.summary || ""),
      ctrl_overview:  norm(T?.ctrl_overview || src?.ctrl_overview || ""),
      ctrl_overviewQ: norm(T?.ctrl_overviewQ || src?.ctrl_overviewQ || ""),
      sequence:       norm(T?.sequence || src?.sequence || ""),
      themepair:      norm(T?.themepair || src?.themepair || ""),
      tips:           norm(T?.tips || src?.tips || ""),

      workWith: isObj(src?.workWith) ? src.workWith : null
    };

    const hasText = !!(isObj(src?.text) || isObj(src?.textV2));
    const chartUrl = norm(
      src?.chartUrl ||
      src?.spiderChartUrl ||
      src?.chart?.spiderUrl ||
      ""
    );

    if (debugMode) {
      const [a1,a2,a3] = splitTipsInto3(P.tips);

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
          hasText,
          fields: {
            summary_len: P.summary.length,
            ctrl_overview_len: P.ctrl_overview.length,
            ctrl_overviewQ_len: P.ctrl_overviewQ.length,
            sequence_len: P.sequence.length,
            themepair_len: P.themepair.length,
            tips_len: P.tips.length
          },
          hasWorkWith: !!P.workWith,
          workWithKeys: P.workWith ? Object.keys(P.workWith).slice(0, 80) : [],
          page6Preview: {
            collabCol_len: combineWorkWithSide(P.workWith, "C").length,
            collabLe_len:  combineWorkWithSide(P.workWith, "T").length
          },
          page7Preview: {
            act1_len: a1.length, act2_len: a2.length, act3_len: a3.length
          }
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

    const L = DEFAULT_LAYOUT.pages;

    // p1: name & date
    if (p1 && P.name)    drawTextBox(p1, fontBold, P.name,    L.p1.name, { maxLines: L.p1.name.maxLines ?? 1 });
    if (p1 && P.dateLbl) drawTextBox(p1, font,     P.dateLbl, L.p1.date, { maxLines: L.p1.date.maxLines ?? 1 });

    // headers (p2–p8)
    const putHeader = (page, spec) => {
      if (!page || !P.name || !spec) return;
      drawTextBox(page, font, P.name, spec, { maxLines: spec.maxLines ?? 1 });
    };
    putHeader(p2, L.p2.hdrName);
    putHeader(p3, L.p3.hdrName);
    putHeader(p4, L.p4.hdrName);
    putHeader(p5, L.p5.hdrName);
    putHeader(p6, L.p6.hdrName);
    putHeader(p7, L.p7.hdrName);
    putHeader(p8, L.p8.hdrName);

    // chart on page 4
    const CHART4 = { ...L.p4Text.chart };
    if (q.cx != null) CHART4.x = N(q.cx, CHART4.x);
    if (q.cy != null) CHART4.y = N(q.cy, CHART4.y);
    if (q.cw != null) CHART4.w = N(q.cw, CHART4.w);
    if (q.ch != null) CHART4.h = N(q.ch, CHART4.h);

    if (p4 && chartUrl) {
      chartFetch = await fetchBytes(chartUrl, 9000);
      if (chartFetch.ok && chartFetch.bytes && chartFetch.bytes.length) {
        let img = null;
        if (looksPng(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedPng(chartFetch.bytes);
        else if (looksJpg(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedJpg(chartFetch.bytes);

        if (img) {
          const { x, y, w, h } = CHART4;
          const pageH = p4.getHeight();
          const yBottom = pageH - y - h;
          p4.drawImage(img, { x, y: yBottom, width: w, height: h });
        }
      }
    }

    // p3: summary
    if (p3 && P.summary) {
      drawOverlayBox(p3, fonts, P.summary, L.p3Text.summary);
    }

    // p4: ctrl_overview + ctrl_overviewQ (2 clean blobs, no auto-flow)
    if (p4) {
      if (P.ctrl_overview)  drawOverlayBox(p4, fonts, P.ctrl_overview,  L.p4Text.ctrl_overview);
      if (P.ctrl_overviewQ) drawOverlayBox(p4, fonts, P.ctrl_overviewQ, L.p4Text.ctrl_overviewQ);
    }

    // p5: sequence + theme
    if (p5 && P.sequence) {
      drawOverlayBox(p5, fonts, P.sequence, L.p5Text.sequence);
    }
    if (p5 && P.themepair) {
      drawOverlayBox(p5, fonts, P.themepair, L.p5Text.theme);
    }

    // p6: WorkWith columns (only if payload provides workWith)
    if (p6) {
      const WW = P.workWith;
      const col = combineWorkWithSide(WW, "C");
      const le  = combineWorkWithSide(WW, "T");

      if (col) drawOverlayBox(p6, fonts, col, L.p6WorkWith.collabCol);
      if (le)  drawOverlayBox(p6, fonts, le,  L.p6WorkWith.collabLe);
    }

    // p7: actions (tips split into 3)
    if (p7 && P.tips) {
      const [act1, act2, act3] = splitTipsInto3(P.tips);
      if (act1) drawOverlayBox(p7, fonts, act1, L.p7Actions.act1);
      if (act2) drawOverlayBox(p7, fonts, act2, L.p7Actions.act2);
      if (act3) drawOverlayBox(p7, fonts, act3, L.p7Actions.act3);
    }

    const bytes = await pdfDoc.save();

    const outName = S(q.out || `CTRL_180_${P.name || "Perspective"}_${P.dateLbl || ""}.pdf`)
      .replace(/[^\w.-]+/g, "_");

    // Response headers
    res.setHeader("X-CTRL-TPL", tpl);
    res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
    res.setHeader("X-CTRL-PAYLOAD-SIZE", String(payloadSize || 0));
    res.setHeader("X-CTRL-CHART", chartUrl ? "1" : "0");
    res.setHeader("X-CTRL-CHART-TARGET", "p4");
    res.setHeader("X-CTRL-CHART-FETCH", chartFetch?.ok ? "ok" : (chartFetch?.reason || "no"));
    res.setHeader("X-CTRL-WORKWITH", P.workWith ? "1" : "0");
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
