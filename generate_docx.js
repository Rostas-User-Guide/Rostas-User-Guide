/**
 * Rostas Coordinator Guide — HTML → DOCX Converter
 *
 * Produces a fully self-contained DOCX:
 *   - Static dot-leader TOC (no Word field refresh needed)
 *   - Works in Google Docs, Word Online, LibreOffice, Word
 *   - All images embedded, callout boxes, step boxes, lists
 *
 * Usage:
 *   node generate_docx.js                                   ← HTML + images from GitHub
 *   node generate_docx.js --local ./index.html              ← local HTML, images from GitHub
 *   node generate_docx.js --local ./index.html --images /path/to/imgs  ← fully offline
 *
 * Add this script to the GitHub repo alongside index.html.
 * Run it any time index.html changes to regenerate the guide.
 */

'use strict';

const fs    = require('fs');
const path  = require('path');
const https = require('https');
const http  = require('http');

const cheerio = require('cheerio');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, UnderlineType, Tab,
  TabStopType, LeaderType,
} = require('docx');

// ─── Config ───────────────────────────────────────────────────────────────────
const GITHUB_RAW  = 'https://raw.githubusercontent.com/Rostas-User-Guide/Rostas-User-Guide/main/';
const GITHUB_HTML = GITHUB_RAW + 'index.html';
const OUTPUT_FILE = process.env.OUTPUT_FILE || './Rostas_Coordinator_Guide.docx';

// A4 portrait, ~2cm margins
const PAGE_W    = 11906;
const PAGE_H    = 16838;
const MARGIN    = 1134;
const CONTENT_W = PAGE_W - MARGIN * 2;  // 9638 DXA

// Approximate lines per page (used for page number estimation)
// Calibrated for A4, 11pt Calibri, ~2cm margins
const LINES_PER_PAGE = 38;

// Colour palette
const C = {
  heading:  '1B3A5C',
  caption:  '777777',
  pill:     '555566',
  infoBg:   'EBF4FB', infoBar:  '2196F3',
  warnBg:   'FFF8E1', warnBar:  'FF9800',
  tipBg:    'E8F5E9', tipBar:   '4CAF50',
  stepBg:   'F3F0FF', stepBar:  '7C4DFF',
  footer:   '999999', link:     '0563C1',
  tocH1:    '1B3A5C', tocH2:    '444455',
  tocDot:   'AAAAAA',
};

// docx-js ImageRun.transformation takes pixels (96dpi). 1 inch = 1440 DXA = 96px.
function dxaToEmu(dxa) { return Math.round(dxa * 96 / 1440); }

// Scale image to fit max width, preserving real aspect ratio
function imgSize(cls, imgW, imgH) {
  let maxW;
  if (cls.includes('screenshot-mid')) maxW = Math.round(CONTENT_W * 0.56);
  else if (cls.includes('screenshot-sm'))  maxW = Math.round(CONTENT_W * 0.38);
  else maxW = CONTENT_W;
  // Use actual aspect ratio if we have real dimensions
  if (imgW && imgH && imgH > 0) {
    const w = maxW;
    const h = Math.round(w * imgH / imgW);
    return { w, h };
  }
  // Fallback estimates
  if (cls.includes('screenshot-mid')) return { w: maxW, h: Math.round(maxW * 0.70) };
  if (cls.includes('screenshot-sm'))  return { w: maxW, h: Math.round(maxW * 0.90) };
  return { w: maxW, h: Math.round(maxW * 0.55) };
}

// ─── HTTP fetch ───────────────────────────────────────────────────────────────
function fetchURL(url) {
  return new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? https : http;
    mod.get(url, res => {
      if ([301,302,307,308].includes(res.statusCode))
        return fetchURL(res.headers.location).then(resolve).catch(reject);
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end',  () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ─── Image loading ────────────────────────────────────────────────────────────
function buildLocalMap(dir) {
  if (!dir || !fs.existsSync(dir)) return {};
  const map = {};
  for (const f of fs.readdirSync(dir)) {
    if (!/\.(png|jpe?g|gif|webp)$/i.test(f)) continue;
    const full = path.join(dir, f);
    map[f.toLowerCase()]                  = full;
    map[f.replace(/_/g,'-').toLowerCase()] = full;
    map[f.replace(/-/g,'_').toLowerCase()] = full;
  }
  return map;
}

// Detect real image type and dimensions from file header — never trust the extension
function detectImageInfo(buf) {
  // PNG
  if (buf[0]===0x89&&buf[1]===0x50&&buf[2]===0x4E&&buf[3]===0x47) {
    const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
    return { type: 'png', w, h };
  }
  // JPEG — scan for SOF marker
  if (buf[0]===0xFF&&buf[1]===0xD8&&buf[2]===0xFF) {
    let i = 2;
    while (i < buf.length - 8) {
      if (buf[i] !== 0xFF) break;
      const marker = buf[i+1];
      const len = buf.readUInt16BE(i+2);
      if (marker >= 0xC0 && marker <= 0xC3) {
        const h = buf.readUInt16BE(i+5), w = buf.readUInt16BE(i+7);
        return { type: 'jpeg', w, h };
      }
      i += 2 + len;
    }
    return { type: 'jpeg', w: 0, h: 0 };
  }
  // WebP
  if (buf[0]===0x52&&buf[1]===0x49&&buf[2]===0x46&&buf[3]===0x46&&
      buf[8]===0x57&&buf[9]===0x45&&buf[10]===0x42&&buf[11]===0x50)
    return { type: 'webp', w: 0, h: 0 };
  // GIF
  if (buf[0]===0x47&&buf[1]===0x49&&buf[2]===0x46)
    return { type: 'gif', w: 0, h: 0 };
  return { type: 'png', w: 0, h: 0 };
}

async function loadImage(src, localMap) {
  const local = localMap[src.toLowerCase()];
  if (local && fs.existsSync(local)) {
    const data = fs.readFileSync(local);
    const info = detectImageInfo(data);
    return { data, ...info };
  }
  try {
    const data = await fetchURL(GITHUB_RAW + encodeURIComponent(src));
    const info = detectImageInfo(data);
    return { data, ...info };
  } catch { return null; }
}

// ─── Inline text runs ─────────────────────────────────────────────────────────
function inlineRuns($, node, style = {}) {
  const runs = [];
  function walk(n, s) {
    if (n.type === 'text') {
      const txt = n.data
        .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
        .replace(/&nbsp;/g,' ').replace(/&#x2019;/g,'\u2019')
        .replace(/&#x201C;/g,'\u201C').replace(/&#x201D;/g,'\u201D')
        .replace(/&#x2018;/g,'\u2018');
      if (txt) runs.push(new TextRun({ text: txt, ...s }));
      return;
    }
    if (n.type !== 'tag') return;
    const tag = n.name.toLowerCase();
    const ch  = n.children || [];
    switch (tag) {
      case 'strong': case 'b': ch.forEach(c => walk(c, { ...s, bold: true })); break;
      case 'em':     case 'i': ch.forEach(c => walk(c, { ...s, italics: true })); break;
      case 'code':
        runs.push(new TextRun({ text: $(n).text(), font: 'Courier New', size: 20, ...s })); break;
      case 'br': runs.push(new TextRun({ break: 1 })); break;
      case 'a': {
        const href = $(n).attr('href') || '';
        if (href.startsWith('http')) {
          runs.push(new ExternalHyperlink({ link: href,
            children: [new TextRun({ text: $(n).text(), style: 'Hyperlink' })] }));
        } else ch.forEach(c => walk(c, s));
        break;
      }
      case 'span':
        if ($(n).hasClass('callout-icon')) break;
        ch.forEach(c => walk(c, s)); break;
      default: ch.forEach(c => walk(c, s));
    }
  }
  (node.children || []).forEach(n => walk(n, style));
  return runs.length ? runs : [new TextRun({ text: '' })];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function spacer(pts = 80) {
  return new Paragraph({ children: [], spacing: { before: pts, after: pts } });
}

function calloutBox(paras, barCol, bgCol) {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [140, CONTENT_W - 140],
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 140, type: WidthType.DXA },
        shading: { fill: barCol, type: ShadingType.CLEAR },
        borders: { top: none, bottom: none, left: none, right: none },
        children: [new Paragraph({ children: [] })],
      }),
      new TableCell({
        width: { size: CONTENT_W - 140, type: WidthType.DXA },
        shading: { fill: bgCol, type: ShadingType.CLEAR },
        borders: { top: none, bottom: none, left: none, right: none },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: paras,
      }),
    ]})],
  });
}

function blockChildren($, node) {
  const out = [];
  for (const ch of (node.children || [])) {
    if (ch.type !== 'tag') continue;
    const t = ch.name.toLowerCase();
    if (t === 'p') {
      out.push(new Paragraph({ children: inlineRuns($, ch), spacing: { before: 60, after: 80 } }));
    } else if (t === 'ul') {
      $(ch).children('li').each((_,li) => out.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        children: inlineRuns($, li), spacing: { before: 40, after: 40 },
      })));
    } else if (t === 'ol') {
      $(ch).children('li').each((_,li) => out.push(new Paragraph({
        numbering: { reference: 'numbers', level: 0 },
        children: inlineRuns($, li), spacing: { before: 40, after: 40 },
      })));
    }
  }
  return out;
}

// ─── Page estimation ──────────────────────────────────────────────────────────
// Estimate the "line cost" of each content element so we can predict page numbers.
// This is intentionally rough — good enough for navigation, not for citation.
function estimateCost(el) {
  const $el = global._$ ? global._$(el) : null;
  if (!el || !el.name) return 0;
  const tag = el.name.toLowerCase();
  if (tag === 'h2') return LINES_PER_PAGE;           // always starts new page
  if (tag === 'h3') return 3;                        // subheading
  if (tag === 'p')  return 2;                        // paragraph
  if (tag === 'ul' || tag === 'ol') {
    const items = (el.children || []).filter(c => c.name === 'li').length;
    return items * 1.2;
  }
  if (tag === 'figure') return 14;                   // image + caption
  if (tag === 'img')    return 12;
  if (tag === 'div') {
    if ($el && $el.hasClass('callout')) return 3;
    if ($el && $el.hasClass('step-box')) return 4;
    if ($el && ($el.hasClass('pill-row') || $el.hasClass('pill-grid'))) return 1;
    return 0;  // generic div — costs handled by children
  }
  return 0;
}

// ─── Heading extraction (first pass) ─────────────────────────────────────────
// Walk #main and collect all headings in document order with their cumulative
// line costs. Returns: [{ level, text, linesBefore }]
function extractHeadings($) {
  const headings = [];
  let lines = 0;
  global._$ = $;

  function walk(el) {
    if (!el || el.type !== 'tag') return;
    const $el = $(el);
    const tag = el.name.toLowerCase();

    // Skip carousel
    if ($el.hasClass('carousel') || $el.hasClass('carousel-container')) return;

    if (tag === 'h2' && $el.hasClass('section-title')) {
      lines += LINES_PER_PAGE; // page break before each h2
      headings.push({ level: 1, text: $el.text().trim(), linesBefore: lines });
      lines += 2;
      return;
    }
    if (tag === 'h3') {
      headings.push({ level: 2, text: $el.text().trim(), linesBefore: lines });
      lines += 3;
      return;
    }

    // Tally cost of other elements
    const cost = estimateCost(el);
    if (cost > 0 && !['div','section'].includes(tag)) {
      lines += cost;
      return;
    }

    // Recurse into containers
    for (const ch of (el.children || [])) walk(ch);
  }

  const main = $('#main');
  if (!main.length) return headings;
  for (const ch of (main.get(0).children || [])) walk(ch);
  return headings;
}

// Convert cumulative lines → page number
// Cover = 1, TOC = 2-3, content starts page 4
function linesToPage(linesBefore) {
  return 3 + Math.ceil(linesBefore / LINES_PER_PAGE);
}

// ─── Word field TOC (rendered correctly by LibreOffice → PDF) ─────────────────
function buildTOC() {
  return [
    new Paragraph({
      children: [new TextRun({ text: 'Contents', bold: true, size: 36, color: C.heading, font: 'Calibri' })],
      spacing: { before: 200, after: 200 },
    }),
    new TableOfContents('Contents', {
      hyperlink: true,
      headingStyleRange: '1-2',
      stylesWithLevels: [
        { styleId: 'Heading1', level: 1 },
        { styleId: 'Heading2', level: 2 },
      ],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ─── Slides section (2 per page) ─────────────────────────────────────────────
function buildSlides(imgs) {
  const out = [];
  out.push(new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: 'Quick Start Guide — Orchestrating with Empathy', bold: true, color: C.heading })],
    spacing: { before: 200, after: 240 },
  }));

  // Two slides per page — half-width each to fit side by side isn't possible in
  // docx (no inline columns), so we stack two per page with a page break every 2
  const SLIDE_W = CONTENT_W;
  const pairs = [];
  for (let i = 1; i <= 15; i++) {
    const src = 'slide-' + String(i).padStart(2, '0') + '.png';
    pairs.push(src);
  }

  for (let i = 0; i < pairs.length; i++) {
    const src = pairs[i];
    if (imgs[src]) {
      // Half-height to fit two per page (A4 usable height ~13200 DXA, half = 6600)
      const aspectW = imgs[src].w || 16;
      const aspectH = imgs[src].h || 9;
      const w = SLIDE_W;
      const h = Math.round(w * aspectH / aspectW);
      // Cap height to half-page so two fit
      const maxH = 5600; // ~half A4 page in DXA
      const finalW = h > maxH ? Math.round(w * maxH / h) : w;
      const finalH = h > maxH ? maxH : h;

      out.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: imgs[src].data, type: imgs[src].type,
          transformation: { width: dxaToEmu(finalW), height: dxaToEmu(finalH) } })],
        spacing: { before: 60, after: 60 },
      }));

      // Page break after every 2nd slide, or after the last one
      const isSecond = (i + 1) % 2 === 0;
      const isLast   = i === pairs.length - 1;
      if (isSecond || isLast) {
        out.push(new Paragraph({ children: [new PageBreak()] }));
      }
    }
  }
  return out;
}

// ─── Content parser (second pass) ────────────────────────────────────────────
async function buildContent($, imgs) {
  const out = [];
  let firstSection = true;

  async function processEl(el) {
    const $el = $(el);
    const tag = (el.name || '').toLowerCase();

    if ($el.hasClass('carousel') || $el.hasClass('carousel-container') ||
        $el.hasClass('carousel-eyebrow') || $el.hasClass('carousel-title')) return;

    // h2 → Heading 1 (page break before, except first)
    if (tag === 'h2' && $el.hasClass('section-title')) {
      out.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: $el.text().trim(), bold: true, color: C.heading })],
        spacing: { before: firstSection ? 0 : 400, after: 140 },
        pageBreakBefore: !firstSection,
      }));
      firstSection = false;
      return;
    }

    // h3 → Heading 2
    if (tag === 'h3') {
      out.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: $el.text().trim(), bold: true, color: C.heading })],
        spacing: { before: 220, after: 80 },
      }));
      return;
    }

// Pill row → render actual pill images inline (fall back to text if missing)
    if ($el.hasClass('pill-row') || $el.hasClass('pill-grid')) {
      const children = [];
      // 34px CSS height → DXA then to pixels via dxaToEmu
      const PILL_H_DXA = Math.round(34 * 1440 / 96); // 510 DXA
      $el.find('img.nav-pill').each((_, img) => {
        const src = $(img).attr('src') || '';
        const lbl = ($(img).attr('alt') || '').replace(/ view$/i, '').trim();
        if (src && imgs[src]) {
          const { w: iw, h: ih } = imgs[src];
          const pillW = (iw && ih) ? Math.round(PILL_H_DXA * iw / ih)
                                   : Math.round(PILL_H_DXA * 3.2);
          children.push(new ImageRun({
            data: imgs[src].data,
            type: imgs[src].type,
            transformation: { width: dxaToEmu(pillW), height: dxaToEmu(PILL_H_DXA) },
          }));
          children.push(new TextRun({ text: '  ' })); // gap between pills
        } else {
          // Fallback: text label if image missing
          if (children.length) children.push(new TextRun({ text: '  •  ', size: 18, color: C.pill }));
          children.push(new TextRun({ text: lbl, size: 18, color: C.pill }));
        }
      });
      if (children.length) out.push(new Paragraph({
        children,
        spacing: { before: 60, after: 80 },
      }));
      return;
    }

    // Callout
    if (tag === 'div' && $el.hasClass('callout')) {
      let bar = C.infoBar, bg = C.infoBg;
      if ($el.hasClass('callout-warn')) { bar = C.warnBar; bg = C.warnBg; }
      if ($el.hasClass('callout-tip'))  { bar = C.tipBar;  bg = C.tipBg;  }
      const clone = $el.clone(); clone.find('.callout-icon').remove();
      const paras = clone.children('p,ul,ol').length
        ? blockChildren($, clone.get(0))
        : [new Paragraph({ children: inlineRuns($, clone.get(0)), spacing: { before: 60, after: 60 } })];
      out.push(spacer(80));
      out.push(calloutBox(paras, bar, bg));
      out.push(spacer(80));
      return;
    }

    // Step box
    if (tag === 'div' && $el.hasClass('step-box')) {
      const label = $el.find('.step-label').first().text().trim();
      const clone = $el.clone(); clone.find('.step-label').remove();
      const paras = [];
      if (label) paras.push(new Paragraph({
        children: [new TextRun({ text: label, bold: true, color: C.stepBar, size: 22 })],
        spacing: { before: 40, after: 60 },
      }));
      paras.push(...blockChildren($, clone.get(0)));
      if (!paras.length) paras.push(new Paragraph({ children: [new TextRun($el.text().trim())] }));
      out.push(spacer(80));
      out.push(calloutBox(paras, C.stepBar, C.stepBg));
      out.push(spacer(80));
      return;
    }

    // Paragraph
    if (tag === 'p') {
      out.push(new Paragraph({ children: inlineRuns($, el), spacing: { before: 60, after: 100 } }));
      return;
    }

    // Lists
    if (tag === 'ul') {
      $el.children('li').each((_,li) => out.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        children: inlineRuns($, li), spacing: { before: 40, after: 40 },
      })));
      return;
    }
    if (tag === 'ol') {
      $el.children('li').each((_,li) => out.push(new Paragraph({
        numbering: { reference: 'numbers', level: 0 },
        children: inlineRuns($, li), spacing: { before: 40, after: 40 },
      })));
      return;
    }

    // Figure
    if (tag === 'figure') {
      const img = $el.find('img').first();
      const cap = $el.find('figcaption').first();
      const src = img.attr('src') || '';
      const cls = img.attr('class') || '';
      if (src && imgs[src]) {
        const { w, h } = imgSize(cls, imgs[src].w, imgs[src].h);
        out.push(spacer(100));
        out.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new ImageRun({ data: imgs[src].data, type: imgs[src].type,
            transformation: { width: dxaToEmu(w), height: dxaToEmu(h) } })],
        }));
        if (cap.length && cap.text().trim()) {
          out.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: cap.text().trim(), italics: true, color: C.caption, size: 18 })],
            spacing: { before: 40, after: 120 },
          }));
        }
      } else if (src && !imgs[src]) {
        out.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `[Image: ${src}]`, italics: true, color: 'BBBBBB', size: 18 })],
          spacing: { before: 40, after: 60 },
        }));
      }
      return;
    }

    // Standalone img
    if (tag === 'img') {
      const src = $el.attr('src') || '';
      const cls = $el.attr('class') || '';
      if (src && !src.startsWith('pill-') && !src.startsWith('slide-') && imgs[src]) {
        const { w, h } = imgSize(cls, imgs[src].w, imgs[src].h);
        out.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new ImageRun({ data: imgs[src].data, type: imgs[src].type,
            transformation: { width: dxaToEmu(w), height: dxaToEmu(h) } })],
          spacing: { before: 100, after: 100 },
        }));
      }
      return;
    }

    // Containers — recurse
    if (['div','section','article','aside','main'].includes(tag)) {
      for (const ch of (el.children || [])) {
        if (ch.type === 'tag') await processEl(ch);
      }
    }
  }

  const main = $('#main');
  if (!main.length) { console.error('  ✗ #main not found'); return out; }
  for (const ch of (main.get(0).children || [])) {
    if (ch.type === 'tag') await processEl(ch);
  }
  return out;
}

// ─── Entry point ──────────────────────────────────────────────────────────────
async function main() {
  const args     = process.argv.slice(2);
  const localIdx = args.indexOf('--local');
  const imgsIdx  = args.indexOf('--images');
  const localHTML = localIdx >= 0 ? args[localIdx + 1] : null;
  const imgsDir   = imgsIdx  >= 0 ? args[imgsIdx  + 1] : '/mnt/project';

  console.log('🚀  Rostas Coordinator Guide  →  DOCX');
  console.log('═'.repeat(45));

  // 1. Load HTML
  let html;
  if (localHTML) {
    console.log(`📄  Reading: ${localHTML}`);
    html = fs.readFileSync(localHTML, 'utf8');
  } else {
    console.log('📥  Fetching HTML from GitHub...');
    html = (await fetchURL(GITHUB_HTML)).toString('utf8');
  }
  console.log(`    ✓ ${Math.round(html.length/1024)} KB`);

  const $ = cheerio.load(html, { decodeEntities: false });

  // 2. First pass — extract headings for TOC
  console.log('\n📑  Extracting headings for TOC...');
  const headings = extractHeadings($);
  console.log(`    ✓ ${headings.filter(h=>h.level===1).length} sections,` +
              ` ${headings.filter(h=>h.level===2).length} sub-sections`);

  // 3. Load images
  const srcs = new Set();
  $('img').each((_,el) => {
    const s = $(el).attr('src') || '';
    if (s && !s.startsWith('http') && !s.startsWith('data:')) srcs.add(s);
  });
  // Explicitly include all 15 slides (they load dynamically in the HTML carousel)
  for (let i = 1; i <= 15; i++) {
    srcs.add('slide-' + String(i).padStart(2, '0') + '.png');
  }

  const localMap = buildLocalMap(imgsDir);

  console.log(`\n📸  Loading ${srcs.size} images...`);
  const imgs = {};
  let i = 0;
  for (const src of srcs) {
    i++;
    process.stdout.write(`    [${String(i).padStart(2)}/${srcs.size}] ${src.padEnd(32)}`);
    imgs[src] = await loadImage(src, localMap);
    console.log(imgs[src] ? '✓' : '✗');
  }
  const found = Object.values(imgs).filter(Boolean).length;
  console.log(`    ${found}/${srcs.size} images loaded`);

  // 4. Build Word field TOC (LibreOffice will render correct page numbers in PDF)
  console.log('\n📋  Building TOC...');
  const tocParas = buildTOC();
  console.log(`    ✓ TOC field inserted (${headings.length} headings detected)`);

  // 5. Parse content
  console.log('\n📝  Parsing content...');
  const content = await buildContent($, imgs);
  console.log(`    ✓ ${content.length} elements`);
  if (content.length < 5) { console.error('    ✗ Too few elements — aborting'); process.exit(1); }

  // 6. Cover page
  const logo = imgs['Logo.png'];
  const cover = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2800, after: 600 },
      children: logo
        ? [new ImageRun({ data: logo.data, type: logo.type,
            transformation: { width: dxaToEmu(3600), height: dxaToEmu(1440) } })]
        : [new TextRun({ text: 'Rostas', size: 64, bold: true, color: C.heading })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 500, after: 200 },
      children: [new TextRun({ text: 'Coordinator Guide', size: 52, bold: true, color: C.heading, font: 'Calibri' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 160 },
      children: [new TextRun({ text: 'Knox Church Waitara', size: 30, color: '444444', font: 'Calibri' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: new Date().toLocaleDateString('en-NZ', { month: 'long', year: 'numeric' }),
        size: 24, color: '999999', font: 'Calibri',
      })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];

  // 7. Footer
  const footer = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' } },
      children: [
        new TextRun({ text: 'Rostas Coordinator Guide  •  Knox Church Waitara  •  p.\u00A0', size: 18, color: C.footer }),
        new TextRun({ children: [PageNumber.CURRENT], size: 18, color: C.footer }),
      ],
    })],
  });

  // 8. Assemble
  console.log('\n📄  Assembling DOCX...');
  const doc = new Document({
    numbering: {
      config: [
        { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ],
    },
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 30, bold: true, font: 'Calibri', color: C.heading },
          paragraph: { spacing: { before: 400, after: 140 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Calibri', color: C.heading },
          paragraph: { spacing: { before: 220, after: 80 }, outlineLevel: 1 } },
        { id: 'Hyperlink', name: 'Hyperlink', basedOn: 'Default Paragraph Font',
          run: { color: C.link, underline: { type: UnderlineType.SINGLE } } },
      ],
    },
    features: { updateFields: true },
    sections: [
      // Cover — no footer
      {
        properties: { page: { size: { width: PAGE_W, height: PAGE_H },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } },
        children: cover,
      },
      // TOC + content — with footer
      {
        properties: { page: { size: { width: PAGE_W, height: PAGE_H },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } },
        footers: { default: footer },
        children: [...tocParas, ...buildSlides(imgs), ...content],
      },
    ],
  });

  // 9. Write
  console.log(`\n💾  Writing ${OUTPUT_FILE}`);
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT_FILE, buf);

  console.log(`\n✅  Done — ${(buf.length/1024/1024).toFixed(1)} MB`);
  console.log('   No refresh needed — TOC is static and works in any app.');
  console.log('   Open in Google Docs and print to PDF for the final version.\n');
}

main().catch(err => {
  console.error('\n❌  Error:', err.message);
  console.error(err.stack);
  process.exit(1);
});
