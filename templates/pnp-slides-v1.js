/**
 * pnp-slides-v1.js
 * PNP / Playbook slide deck renderer — hosted on pnp-skill-assets CDN.
 * Entry point: buildPptx(slidesData, opts)
 *
 * Canvas: 13.33" × 7.5" (LAYOUT_WIDE, 1920×1080 at 144 DPI)
 * Font:   pt(cssPx) = cssPx / 2  →  CSS 40px = 20pt ✓
 * Coords: ix(px) = px * 13.33/1920,  iy(px) = px * 7.5/1080
 */

/* ── Logo assets (PNG, three colour variants) ─────────────────────────── */
const _CDN = 'https://cdn.jsdelivr.net/gh/aromero-pnptc/pnp-skill-assets@main/logos/';
const LOGO_URLS = {
  pnp:      { color: _CDN + 'pnp-logo%201D2A57.png',
              muted: _CDN + 'pnp-logo%2094A0B8.png',
              white: _CDN + 'pnp-logo%20white.png' },
  playbook: { color: _CDN + 'playbook-logo%201D2A57.png',
              muted: _CDN + 'playbook-logo%2094A0B8.png',
              white: _CDN + 'playbook-logo%20white.png' },
};

/* ── Slide canvas ─────────────────────────────────────────────────────── */
const SW = 13.33, SH = 7.5;
const ix = n => +(n * SW / 1920).toFixed(4);
const iy = n => +(n * SH / 1080).toFixed(4);
const pt = n => n / 2;   // CSS px → PPT pt at 144 DPI

/* ── Design tokens (light theme — v1 only) ────────────────────────────── */
const C = {
  bg:           'F5F5F7',   // --pb-athens-gray
  card:         'FFFFFF',
  cardBorder:   'E5E9F0',   // --pb-seattle
  rule:         'C6CEDC',   // --pb-heather
  text1:        '1D2A57',   // --pb-oxford-blue
  text2:        '44527A',   // --pb-dove-gray
  muted:        '94A0B8',   // --pb-suva-gray
  accent:       '3333CC',   // --pb-playbook-purple
  accentChip:   'EBEBFA',   // --pb-playbook-100 (tag bg)
  photoBg:      'D8DCEB',   // photo placeholder
};

/* ── Chrome geometry (px on 1920×1080 grid) ───────────────────────────── */
// All values derived from playbook.css measurements.
const G = {
  HEADER_H:   100,  // padding:40 + logo:20 + padding:40
  BAR_Y:      100,  // immediately after header
  BAR_W:       40,  // width:40px
  BAR_H:        4,  // height:4px
  TITLE_Y:    118,  // BAR_Y + BAR_H + margin-bottom:14
  TITLE_H:     60,  // 40px font × 1.05 lh + buffer
  CONTENT_Y:  200,  // TITLE_Y + TITLE_H + 16px content padding-top ≈ 194 → 200
  CONTENT_B: 1000,  // 1080 - 80px bottom padding
};
G.CONTENT_H = G.CONTENT_B - G.CONTENT_Y;  // 800

/* ── Zone helpers ─────────────────────────────────────────────────────── */
function contentZone() {
  return { x: 80, y: G.CONTENT_Y, w: 1760, h: G.CONTENT_H };
}

// Divide zone into cols × rows equal cells separated by gap px.
function gridZones(zone, cols, rows, gap) {
  rows = rows || 1; gap = gap ?? 24;
  const cw = (zone.w - gap * (cols - 1)) / cols;
  const ch = (zone.h - gap * (rows - 1)) / rows;
  const out = [];
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++)
      out.push({ x: zone.x + c * (cw + gap), y: zone.y + r * (ch + gap), w: cw, h: ch });
  return out;
}

// Stack n equal-height rows within zone, separated by gap px.
function stackZones(zone, n, gap) {
  return gridZones(zone, 1, n, gap ?? 24);
}

/* ── Drawing primitives ───────────────────────────────────────────────── */
function _txt(slide, text, zone, opts) {
  opts = opts || {};
  slide.addText(text || '', {
    x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h),
    fontSize:    opts.size   || 9,
    bold:        opts.bold   || false,
    color:       opts.color  || C.text1,
    fontFace:    'Inter',
    align:       opts.align  || 'left',
    valign:      opts.valign || 'middle',
    wrap:        opts.wrap !== false,
    margin:      0,
    charSpacing: opts.tracking || 0,
  });
}

function _rect(slide, zone, fill, opts) {
  opts = opts || {};
  slide.addShape('rect', {
    x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h),
    fill: { color: fill },
    line: opts.border
      ? { color: opts.borderColor || C.cardBorder, width: 1 }
      : { type: 'none' },
  });
}

// Rounded rect. rectRadius in inches: ix(px) gives correct scale conversion.
// pill:true → use rectRadius:0.5 (PptxGenJS caps at max, gives full pill).
function _rRect(slide, zone, fill, opts) {
  opts = opts || {};
  slide.addShape('roundRect', {
    x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h),
    fill: { color: fill },
    line: opts.border
      ? { color: opts.borderColor || C.cardBorder, width: 1 }
      : { type: 'none' },
    rectRadius: opts.pill ? 0.5 : (opts.r != null ? opts.r : ix(12)),
  });
  // ix(12) = 12 * 13.33/1920 ≈ 0.0833" — matches CSS border-radius:12px at this canvas size
}

/* ── Image loading ────────────────────────────────────────────────────── */
const _cache = new Map();

async function _fetch(url) {
  if (!url) return null;
  if (_cache.has(url)) return _cache.get(url);
  try {
    const res  = await fetch(url);
    const blob = await res.blob();
    const data = await new Promise((ok, err) => {
      const fr = new FileReader();
      fr.onload  = () => ok(fr.result);
      fr.onerror = err;
      fr.readAsDataURL(blob);
    });
    _cache.set(url, data);
    return data;
  } catch (_) { return null; }
}

async function _preload(urls) {
  await Promise.all([...new Set(urls)].filter(Boolean).map(_fetch));
}

function _img(slide, url, zone, opts) {
  opts = opts || {};
  // Always draw placeholder so the region is never blank
  const ph = { fill: { color: opts.ph || C.photoBg }, line: { type: 'none' } };
  if (opts.r != null || opts.pill) {
    ph.rectRadius = opts.pill ? 0.5 : opts.r;
    slide.addShape('roundRect', { x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h), ...ph });
  } else {
    slide.addShape('rect', { x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h), ...ph });
  }
  const data = _cache.get(url);
  if (!data) return;
  try {
    slide.addImage({ data, x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h) });
  } catch (_) {}
}

/* ── Logo loading ─────────────────────────────────────────────────────── */
async function _loadLogos() {
  const urls = [];
  for (const k of ['pnp', 'playbook'])
    for (const v of ['color', 'muted', 'white'])
      urls.push(LOGO_URLS[k][v]);
  await _preload(urls);
}

/* ── Gradient overlay (for cover-photo blend edge) ───────────────────── */
function _gradientPng(hex, w) {
  const r = parseInt(hex.slice(0,2), 16),
        g = parseInt(hex.slice(2,4), 16),
        b = parseInt(hex.slice(4,6), 16);
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = 2;
  const ctx = cv.getContext('2d');
  const grad = ctx.createLinearGradient(0,0,w,0);
  grad.addColorStop(0, `rgba(${r},${g},${b},1)`);
  grad.addColorStop(1, `rgba(${r},${g},${b},0)`);
  ctx.fillStyle = grad; ctx.fillRect(0,0,w,2);
  return cv.toDataURL('image/png');
}

/* ── Chrome helpers ───────────────────────────────────────────────────── */
function _drawHeader(slide, ctx) {
  // Presentation title (left)
  _txt(slide, ctx.title || '',
       { x: 40, y: 0, w: 900, h: G.HEADER_H },
       { size: pt(15), color: C.text2, valign: 'middle' });

  // Page number (right, before logos)
  if (ctx.slideNum != null) {
    slide.addText([
      { text: String(ctx.slideNum).padStart(2,'0'), options: { bold: true, color: C.text1 } },
      { text: ctx.total ? ' / ' + String(ctx.total).padStart(2,'0') : '',
        options: { bold: false, color: C.muted } },
    ], {
      x: ix(1180), y: iy(0), w: ix(240), h: iy(G.HEADER_H),
      fontSize: pt(15), fontFace: 'Inter', align: 'right', valign: 'middle', margin: 0,
    });
  }

  // Logos (right edge, muted on light slides)
  // PNG dimensions from SVG viewBoxes: PNP = 400.8×75.4, Playbook = 168×40.3
  const pnpH = 20, pbH = 30;
  const pnpW = Math.round(pnpH * (400.8 / 75.4));  // ≈ 106px
  const pbW  = Math.round(pbH  * (168   / 40.3));  // ≈ 125px
  const RIGHT = 1880, GAP = 24;
  const pnpX = RIGHT - pnpW;

  const logos = ctx.logos || ['pnp'];
  if (logos.includes('playbook')) {
    const pbX = pnpX - GAP - pbW;
    const pbData = _cache.get(LOGO_URLS.playbook.muted);
    if (pbData) slide.addImage({ data: pbData,
      x: ix(pbX), y: iy((G.HEADER_H - pbH) / 2), w: ix(pbW), h: iy(pbH) });
  }
  const pnpData = _cache.get(LOGO_URLS.pnp.muted);
  if (pnpData) slide.addImage({ data: pnpData,
    x: ix(pnpX), y: iy((G.HEADER_H - pnpH) / 2), w: ix(pnpW), h: iy(pnpH) });
}

function _drawTitle(slide, title) {
  // Purple accent bar (pill shape)
  _rRect(slide, { x: 80, y: G.BAR_Y, w: G.BAR_W, h: G.BAR_H }, C.accent, { pill: true });
  // Slide title
  _txt(slide, title || '',
       { x: 80, y: G.TITLE_Y, w: 1760, h: G.TITLE_H },
       { size: pt(40), bold: true, color: C.text1, valign: 'top' });
}

/* ── Tag chip helper ─────────────────────────────────────────────────── */
function _tag(slide, label, anchor) {
  const w = Math.min(260, String(label).length * 8 + 28);
  _rRect(slide, { x: anchor.x, y: anchor.y, w, h: 28 }, C.accentChip, { pill: true });
  _txt(slide, label, { x: anchor.x, y: anchor.y, w, h: 28 },
       { size: pt(13), color: C.accent, align: 'center', valign: 'middle' });
  return w;
}

/* ══════════════════════════════════════════════════════════════════════
   CONTENT COMPONENTS
   Each function draws into `zone` (px on 1920×1080 grid).
   ══════════════════════════════════════════════════════════════════════ */

/* Card — white bg, 12px radius, 32px padding. title + body + optional tag.
   Matches .pb-card { padding:32px; gap:20px; border-radius:12px } */
function _drawCard(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const PAD = 32, GAP = 20;
  let cy = zone.y + PAD;
  const iw = zone.w - PAD * 2;

  if (data.tag) {
    _tag(slide, data.tag, { x: zone.x + PAD, y: cy });
    cy += 28 + GAP;
  }

  // Title (.pb-h3: 24px bold, line-height 1.2)
  const titleH = 64;  // 24px × 1.2lh × ~2 lines
  _txt(slide, data.title || '',
       { x: zone.x + PAD, y: cy, w: iw, h: titleH },
       { size: pt(24), bold: true, color: C.text1, valign: 'top' });
  cy += titleH;

  // Body (.pb-body: 18px, color text2)
  if (data.body) {
    _txt(slide, data.body,
         { x: zone.x + PAD, y: cy, w: iw, h: zone.y + zone.h - PAD - cy },
         { size: pt(18), color: C.text2, valign: 'top' });
  }
}

/* Metric (column-divider style).
   Matches .metrics__cell with num/value/label/sub stacked. */
function _drawMetric(slide, zone, data, small) {
  const PAD    = 32;
  const valSz  = small ? pt(72) : pt(96);
  const valH   = small ? 86    : 116;
  let cy = zone.y;

  // Number badge (01, 02…)
  if (data.number) {
    _txt(slide, data.number,
         { x: zone.x + PAD, y: cy, w: zone.w - PAD*2, h: 36 },
         { size: pt(18), bold: true, color: C.accent, tracking: 1, valign: 'top' });
    cy += 44;
  }

  // Big value
  _txt(slide, data.value || '',
       { x: zone.x + PAD, y: cy, w: zone.w - PAD*2, h: valH },
       { size: valSz, bold: true, color: C.text1, valign: 'top' });
  cy += valH + 8;

  // Label (.metrics__label: 22px)
  _txt(slide, data.label || '',
       { x: zone.x + PAD, y: cy, w: zone.w - PAD*2, h: 56 },
       { size: pt(22), color: C.text1, valign: 'top' });
  cy += 46;

  // Sub (.metrics__sub: 15px text2)
  if (data.sub) {
    _txt(slide, data.sub,
         { x: zone.x + PAD, y: cy, w: zone.w - PAD*2, h: 56 },
         { size: pt(15), color: C.text2, valign: 'top' });
  }
}

/* Takeaway — white card + purple left accent bar.
   Matches .takeaway: padding:36px 40px, bar 4px, gap:32px. */
function _drawTakeaway(slide, zone, item) {
  _rRect(slide, zone, C.card, { border: true });
  const padX = 40, padY = 36, barGap = 32;
  // Purple bar (full inner height)
  _rect(slide,
        { x: zone.x + padX, y: zone.y + padY, w: 4, h: zone.h - padY*2 },
        C.accent);
  const text = typeof item === 'string' ? item : (item.text || item.title || '');
  _txt(slide, text,
       { x: zone.x + padX + 4 + barGap, y: zone.y + padY,
         w: zone.w - padX*2 - 4 - barGap, h: zone.h - padY*2 },
       { size: pt(32), bold: true, color: C.text1, valign: 'middle' });
}

/* Bullet — purple left rule + main text + optional detail.
   Matches .bullets__item: rule 3px, gap:16px. */
function _drawBullet(slide, zone, item) {
  const text   = typeof item === 'string' ? item : (item.text || '');
  const detail = typeof item === 'object' ? item.detail : null;
  const RW = 3, gap = 16;

  _rect(slide, { x: zone.x, y: zone.y + 4, w: RW, h: zone.h - 8 }, C.accent);

  if (detail) {
    _txt(slide, text,
         { x: zone.x + RW + gap, y: zone.y, w: zone.w - RW - gap, h: zone.h * 0.55 },
         { size: pt(18), color: C.text1, valign: 'top' });
    _txt(slide, detail,
         { x: zone.x + RW + gap, y: zone.y + zone.h * 0.55,
           w: zone.w - RW - gap, h: zone.h * 0.45 },
         { size: pt(14), color: C.text2, valign: 'top' });
  } else {
    _txt(slide, text,
         { x: zone.x + RW + gap, y: zone.y, w: zone.w - RW - gap, h: zone.h },
         { size: pt(18), color: C.text1, valign: 'middle' });
  }
}

/* Statement — centered pull-quote in white card.
   Matches .statement: padding 80px 120px, quote 72px. */
function _drawStatement(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const padX = Math.min(120, zone.w * 0.12);
  const padY = Math.min(80,  zone.h * 0.12);
  const qSz  = zone.w < 1200 ? pt(56) : pt(72);
  const citeH = data.cite ? 44 : 0;

  _txt(slide, data.text || '',
       { x: zone.x + padX, y: zone.y + padY,
         w: zone.w - padX*2, h: zone.h - padY*2 - citeH },
       { size: qSz, bold: true, color: C.text1, align: 'center', valign: 'middle' });

  if (data.cite) {
    _txt(slide, data.cite,
         { x: zone.x + padX, y: zone.y + zone.h - padY - citeH,
           w: zone.w - padX*2, h: citeH },
         { size: pt(18), color: C.text2, align: 'center', valign: 'middle' });
  }
}

/* Row — full-width card: bold title left, body right, vertical rule between.
   Matches .row: grid 320px | 1fr, padding:32px 40px, gap:48px. */
function _drawRow(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const padX = 40, padY = 32;
  const titleW = 320, divGap = 48;

  _txt(slide, data.title || '',
       { x: zone.x + padX, y: zone.y + padY, w: titleW, h: zone.h - padY*2 },
       { size: pt(24), bold: true, color: C.text1, valign: 'middle' });

  // Vertical divider rule
  _rect(slide,
        { x: zone.x + padX + titleW + divGap/2, y: zone.y + padY, w: 1, h: zone.h - padY*2 },
        C.rule);

  _txt(slide, data.body || '',
       { x: zone.x + padX + titleW + divGap, y: zone.y + padY,
         w: zone.w - padX*2 - titleW - divGap, h: zone.h - padY*2 },
       { size: pt(18), color: C.text2, valign: 'middle' });
}

/* Agenda — time · title · duration with horizontal dividers.
   Matches .agenda__row: grid 180px | 1fr | 120px, gap:80px. */
function _drawAgenda(slide, zone, items) {
  const rows = stackZones(zone, items.length, 0);
  // Top border
  _rect(slide, { x: zone.x, y: zone.y, w: zone.w, h: 1 }, C.rule);

  items.forEach((it, i) => {
    const r = rows[i];
    // Bottom border for each row
    _rect(slide, { x: r.x, y: r.y + r.h, w: r.w, h: 1 }, C.rule);

    _txt(slide, it.time || '',
         { x: r.x, y: r.y, w: 180, h: r.h },
         { size: pt(28), bold: true, color: C.accent, valign: 'middle' });

    _txt(slide, it.title || '',
         { x: r.x + 260, y: r.y, w: r.w - 260 - (it.duration ? 120 : 0), h: r.h },
         { size: pt(28), bold: true, color: C.text1, valign: 'middle' });

    if (it.duration) {
      _txt(slide, it.duration,
           { x: r.x + r.w - 120, y: r.y, w: 120, h: r.h },
           { size: pt(18), color: C.text2, align: 'right', valign: 'middle' });
    }
  });
}

/* Contents / TOC — numbered rows, 1- or 2-column.
   Matches .contents: num 56px purple, title 28px, page label. */
function _drawContents(slide, zone, items, cols) {
  cols = cols || 2;
  const perCol  = Math.ceil(items.length / cols);
  const colGap  = 96;
  const colZones = gridZones(zone, cols, 1, colGap);

  items.forEach((it, i) => {
    const colIdx = Math.floor(i / perCol);
    const rowIdx = i % perCol;
    const col    = colZones[colIdx];
    const rows   = stackZones(col, perCol, 0);
    const r      = rows[rowIdx];

    if (rowIdx < perCol - 1)
      _rect(slide, { x: r.x, y: r.y + r.h - 1, w: r.w, h: 1 }, C.rule);

    _txt(slide, String(i + 1).padStart(2, '0'),
         { x: r.x, y: r.y, w: 80, h: r.h },
         { size: pt(56), bold: true, color: C.accent, valign: 'middle' });

    _txt(slide, it.title || '',
         { x: r.x + 112, y: r.y, w: r.w - 112 - (it.page ? 56 : 0), h: r.h },
         { size: pt(28), bold: true, color: C.text1, valign: 'middle' });

    if (it.page) {
      _txt(slide, String(it.page),
           { x: r.x + r.w - 56, y: r.y, w: 56, h: r.h },
           { size: pt(14), color: C.muted, align: 'right', valign: 'middle' });
    }
  });
}

/* ══════════════════════════════════════════════════════════════════════
   SLIDE BUILDERS
   ══════════════════════════════════════════════════════════════════════ */

function _buildCover(prs, data) {
  const slide = prs.addSlide();
  const t     = data.type;
  const dark  = t === 'cover-dark'   || t === 'cover-purple';
  const purple = t === 'cover-purple';
  const bg    = purple ? '3333CC' : (dark ? '161F41' : C.bg);
  const fg    = dark ? 'FFFFFF' : C.text1;
  const fg2   = dark ? 'C6CEDC' : C.text2;
  const bar   = dark ? (purple ? 'FFFFFF' : 'ACA5FA') : C.accent;

  slide.background = { color: bg };

  // cover-photo: solid left panel + photo right + gradient blend
  if (t === 'cover-photo') {
    const isDark  = data.variant !== 'light';
    const cvBg    = isDark ? '161F41' : C.bg;
    const cvFg    = isDark ? 'FFFFFF' : C.text1;
    const cvFg2   = isDark ? 'C6CEDC' : C.text2;
    slide.background = { color: cvBg };
    const PHOTO_X = 480;
    if (data.image_path) {
      _img(slide, data.image_path, { x: PHOTO_X, y: 0, w: 1920 - PHOTO_X, h: 1080 });
    }
    // Gradient blend
    slide.addImage({ data: _gradientPng(cvBg, 700),
      x: ix(PHOTO_X), y: 0, w: ix(700), h: SH });

    // Logos (top-left, colour variant on light, white on dark)
    const pnpH = 28, pbH = 34;
    const pnpW = Math.round(pnpH * (400.8 / 75.4));
    const pbW  = Math.round(pbH  * (168   / 40.3));
    const pnpVariant = isDark ? 'white' : 'color';
    const pbVariant  = isDark ? 'white' : 'color';
    const pnpD  = _cache.get(LOGO_URLS.pnp[pnpVariant]);
    const pbD   = _cache.get(LOGO_URLS.playbook[pbVariant]);
    if (pnpD) slide.addImage({ data: pnpD, x: ix(80), y: iy(80),       w: ix(pnpW), h: iy(pnpH) });
    if (pbD)  slide.addImage({ data: pbD,  x: ix(80 + pnpW + 32), y: iy(80), w: ix(pbW), h: iy(pbH) });

    // Accent bar + title + subtitle + date (same positions as cover-light)
    _rRect(slide, { x: 80, y: 290, w: 100, h: 8 }, C.accent, { pill: true });
    _txt(slide, data.title || '',   { x: 80, y: 320, w: 820, h: 460 },
         { size: pt(80), bold: true, color: cvFg, valign: 'top' });
    if (data.subtitle)
      _txt(slide, data.subtitle,   { x: 80, y: 800, w: 820, h: 60 },
           { size: pt(22), color: cvFg2, valign: 'top' });
    if (data.date)
      _txt(slide, data.date,       { x: 80, y: 980, w: 820, h: 40 },
           { size: pt(18), color: cvFg2 });
    return;
  }

  // cover-light / cover-dark / cover-purple
  const pnpH = 28, pnpW = Math.round(pnpH * (400.8/75.4));
  const pnpD = _cache.get(dark ? LOGO_URLS.pnp.white : LOGO_URLS.pnp.color);
  if (pnpD) slide.addImage({ data: pnpD, x: ix(80), y: iy(80), w: ix(pnpW), h: iy(pnpH) });

  _rRect(slide, { x: 80, y: 290, w: 100, h: 8 }, bar, { pill: true });
  _txt(slide, data.title || '',    { x: 80, y: 320, w: 800, h: 460 },
       { size: pt(80), bold: true, color: fg, valign: 'top' });
  if (data.subtitle)
    _txt(slide, data.subtitle,     { x: 80, y: 800, w: 800, h: 60  },
         { size: pt(22), color: fg2, valign: 'top' });
  if (data.date)
    _txt(slide, data.date,         { x: 80, y: 980, w: 800, h: 40  },
         { size: pt(18), color: fg2 });
}

function _buildSection(prs, data, ctx) {
  const slide = prs.addSlide();
  const isPurple = data.type === 'section-banner'
                || data.type === 'closing-purple';

  slide.background = { color: isPurple ? C.accent : C.bg };
  if (!isPurple) {
    _drawHeader(slide, ctx);
    _drawTitle(slide, data.slide_title || '');
  }
  const z = contentZone();

  if (data.type === 'section-number') {
    // Giant background number + section title below
    const numH = 260;
    const numY  = z.y + Math.max(0, (z.h - numH - 80) / 2);
    _txt(slide, data.number || '01',
         { x: z.x, y: numY, w: 820, h: numH },
         { size: pt(300), bold: true, color: C.accent, valign: 'top' });
    _txt(slide, data.title || '',
         { x: z.x, y: numY + numH + 8, w: z.w, h: 80 },
         { size: pt(56), bold: true, color: C.text1, valign: 'top' });
    if (data.subtitle)
      _txt(slide, data.subtitle,
           { x: z.x, y: z.y + z.h - 56, w: z.w, h: 56 },
           { size: pt(22), color: C.text2 });

  } else if (data.type === 'section-banner') {
    // Full accent-purple content zone with large white title
    _rRect(slide, z, C.accent, { r: ix(12) });
    _txt(slide, data.title || '', z,
         { size: pt(96), bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });

  } else if (data.type === 'closing-light' || data.type === 'closing-purple') {
    const fg  = isPurple ? 'FFFFFF' : C.text1;
    const fg2 = isPurple ? 'C6CEDC' : C.text2;
    const headH = 240, subH = 56, ctcH = 48, gap = 20;
    _txt(slide, data.heading || 'Thank You',
         { x: z.x, y: z.y, w: z.w, h: headH },
         { size: pt(160), bold: true, color: fg, valign: 'top' });
    if (data.subtitle)
      _txt(slide, data.subtitle,
           { x: z.x, y: z.y + headH + gap, w: z.w, h: subH },
           { size: pt(22), color: fg2 });
    if (data.contact)
      _txt(slide, data.contact,
           { x: z.x, y: z.y + z.h - ctcH, w: z.w, h: ctcH },
           { size: pt(22), color: isPurple ? 'FFFFFF' : C.accent });
  }
}

function _buildContent(prs, data, ctx) {
  const slide = prs.addSlide();
  slide.background = { color: C.bg };
  _drawHeader(slide, ctx);
  _drawTitle(slide, data.slide_title || '');
  const z = contentZone();
  const t = data.type;

  if (t === 'cards') {
    const items = data.items || data.columns || [];
    const cols  = data.cols || Math.min(items.length, 3);
    const rows  = Math.ceil(items.length / cols);
    const cells = gridZones(z, cols, rows, 24);
    items.forEach((c, i) => cells[i] && _drawCard(slide, cells[i], c));

  } else if (t === 'cards-2x2') {
    const items = (data.items || []).slice(0, 4);
    gridZones(z, 2, 2, 24).forEach((cell, i) => _drawCard(slide, cell, items[i] || {}));

  } else if (t === 'metrics-4') {
    _rect(slide, { x: z.x, y: z.y, w: z.w, h: 1 }, C.rule);
    const inner = { ...z, y: z.y + 40, h: z.h - 40 };
    const stats = (data.stats || []).slice(0, 4);
    const cells = gridZones(inner, stats.length, 1, 0);
    stats.forEach((s, i) => {
      _drawMetric(slide, cells[i], { number: String(i+1).padStart(2,'0'), ...s });
      if (i < stats.length - 1)
        _rect(slide, { x: cells[i].x + cells[i].w, y: inner.y, w: 1, h: inner.h }, C.rule);
    });

  } else if (t === 'metrics-8') {
    _rect(slide, { x: z.x, y: z.y, w: z.w, h: 1 }, C.rule);
    const inner = { ...z, y: z.y + 40, h: z.h - 40 };
    const stats = (data.stats || []).slice(0, 8);
    const cells = gridZones(inner, 4, 2, 0);
    stats.forEach((s, i) => {
      _drawMetric(slide, cells[i], { number: String(i+1).padStart(2,'0'), ...s }, true);
      if ((i+1) % 4 !== 0)
        _rect(slide, { x: cells[i].x + cells[i].w, y: cells[i].y, w: 1, h: cells[i].h }, C.rule);
    });
    _rect(slide, { x: inner.x, y: inner.y + inner.h / 2, w: inner.w, h: 1 }, C.rule);

  } else if (t === 'takeaways') {
    const items = data.items || [];
    stackZones(z, items.length).forEach((cell, i) => _drawTakeaway(slide, cell, items[i]));

  } else if (t === 'bullets') {
    const items = data.items || [];
    if (data.lead) {
      _txt(slide, data.lead,
           { x: z.x, y: z.y, w: Math.min(z.w, 560), h: 80 },
           { size: pt(18), color: C.text2, valign: 'top' });
    }
    const listZ = data.lead ? { ...z, y: z.y + 96, h: z.h - 96 } : z;
    stackZones(listZ, items.length).forEach((cell, i) => _drawBullet(slide, cell, items[i]));

  } else if (t === 'statement') {
    _drawStatement(slide, z, data);

  } else if (t === 'rows') {
    const items = data.items || [];
    stackZones(z, items.length).forEach((cell, i) => _drawRow(slide, cell, items[i]));

  } else if (t === 'agenda') {
    _drawAgenda(slide, z, data.items || []);

  } else if (t === 'contents') {
    _drawContents(slide, z, data.items || [], data.cols || 2);
  }
}

/* ══════════════════════════════════════════════════════════════════════
   ENTRY POINT
   buildPptx(slidesData, opts)
   opts: { title, logos, filename }
   ══════════════════════════════════════════════════════════════════════ */
async function buildPptx(slidesData, opts) {
  opts = opts || {};
  const setStatus = msg => {
    const el = document.getElementById('status');
    if (el) el.textContent = msg;
  };

  setStatus('Loading logos…');
  await _loadLogos();

  // Collect all image URLs referenced in slides and preload them
  const imgUrls = [];
  for (const s of slidesData) {
    if (s.image_path) imgUrls.push(s.image_path);
    if (s.bleed_image) imgUrls.push(s.bleed_image);
    if (s.items) s.items.forEach(it => { if (it && it.image_path) imgUrls.push(it.image_path); });
    if (s.columns) s.columns.forEach(c => { if (c && c.image_path) imgUrls.push(c.image_path); });
  }
  if (imgUrls.length) { setStatus('Loading images…'); await _preload(imgUrls); }

  setStatus('Building presentation…');
  const prs = new PptxGenJS();
  prs.defineLayout({ name: 'PNP', width: SW, height: SH });
  prs.layout = 'PNP';

  const title  = opts.title  || 'Presentation';
  const logos  = opts.logos  || ['pnp', 'playbook'];
  const total  = slidesData.length;

  const COVER_TYPES   = ['cover-light','cover-dark','cover-purple','cover-photo'];
  const SECTION_TYPES = ['section-number','section-banner','closing-light','closing-purple'];

  slidesData.forEach((s, i) => {
    const ctx = { title, logos, slideNum: i + 1, total };
    if (COVER_TYPES.includes(s.type))   _buildCover(prs, s);
    else if (SECTION_TYPES.includes(s.type)) _buildSection(prs, s, ctx);
    else                                _buildContent(prs, s, ctx);
  });

  setStatus('Saving…');
  const buf  = await prs.write('arraybuffer');
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href     = url;
  a.download = (opts.filename || title).replace(/[^a-z0-9\s\-_]/gi, '').replace(/\s+/g, '-') + '.pptx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus('✓ Saved to your downloads folder.');
}
