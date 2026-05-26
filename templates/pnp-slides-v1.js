/**
 * pnp-slides-v1.js  — rev 2
 * PNP / Playbook slide deck renderer — hosted on pnp-skill-assets CDN.
 * Entry point: buildPptx(slidesData, opts)
 *
 * Canvas: 20" × 11.25" = 1920×1080 at 96 DPI
 *   (matches Playbook Content Components.pptx exactly)
 * Coords: ix(px) = px / 96  →  80px = 0.8333"
 * Fonts:  all sizes are exact pt values extracted from the PPTX template
 *         (not derived from CSS px)
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
const SW = 20, SH = 11.25;          // 20" × 11.25" = 1920×1080 @ 96 DPI
const ix = n => +(n / 96).toFixed(4);  // CSS px → inches (96 DPI)
const iy = n => +(n / 96).toFixed(4);  // same DPI both axes

/* ── Typography — exact pt values from Playbook Content Components.pptx ─ */
// Extracted from slide XML (sz attribute ÷ 100).
const T = {
  header:      11,     // header meta text, sub-labels          sz=1100
  slideTitle:  25.2,   // slide title                           sz=2520
  metricNum:   13,     // metric row number "01"…"04"           sz=1300
  metricVal:   55.44,  // metric big KPI value                  sz=5544
  metricLabel: 16,     // metric label text                     sz=1600
  metricSub:   11,     // metric sub-label (muted)              sz=1100
  takeaway:    23.04,  // takeaway card text                    sz=2304
  cardTitle:   20,     // card heading (estimated from box h)
  cardBody:    13,     // card body text                        sz=1300
  bulletMain:  16,     // bullet main text
  bulletDet:   13,     // bullet detail text
  rowTitle:    20,     // row left title
  rowBody:     13,     // row right body
  agendaTime:  22,     // agenda time
  agendaTitle: 20,     // agenda session title
  agendaDur:   13,     // agenda duration
  tocNum:      36,     // TOC section number
  tocTitle:    20,     // TOC section title
  tocPage:     11,     // TOC page number (muted)
  statement:   28,     // pull-quote text
  statementCite: 13,   // pull-quote attribution
  coverTitle:  60,     // cover main heading
  coverSub:    16,     // cover subtitle
  coverDate:   13,     // cover date
  sectionNum:  120,    // large section display number
  sectionTitle: 36,    // section heading
  sectionSub:  16,     // section sub-heading
  closingHead: 80,     // "Thank You" heading
  closingSub:  20,     // closing subtitle
  closingCtc:  16,     // closing contact line
  tagLabel:    11,     // tag chip label
};

/* ── Design tokens ────────────────────────────────────────────────────── */
const C = {
  bg:         'F5F5F7',   // canvas background
  card:       'FFFFFF',
  cardBorder: 'E5E9F0',
  rule:       'C6CEDC',
  text1:      '1D2A57',   // oxford blue
  text2:      '44527A',   // dove gray
  muted:      '94A0B8',   // suva gray
  accent:     '3333CC',   // playbook purple
  accentChip: 'EBEBFA',   // accent bg for tags
  photoBg:    'D8DCEB',
  dark:       '161F41',
  navy:       '1D2A57',
};

/* ── Chrome geometry (px on 1920×1080 grid, verified from PPTX template) */
const G = {
  //  header area:  logos, page number, presentation name
  HEADER_Y:    40,    // header text top (logo center ≈ y=51)
  HEADER_H:   100,    // full header zone height
  //  accent bar:   40px × 4px pill, purple
  BAR_X:       80,
  BAR_Y:      102.5,  // verified: 976312 EMU → 102.49px
  BAR_W:       40,
  BAR_H:        4,
  //  slide title:  25.2pt bold Inter, oxford blue
  TITLE_X:     80,
  TITLE_Y:    120.5,  // verified: 1147762 EMU → 120.49px
  TITLE_H:     46,    // verified: 438151 EMU → 45.95px
  //  content zone
  CONTENT_X:   80,
  CONTENT_Y:  198.5,  // verified from takeaway/card slides
  CONTENT_B: 1000,    // 1080 − 80px bottom margin
};
G.CONTENT_H = G.CONTENT_B - G.CONTENT_Y;   // 801.5 px
G.CONTENT_W = 1920 - 2 * G.CONTENT_X;      // 1760 px

/* ── Card corner radius ───────────────────────────────────────────────── */
// Template: roundRect adj=4778 on a card h≈251px → cornerRadius = 12px
// At 96 DPI: 12px / 96 = 0.125" exactly.
const CARD_R = 12 / 96;   // 0.125" — matches CSS border-radius:12px

/* ── Zone helpers ─────────────────────────────────────────────────────── */
function contentZone() {
  return { x: G.CONTENT_X, y: G.CONTENT_Y, w: G.CONTENT_W, h: G.CONTENT_H };
}

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

function stackZones(zone, n, gap) {
  return gridZones(zone, 1, n, gap ?? 24);
}

/* ── Drawing primitives ───────────────────────────────────────────────── */
function _txt(slide, text, zone, opts) {
  opts = opts || {};
  const params = {
    x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h),
    fontSize:  opts.size  || T.body || 13,
    bold:      opts.bold  || false,
    color:     opts.color || C.text1,
    fontFace:  'Inter',
    align:     opts.align  || 'left',
    valign:    opts.valign || 'middle',
    wrap:      opts.wrap !== false,
    margin:    0,
  };
  if (opts.lineSpacing) params.lineSpacingMultiple = opts.lineSpacing;
  if (opts.charSpacing) params.charSpacing = opts.charSpacing;
  slide.addText(text || '', params);
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

// Rounded rect. CARD_R (0.125") is the standard 12px card corner.
// pill:true uses rectRadius:0.5 (PptxGenJS caps, gives full pill).
function _rRect(slide, zone, fill, opts) {
  opts = opts || {};
  slide.addShape('roundRect', {
    x: ix(zone.x), y: iy(zone.y), w: ix(zone.w), h: iy(zone.h),
    fill: { color: fill },
    line: opts.border
      ? { color: opts.borderColor || C.cardBorder, width: 1 }
      : { type: 'none' },
    rectRadius: opts.pill ? 0.5 : (opts.r != null ? opts.r : CARD_R),
  });
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

/* ── Gradient overlay helper (cover-photo blend) ─────────────────────── */
function _gradientPng(hex, w) {
  const r = parseInt(hex.slice(0,2),16), g = parseInt(hex.slice(2,4),16), b = parseInt(hex.slice(4,6),16);
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = 2;
  const ctx = cv.getContext('2d');
  const grad = ctx.createLinearGradient(0,0,w,0);
  grad.addColorStop(0, `rgba(${r},${g},${b},1)`);
  grad.addColorStop(1, `rgba(${r},${g},${b},0)`);
  ctx.fillStyle = grad; ctx.fillRect(0,0,w,2);
  return cv.toDataURL('image/png');
}

/* ══════════════════════════════════════════════════════════════════════
   CHROME HELPERS
   ══════════════════════════════════════════════════════════════════════ */

// Logo display sizes (px at 96 DPI, from PPTX template measurements)
// Template: PNP=72.3×20, Playbook=84.5×20 (both at 55% muted opacity in header)
const LOGO = {
  pnpW: 72, pnpH: 20,
  pbW:  85, pbH:  20,
  gap:  10,
};

function _drawHeader(slide, ctx) {
  // Presentation title (left, 11pt muted)
  _txt(slide, ctx.title || '',
       { x: 40, y: 0, w: 900, h: G.HEADER_H },
       { size: T.header, color: C.text2, valign: 'middle' });

  // Page number (right, before logos)
  if (ctx.slideNum != null) {
    const numRight = 1920 - LOGO.pbW - LOGO.gap - LOGO.pnpW - LOGO.gap - 80;
    slide.addText([
      { text: String(ctx.slideNum).padStart(2,'0'),
        options: { bold: true, color: C.text1 } },
      { text: ctx.total ? ' / ' + String(ctx.total).padStart(2,'0') : '',
        options: { bold: false, color: C.muted } },
    ], {
      x: ix(numRight - 80), y: iy(0), w: ix(80), h: iy(G.HEADER_H),
      fontSize: T.header, fontFace: 'Inter', align: 'right', valign: 'middle', margin: 0,
    });
  }

  // Logos right-aligned (muted variant in chrome)
  const logos = ctx.logos || ['pnp'];
  const RIGHT = 1880;
  let logoX = RIGHT;

  if (logos.includes('pnp')) {
    logoX -= LOGO.pnpW;
    const d = _cache.get(LOGO_URLS.pnp.muted);
    if (d) slide.addImage({ data: d,
      x: ix(logoX), y: iy((G.HEADER_H - LOGO.pnpH) / 2),
      w: ix(LOGO.pnpW), h: iy(LOGO.pnpH) });
    logoX -= LOGO.gap;
  }
  if (logos.includes('playbook')) {
    logoX -= LOGO.pbW;
    const d = _cache.get(LOGO_URLS.playbook.muted);
    if (d) slide.addImage({ data: d,
      x: ix(logoX), y: iy((G.HEADER_H - LOGO.pbH) / 2),
      w: ix(LOGO.pbW), h: iy(LOGO.pbH) });
  }
}

function _drawTitle(slide, title) {
  // Accent bar — pill-shaped roundRect (template: prst="roundRect" adj=50000)
  _rRect(slide, { x: G.BAR_X, y: G.BAR_Y, w: G.BAR_W, h: G.BAR_H }, C.accent, { pill: true });
  // Slide title — 25.2pt bold Inter, oxford blue
  _txt(slide, title || '',
       { x: G.TITLE_X, y: G.TITLE_Y, w: G.CONTENT_W, h: G.TITLE_H },
       { size: T.slideTitle, bold: true, color: C.text1, valign: 'top' });
}

/* ── Tag chip ─────────────────────────────────────────────────────────── */
function _tag(slide, label, anchor) {
  const w = Math.min(240, String(label).length * 7 + 24);
  _rRect(slide, { x: anchor.x, y: anchor.y, w, h: 26 }, C.accentChip, { pill: true });
  _txt(slide, label, { x: anchor.x, y: anchor.y, w, h: 26 },
       { size: T.tagLabel, color: C.accent, align: 'center', valign: 'middle' });
  return w;
}

/* ══════════════════════════════════════════════════════════════════════
   CONTENT COMPONENTS
   ══════════════════════════════════════════════════════════════════════ */

/* Card — white rounded rect, title + body + optional tag chip.
   Template: roundRect adj=3087/4778 → 12px corner. Padding ≈ 32px.
   Font: title=20pt bold, body=13pt medium. */
function _drawCard(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const PAD = 32, GAP = 16;
  let cy = zone.y + PAD;
  const iw = zone.w - PAD * 2;

  if (data.tag) {
    _tag(slide, data.tag, { x: zone.x + PAD, y: cy });
    cy += 26 + GAP;
  }

  const titleH = Math.min(80, zone.h * 0.3);
  _txt(slide, data.title || '',
       { x: zone.x + PAD, y: cy, w: iw, h: titleH },
       { size: T.cardTitle, bold: true, color: C.text1, valign: 'top' });
  cy += titleH + GAP * 0.5;

  if (data.body) {
    _txt(slide, data.body,
         { x: zone.x + PAD, y: cy, w: iw, h: zone.y + zone.h - PAD - cy },
         { size: T.cardBody, color: C.text2, valign: 'top', lineSpacing: 1.5 });
  }
}

/* Metrics — four columns divided by 1px rules.
   Template layout (verified from slide21.xml):
     - Horizontal rule at y=406px (breathing space below title)
     - Content starts at y=448px (41px below rule)
     - Within each column: num(13pt) | gap12 | value(55.44pt) | gap12 | label(16pt) | gap12 | sub(11pt)
     - Vertical dividers at mid-gap between columns
   small=true → use smaller value size (for metrics-8 2-row grid). */
function _drawMetric(slide, zone, data, small) {
  const valSz = small ? 36 : T.metricVal;
  let y = zone.y;

  // Number badge "01"…
  if (data.number) {
    _txt(slide, data.number,
         { x: zone.x, y, w: zone.w, h: 31 },
         { size: T.metricNum, bold: true, color: C.accent });
    y += 43;   // gap between num bottom and value top (verified: 490.7−447.7)
  }

  // Big value
  const valH = small ? 72 : 95;
  _txt(slide, data.value || '',
       { x: zone.x, y, w: zone.w, h: valH },
       { size: valSz, bold: true, color: C.text1, valign: 'top' });
  y += valH + 12;   // verified gap before label (597.9−585.9)

  // Label
  _txt(slide, data.label || '',
       { x: zone.x, y, w: zone.w, h: 37 },
       { size: T.metricLabel, color: C.text1, valign: 'top' });
  y += 37 + 12;   // verified gap before sub (646.9−633.9)

  // Sub-label
  if (data.sub) {
    _txt(slide, data.sub,
         { x: zone.x, y, w: zone.w, h: 27 },
         { size: T.metricSub, color: C.text2, valign: 'top' });
  }
}

/* Takeaway — white card + purple 4px vertical pill bar on left.
   Template (slide28): card h=251px, bar 4×177px, text 23.04pt bold.
   Bar padding top/bottom = 37px. Text left offset = 77px from card. */
function _drawTakeaway(slide, zone, item) {
  // Card background
  _rRect(slide, zone, C.card, { border: true });

  // Purple left bar — pill shape, 4px wide, inset 40px from card left
  const BAR_INSET = 41;   // card_left → bar_left (template: 120.6−80)
  const BAR_PAD   = 37;   // top/bottom padding inside card
  const BAR_W     = 4;
  _rRect(slide,
    { x: zone.x + BAR_INSET, y: zone.y + BAR_PAD,
      w: BAR_W, h: zone.h - BAR_PAD * 2 },
    C.accent, { pill: true });

  // Text — starts 77px from card left, centered vertically
  const TEXT_X = 77;    // template: 156.5−80
  const TEXT_PAD = 32;  // vertical padding for text box
  const text = typeof item === 'string' ? item : (item.text || item.title || '');
  _txt(slide, text,
       { x: zone.x + TEXT_X, y: zone.y + TEXT_PAD,
         w: zone.w - TEXT_X - 32, h: zone.h - TEXT_PAD * 2 },
       { size: T.takeaway, bold: true, color: C.text1, valign: 'middle', lineSpacing: 1.2 });
}

/* Bullet — purple 3px left rule, main text + optional detail.  */
function _drawBullet(slide, zone, item) {
  const text   = typeof item === 'string' ? item : (item.text || '');
  const detail = typeof item === 'object' ? item.detail : null;
  const RW = 3, GAP = 16;

  _rect(slide, { x: zone.x, y: zone.y + 4, w: RW, h: zone.h - 8 }, C.accent);

  if (detail) {
    _txt(slide, text,
         { x: zone.x + RW + GAP, y: zone.y, w: zone.w - RW - GAP, h: zone.h * 0.55 },
         { size: T.bulletMain, bold: true, color: C.text1, valign: 'top' });
    _txt(slide, detail,
         { x: zone.x + RW + GAP, y: zone.y + zone.h * 0.55,
           w: zone.w - RW - GAP, h: zone.h * 0.45 },
         { size: T.bulletDet, color: C.text2, valign: 'top' });
  } else {
    _txt(slide, text,
         { x: zone.x + RW + GAP, y: zone.y, w: zone.w - RW - GAP, h: zone.h },
         { size: T.bulletMain, bold: true, color: C.text1, valign: 'middle' });
  }
}

/* Statement — centered pull-quote in white card.  */
function _drawStatement(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const padX = Math.min(100, zone.w * 0.08);
  const padY = Math.min(60,  zone.h * 0.10);
  const citeH = data.cite ? 36 : 0;

  _txt(slide, data.text || '',
       { x: zone.x + padX, y: zone.y + padY,
         w: zone.w - padX * 2, h: zone.h - padY * 2 - citeH },
       { size: T.statement, bold: true, color: C.text1, align: 'center', valign: 'middle',
         lineSpacing: 1.3 });

  if (data.cite) {
    _txt(slide, data.cite,
         { x: zone.x + padX, y: zone.y + zone.h - padY - citeH,
           w: zone.w - padX * 2, h: citeH },
         { size: T.statementCite, color: C.text2, align: 'center' });
  }
}

/* Row — full-width card, title left | body right.  */
function _drawRow(slide, zone, data) {
  _rRect(slide, zone, C.card, { border: true });
  const padX = 40, padY = 28;
  const TITLE_W = 300, DIV_GAP = 48;

  _txt(slide, data.title || '',
       { x: zone.x + padX, y: zone.y + padY, w: TITLE_W, h: zone.h - padY * 2 },
       { size: T.rowTitle, bold: true, color: C.text1, valign: 'middle' });

  _rect(slide,
        { x: zone.x + padX + TITLE_W + DIV_GAP / 2, y: zone.y + padY,
          w: 1, h: zone.h - padY * 2 }, C.rule);

  _txt(slide, data.body || '',
       { x: zone.x + padX + TITLE_W + DIV_GAP, y: zone.y + padY,
         w: zone.w - padX * 2 - TITLE_W - DIV_GAP, h: zone.h - padY * 2 },
       { size: T.rowBody, color: C.text2, valign: 'middle', lineSpacing: 1.5 });
}

/* Agenda — time · title · duration with dividers.  */
function _drawAgenda(slide, zone, items) {
  const rows = stackZones(zone, items.length, 0);
  _rect(slide, { x: zone.x, y: zone.y, w: zone.w, h: 1 }, C.rule);

  items.forEach((it, i) => {
    const r = rows[i];
    _rect(slide, { x: r.x, y: r.y + r.h, w: r.w, h: 1 }, C.rule);

    _txt(slide, it.time || '',
         { x: r.x, y: r.y, w: 160, h: r.h },
         { size: T.agendaTime, bold: true, color: C.accent, valign: 'middle' });

    _txt(slide, it.title || '',
         { x: r.x + 220, y: r.y,
           w: r.w - 220 - (it.duration ? 140 : 0), h: r.h },
         { size: T.agendaTitle, bold: true, color: C.text1, valign: 'middle' });

    if (it.duration) {
      _txt(slide, it.duration,
           { x: r.x + r.w - 140, y: r.y, w: 140, h: r.h },
           { size: T.agendaDur, color: C.text2, align: 'right', valign: 'middle' });
    }
  });
}

/* Contents / TOC — numbered rows, 1- or 2-column.  */
function _drawContents(slide, zone, items, cols) {
  cols = cols || 2;
  const perCol   = Math.ceil(items.length / cols);
  const colZones = gridZones(zone, cols, 1, 80);

  items.forEach((it, i) => {
    const colIdx = Math.floor(i / perCol);
    const rowIdx = i % perCol;
    const col    = colZones[Math.min(colIdx, cols - 1)];
    const rows   = stackZones(col, perCol, 0);
    const r      = rows[rowIdx] || rows[rows.length - 1];

    if (rowIdx < perCol - 1)
      _rect(slide, { x: r.x, y: r.y + r.h - 1, w: r.w, h: 1 }, C.rule);

    _txt(slide, String(i + 1).padStart(2, '0'),
         { x: r.x, y: r.y, w: 64, h: r.h },
         { size: T.tocNum, bold: true, color: C.accent, valign: 'middle' });

    _txt(slide, it.title || '',
         { x: r.x + 88, y: r.y, w: r.w - 88 - (it.page ? 48 : 0), h: r.h },
         { size: T.tocTitle, bold: true, color: C.text1, valign: 'middle' });

    if (it.page) {
      _txt(slide, String(it.page),
           { x: r.x + r.w - 48, y: r.y, w: 48, h: r.h },
           { size: T.tocPage, color: C.muted, align: 'right', valign: 'middle' });
    }
  });
}

/* ══════════════════════════════════════════════════════════════════════
   SLIDE BUILDERS
   ══════════════════════════════════════════════════════════════════════ */

function _buildCover(prs, data) {
  const slide = prs.addSlide();
  const t = data.type;

  /* cover-photo — solid left panel + photo right + gradient blend */
  if (t === 'cover-photo') {
    const isDark = data.variant !== 'light';
    const cvBg   = isDark ? C.dark : C.bg;
    const cvFg   = isDark ? 'FFFFFF' : C.text1;
    const cvFg2  = isDark ? 'C6CEDC' : C.text2;
    slide.background = { color: cvBg };

    const PHOTO_X = 480;
    if (data.image_path)
      _img(slide, data.image_path, { x: PHOTO_X, y: 0, w: 1920 - PHOTO_X, h: 1080 });

    // Gradient blend from solid to transparent
    slide.addImage({ data: _gradientPng(cvBg, 700),
      x: ix(PHOTO_X), y: 0, w: ix(700), h: SH });

    // Logos top-left
    const pnpVar = isDark ? 'white' : 'color';
    const pbVar  = isDark ? 'white' : 'color';
    const pnpH2 = 22, pnpW2 = Math.round(pnpH2 * (LOGO.pnpW / LOGO.pnpH));
    const pbH2  = 22, pbW2  = Math.round(pbH2  * (LOGO.pbW  / LOGO.pbH));
    const pnpD  = _cache.get(LOGO_URLS.pnp[pnpVar]);
    const pbD   = _cache.get(LOGO_URLS.playbook[pbVar]);
    if (pnpD) slide.addImage({ data: pnpD, x: ix(80), y: iy(72), w: ix(pnpW2), h: iy(pnpH2) });
    if (pbD)  slide.addImage({ data: pbD,  x: ix(80 + pnpW2 + 24), y: iy(72), w: ix(pbW2), h: iy(pbH2) });

    // Accent bar + copy
    _rRect(slide, { x: 80, y: 290, w: 80, h: 6 }, C.accent, { pill: true });
    _txt(slide, data.title || '',   { x: 80, y: 310, w: 780, h: 440 },
         { size: T.coverTitle, bold: true, color: cvFg, valign: 'top', lineSpacing: 1.1 });
    if (data.subtitle)
      _txt(slide, data.subtitle,   { x: 80, y: 770, w: 780, h: 60 },
           { size: T.coverSub, color: cvFg2, valign: 'top' });
    if (data.date)
      _txt(slide, data.date,       { x: 80, y: 980, w: 780, h: 36 },
           { size: T.coverDate, color: cvFg2 });
    return;
  }

  /* cover-light / cover-dark / cover-purple */
  const dark   = t === 'cover-dark' || t === 'cover-purple';
  const purple = t === 'cover-purple';
  const bg     = purple ? C.accent : (dark ? C.dark : C.bg);
  const fg     = dark ? 'FFFFFF' : C.text1;
  const fg2    = dark ? 'C6CEDC' : C.text2;
  const bar    = dark ? (purple ? 'FFFFFF' : 'ACA5FA') : C.accent;

  slide.background = { color: bg };
  const pnpVar2 = dark ? 'white' : 'color';
  const pnpH3 = 22, pnpW3 = Math.round(pnpH3 * (LOGO.pnpW / LOGO.pnpH));
  const pnpD2 = _cache.get(LOGO_URLS.pnp[pnpVar2]);
  if (pnpD2) slide.addImage({ data: pnpD2, x: ix(80), y: iy(72), w: ix(pnpW3), h: iy(pnpH3) });

  _rRect(slide, { x: 80, y: 290, w: 80, h: 6 }, bar, { pill: true });
  _txt(slide, data.title || '',   { x: 80, y: 310, w: 800, h: 440 },
       { size: T.coverTitle, bold: true, color: fg, valign: 'top', lineSpacing: 1.1 });
  if (data.subtitle)
    _txt(slide, data.subtitle,   { x: 80, y: 770, w: 800, h: 60 },
         { size: T.coverSub, color: fg2, valign: 'top' });
  if (data.date)
    _txt(slide, data.date,       { x: 80, y: 980, w: 800, h: 36 },
         { size: T.coverDate, color: fg2 });
}

function _buildSection(prs, data, ctx) {
  const slide = prs.addSlide();
  const isPurple = data.type === 'section-banner' || data.type === 'closing-purple';
  slide.background = { color: isPurple ? C.accent : C.bg };

  if (!isPurple) {
    _drawHeader(slide, ctx);
    _drawTitle(slide, data.slide_title || '');
  }
  const z = contentZone();

  if (data.type === 'section-number') {
    // Giant number left + section title + subtitle
    const numH = 220;
    const numY  = z.y + Math.max(0, (z.h - numH - 100) / 2);
    _txt(slide, data.number || '01',
         { x: z.x, y: numY, w: 600, h: numH },
         { size: T.sectionNum, bold: true, color: C.accent, valign: 'top' });
    _txt(slide, data.title || '',
         { x: z.x, y: numY + numH + 16, w: z.w, h: 80 },
         { size: T.sectionTitle, bold: true, color: C.text1, valign: 'top' });
    if (data.subtitle)
      _txt(slide, data.subtitle,
           { x: z.x, y: numY + numH + 16 + 80 + 12, w: z.w, h: 48 },
           { size: T.sectionSub, color: C.text2 });

  } else if (data.type === 'section-banner') {
    _rRect(slide, z, C.accent, { r: CARD_R });
    _txt(slide, data.title || '', z,
         { size: T.sectionTitle * 2, bold: true, color: 'FFFFFF',
           align: 'center', valign: 'middle' });

  } else if (data.type === 'closing-light' || data.type === 'closing-purple') {
    const fg  = isPurple ? 'FFFFFF' : C.text1;
    const fg2 = isPurple ? 'C6CEDC' : C.text2;
    _txt(slide, data.heading || 'Thank You',
         { x: z.x, y: z.y, w: z.w, h: 200 },
         { size: T.closingHead, bold: true, color: fg, valign: 'top' });
    if (data.subtitle)
      _txt(slide, data.subtitle,
           { x: z.x, y: z.y + 220, w: z.w, h: 48 },
           { size: T.closingSub, color: fg2 });
    if (data.contact)
      _txt(slide, data.contact,
           { x: z.x, y: z.y + z.h - 48, w: z.w, h: 48 },
           { size: T.closingCtc, color: isPurple ? 'FFFFFF' : C.accent });
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
    const items = data.items || [];
    const cols  = data.cols || Math.min(items.length, 3);
    const rows  = Math.ceil(items.length / cols);
    gridZones(z, cols, rows, 24).forEach((cell, i) => items[i] && _drawCard(slide, cell, items[i]));

  } else if (t === 'cards-2x2') {
    const items = (data.items || []).slice(0, 4);
    gridZones(z, 2, 2, 24).forEach((cell, i) => _drawCard(slide, cell, items[i] || {}));

  } else if (t === 'metrics-4') {
    // Horizontal rule at y=406 (verified from template — breathing space under title)
    const RULE_Y   = 406;
    const METRIC_Y = 448;   // 41px below rule (template: 447.7px)
    const DIV_GAP  = 53;    // gap between metric columns (template: ~53px)

    _rect(slide, { x: z.x, y: RULE_Y, w: z.w, h: 1 }, C.rule);

    const stats  = (data.stats || []).slice(0, 4);
    const mZone  = { x: z.x, y: METRIC_Y, w: z.w, h: z.y + z.h - METRIC_Y };
    const cells  = gridZones(mZone, stats.length, 1, DIV_GAP);

    stats.forEach((s, i) => {
      _drawMetric(slide, cells[i], { number: String(i+1).padStart(2,'0'), ...s });
      if (i < stats.length - 1) {
        // Vertical divider at center of gap
        const divX = cells[i].x + cells[i].w + Math.round(DIV_GAP / 2);
        _rect(slide, { x: divX, y: METRIC_Y, w: 1, h: 222 }, C.rule);
      }
    });

  } else if (t === 'metrics-8') {
    const RULE_Y   = 406;
    const METRIC_Y = 448;
    const DIV_GAP  = 53;

    _rect(slide, { x: z.x, y: RULE_Y, w: z.w, h: 1 }, C.rule);

    const stats = (data.stats || []).slice(0, 8);
    const totalH = z.y + z.h - METRIC_Y;
    const rowH   = (totalH - 24) / 2;  // 2 rows with 24px gap
    const mZone1 = { x: z.x, y: METRIC_Y,           w: z.w, h: rowH };
    const mZone2 = { x: z.x, y: METRIC_Y + rowH + 24, w: z.w, h: rowH };
    const cells1 = gridZones(mZone1, 4, 1, DIV_GAP);
    const cells2 = gridZones(mZone2, 4, 1, DIV_GAP);

    // Row separator
    _rect(slide, { x: z.x, y: METRIC_Y + rowH + 12, w: z.w, h: 1 }, C.rule);

    stats.slice(0,4).forEach((s, i) => {
      _drawMetric(slide, cells1[i], { number: String(i+1).padStart(2,'0'), ...s }, true);
      if (i < 3) {
        const dx = cells1[i].x + cells1[i].w + Math.round(DIV_GAP/2);
        _rect(slide, { x: dx, y: METRIC_Y, w: 1, h: rowH }, C.rule);
      }
    });
    stats.slice(4).forEach((s, i) => {
      _drawMetric(slide, cells2[i], { number: String(i+5).padStart(2,'0'), ...s }, true);
      if (i < 3) {
        const dx = cells2[i].x + cells2[i].w + Math.round(DIV_GAP/2);
        _rect(slide, { x: dx, y: METRIC_Y + rowH + 24, w: 1, h: rowH }, C.rule);
      }
    });

  } else if (t === 'takeaways') {
    const items = data.items || [];
    stackZones(z, items.length).forEach((cell, i) => _drawTakeaway(slide, cell, items[i]));

  } else if (t === 'bullets') {
    const items = data.items || [];
    let listZ = z;
    if (data.lead) {
      _txt(slide, data.lead,
           { x: z.x, y: z.y, w: Math.min(z.w, 960), h: 72 },
           { size: T.cardBody, color: C.text2, valign: 'top' });
      listZ = { ...z, y: z.y + 88, h: z.h - 88 };
    }
    stackZones(listZ, items.length).forEach((cell, i) => _drawBullet(slide, cell, items[i]));

  } else if (t === 'statement') {
    _drawStatement(slide, z, data);

  } else if (t === 'rows') {
    stackZones(z, (data.items || []).length)
      .forEach((cell, i) => _drawRow(slide, cell, (data.items || [])[i]));

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
  const setStatus = msg => { const el = document.getElementById('status'); if (el) el.textContent = msg; };

  setStatus('Loading logos…');
  await _loadLogos();

  // Pre-load all referenced images
  const imgUrls = [];
  for (const s of slidesData) {
    if (s.image_path) imgUrls.push(s.image_path);
    if (s.bleed_image) imgUrls.push(s.bleed_image);
    (s.items || []).forEach(it => { if (it && it.image_path) imgUrls.push(it.image_path); });
  }
  if (imgUrls.length) { setStatus('Loading images…'); await _preload(imgUrls); }

  setStatus('Building presentation…');
  const prs = new PptxGenJS();
  // Match the exact canvas of Playbook Content Components.pptx
  prs.defineLayout({ name: 'PB_1920', width: SW, height: SH });
  prs.layout = 'PB_1920';

  const title  = opts.title || 'Presentation';
  const logos  = opts.logos || ['pnp', 'playbook'];
  const total  = slidesData.length;

  const COVERS   = ['cover-light','cover-dark','cover-purple','cover-photo'];
  const SECTIONS = ['section-number','section-banner','closing-light','closing-purple'];

  slidesData.forEach((s, i) => {
    const ctx = { title, logos, slideNum: i + 1, total };
    if (COVERS.includes(s.type))         _buildCover(prs, s);
    else if (SECTIONS.includes(s.type))  _buildSection(prs, s, ctx);
    else                                 _buildContent(prs, s, ctx);
  });

  setStatus('Saving…');
  const buf  = await prs.write('arraybuffer');
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href     = url;
  a.download = (opts.filename || title).replace(/[^a-z0-9\s\-_]/gi,'').replace(/\s+/g,'-') + '.pptx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus('✓ Saved to your downloads folder.');
}
