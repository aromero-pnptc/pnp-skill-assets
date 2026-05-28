/*!
 * PnP Presentation Engine v1.0
 * Requires: PptxGenJS 3.x loaded as global before this script
 *
 * Usage:
 *   window.PNP_DECK = { title, theme, slides: [...] };
 *   // Engine auto-runs on load if PNP_DECK is defined
 */
(function () {
  'use strict';

  // Canvas: LAYOUT_WIDE = 13.333" × 7.5"
  const W     = 13.333;
  const H     = 7.5;
  const PAD   = 0.55;          // left/right padding
  const CW    = W - PAD * 2;   // content width
  const HDR_H = 0.52;          // header bar height
  const TTL_Y = 0.54;
  const TTL_H = 0.38;
  const CON_Y = 0.96;          // content area top
  const CON_H = 5.76;          // content area height
  const GAP   = 0.14;          // gap between cards

  const GH = 'https://raw.githubusercontent.com/aromero-pnptc/pnp-skill-assets/main/pnp-presentations';

  const THEME = {
    light: {
      bg: 'F5F5F7', card: 'FFFFFF', border: 'C6CEDC',
      t1: '1D2A57', t2: '44527A', accent: '3333CC',
      logo: GH + '/logos/logo-pnp-color.png',
    },
    dark: {
      bg: '161F41', card: '1D2A57', border: '5D688E',
      t1: 'FFFFFF', t2: 'C6CEDC', accent: 'ACA5FA',
      logo: GH + '/logos/logo-pnp-white.png',
    },
  };

  const PLAYBOOK_LOGO = {
    light: GH + '/logos/logo-playbook-color.png',
    dark:  GH + '/logos/logo-playbook-white.png',
  };

  // ── Low-level helpers ────────────────────────────────────────────────────

  function box(slide, x, y, w, h, fill, lineColor, lineW) {
    slide.addShape('rect', {
      x, y, w, h,
      fill: { color: fill },
      line: lineColor ? { color: lineColor, width: lineW || 0.75 } : { color: fill, width: 0 },
    });
  }

  function txt(slide, text, x, y, w, h, opts) {
    if (text == null || text === '') return;
    slide.addText(String(text), Object.assign(
      { x, y, w, h, fontFace: 'Arial', valign: 'middle', wrap: true, charSpacing: 0 },
      opts
    ));
  }

  function photo(slide, path, x, y, w, h) {
    if (!path) return;
    try {
      slide.addImage({ path, x, y, w, h, sizing: { type: 'cover', w, h } });
    } catch (e) {
      // If image fails (CORS/timeout), skip silently
    }
  }

  // ── Shared components ────────────────────────────────────────────────────

  function hdr(slide, C, presTitle, logos) {
    // Accent bar
    box(slide, 0, 0, W, HDR_H, C.accent);

    // Logo(s)
    const logoList = logos || ['pnp'];
    let lx = PAD * 0.5;
    logoList.forEach(l => {
      const src = l === 'playbook'
        ? PLAYBOOK_LOGO[C === THEME.dark ? 'dark' : 'light']
        : C.logo;
      try { slide.addImage({ path: src, x: lx, y: 0.10, w: 1.1, h: 0.32 }); } catch (e) {}
      lx += 1.25;
    });

    // Presentation title (right)
    if (presTitle) {
      txt(slide, presTitle, W - PAD - 3.8, 0.1, 3.8, 0.32,
        { fontSize: 9, color: 'FFFFFF', align: 'right', valign: 'middle', wrap: false });
    }
  }

  function heading(slide, title, C) {
    if (!title) return;
    txt(slide, title, PAD, TTL_Y, CW, TTL_H,
      { fontSize: 18, bold: true, color: C.t1, valign: 'middle' });
  }

  // Card: white box with optional top photo
  function card(slide, x, y, w, h, col, C, withImg) {
    box(slide, x, y, w, h, C.card, C.border);
    const imgH  = withImg && col.image ? h * 0.44 : 0;
    const txtH  = h - imgH;
    const p     = 0.17;
    let ty = y + p;

    if (col.title) {
      txt(slide, col.title, x + p, ty, w - p * 2, 0.28,
        { fontSize: 13, bold: true, color: C.t1, valign: 'top' });
      ty += 0.31;
    }
    if (col.body) {
      txt(slide, col.body, x + p, ty, w - p * 2, txtH - (ty - y) - p,
        { fontSize: 10.5, color: C.t2, valign: 'top' });
    }
    if (imgH && col.image) {
      photo(slide, col.image, x + 0.01, y + txtH, w - 0.02, imgH - 0.01);
    }
  }

  // ── Slide builders ───────────────────────────────────────────────────────

  // COVER
  function buildCover(slide, d, C) {
    const dark   = d.type === 'cover-dark';
    const bgCol  = dark ? '161F41' : 'F5F5F7';
    const textC  = dark ? 'FFFFFF' : C.t1;
    const subC   = dark ? 'C6CEDC' : C.t2;

    slide.background = { color: bgCol };

    if (d.image) {
      photo(slide, d.image, W * 0.47, 0, W * 0.53, H);
      // Mask left panel so text is readable
      box(slide, 0, 0, W * 0.50, H, bgCol);
    }

    // Left accent stripe
    box(slide, 0, 0, 0.12, H, C.accent);

    // Logo
    try { slide.addImage({ path: C.logo, x: 0.28, y: 0.32, w: 1.7, h: 0.48 }); } catch (e) {}

    // Title
    txt(slide, d.title, 0.28, H * 0.33, W * 0.46 - 0.36, 1.5,
      { fontSize: 33, bold: true, color: textC });
    // Subtitle
    if (d.subtitle) {
      txt(slide, d.subtitle, 0.28, H * 0.62, W * 0.46 - 0.36, 0.52,
        { fontSize: 15, color: subC });
    }
    // Date
    if (d.date) {
      txt(slide, d.date, 0.28, H - 0.62, W * 0.38, 0.35,
        { fontSize: 11, color: subC });
    }
  }

  // SECTION DIVIDER
  function buildSection(slide, d, C) {
    // Left accent bar
    box(slide, PAD, H * 0.22, 0.07, H * 0.56, C.accent);
    let y = H * 0.25;
    if (d.number) {
      txt(slide, d.number, PAD + 0.22, y, 3.5, 1.05,
        { fontSize: 64, bold: true, color: C.accent });
      y += 0.9;
    }
    txt(slide, d.title, PAD + 0.22, y, CW - 0.22, 1.3,
      { fontSize: 38, bold: true, color: C.t1 });
    if (d.subtitle) {
      txt(slide, d.subtitle, PAD + 0.22, y + 1.35, CW * 0.65, 0.62,
        { fontSize: 17, color: C.t2 });
    }
  }

  // INDEX / TOC
  function buildIndex(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    heading(slide, d.heading || 'Contents', C);

    const items = d.items || [];
    const half  = Math.ceil(items.length / 2);
    const cw    = (CW - GAP) / 2;
    const rh    = CON_H / Math.max(half, 1);

    [items.slice(0, half), items.slice(half)].forEach((col, ci) => {
      const x = PAD + ci * (cw + GAP);
      col.forEach((item, i) => {
        const y = CON_Y + i * rh;
        // Row separator
        box(slide, x, y, cw, 0.012, C.border);
        txt(slide, item.number || String(i + 1 + ci * half).padStart(2, '0'),
          x + 0.1, y + 0.06, 0.5, rh - 0.12,
          { fontSize: 12, bold: true, color: C.accent, valign: 'middle' });
        txt(slide, item.title, x + 0.65, y + 0.06, cw - 1.2, rh - 0.12,
          { fontSize: 13, color: C.t1, valign: 'middle' });
        if (item.page) {
          txt(slide, item.page, x + cw - 0.68, y + 0.06, 0.6, rh - 0.12,
            { fontSize: 11, color: C.t2, align: 'right', valign: 'middle' });
        }
      });
    });
  }

  // METRICS-4
  function buildMetrics4(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    heading(slide, d.heading || d.slide_title, C);

    const stats = (d.stats || []).slice(0, 4);
    const sw    = CW / stats.length;
    const cy    = CON_Y + CON_H / 2 - 0.82;

    stats.forEach((s, i) => {
      const x = PAD + i * sw;
      if (i > 0) box(slide, x, CON_Y + CON_H * 0.1, 0.012, CON_H * 0.8, C.border);
      txt(slide, s.value, x, cy, sw, 1.05,
        { fontSize: 50, bold: true, color: C.t1, align: 'center' });
      txt(slide, s.label, x, cy + 1.1, sw, 0.42,
        { fontSize: 12, color: C.t2, align: 'center' });
    });
  }

  // CARDS (2–6 columns, optional photos)
  function buildCards(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    heading(slide, d.heading || d.slide_title, C);

    const cols   = d.cols || d.columns || [];
    const n      = cols.length;
    const cw     = (CW - GAP * (n - 1)) / n;
    const withImg = cols.some(c => c.image);

    cols.forEach((col, i) => {
      card(slide, PAD + i * (cw + GAP), CON_Y, cw, CON_H, col, C, withImg);
    });
  }

  // ROWS (2–4 horizontal cards, optional photos)
  function buildRows(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    heading(slide, d.heading || d.slide_title, C);

    const items  = d.items || d.columns || [];
    const n      = items.length;
    const rh     = (CON_H - GAP * (n - 1)) / n;
    const withImg = items.some(r => r.image);
    const TW     = 2.85;
    const IW     = withImg ? 2.75 : 0;
    const BW     = CW - TW - IW - (withImg ? GAP * 0.5 : 0);

    items.forEach((item, i) => {
      const y = CON_Y + i * (rh + GAP);
      box(slide, PAD, y, CW, rh, C.card, C.border);
      txt(slide, item.title, PAD + 0.18, y + 0.1, TW - 0.18, rh - 0.2,
        { fontSize: 14, bold: true, color: C.t1 });
      txt(slide, item.body, PAD + TW, y + 0.1, BW, rh - 0.2,
        { fontSize: 11, color: C.t2 });
      if (withImg && item.image) {
        photo(slide, item.image,
          PAD + CW - IW + 0.04, y + 0.04, IW - 0.08, rh - 0.08);
      }
    });
  }

  // STATEMENT + PHOTO HALF
  function buildStatement(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    if (d.heading || d.slide_title) heading(slide, d.heading || d.slide_title, C);

    const splitX = W * 0.49;
    box(slide, PAD, CON_Y, splitX - PAD - GAP, CON_H, C.card, C.border);
    txt(slide, d.text, PAD + 0.3, CON_Y + 0.3, splitX - PAD - GAP - 0.6, CON_H - 0.6,
      { fontSize: 24, bold: true, color: C.t1 });

    const px = splitX + GAP;
    const pw = W - px - 0.04;
    if (d.image) {
      photo(slide, d.image, px, CON_Y, pw, CON_H);
    } else {
      box(slide, px, CON_Y, pw, CON_H, C.accent);
    }
  }

  // STARTUPS-4
  function buildStartups4(slide, d, C, PT, logos) {
    hdr(slide, C, PT, logos);
    heading(slide, d.heading || d.slide_title, C);

    const startups = (d.startups || []).slice(0, 4);
    const n      = startups.length || 4;
    const cw     = (CW - GAP * (n - 1)) / n;
    const photoH = CON_H * 0.37;

    startups.forEach((s, i) => {
      const x = PAD + i * (cw + GAP);
      box(slide, x, CON_Y, cw, CON_H, C.card, C.border);
      if (s.image) photo(slide, s.image, x + 0.01, CON_Y + 0.01, cw - 0.02, photoH - 0.02);

      let ty = CON_Y + photoH + 0.14;

      txt(slide, s.name, x + 0.14, ty, cw - 0.28, 0.32,
        { fontSize: 13, bold: true, color: C.t1, valign: 'top' });
      ty += 0.33;

      if (s.tags && s.tags.length) {
        txt(slide, s.tags.join(' · '), x + 0.14, ty, cw - 0.28, 0.24,
          { fontSize: 9.5, bold: true, color: C.accent });
        ty += 0.27;
      }

      if (s.body) {
        txt(slide, s.body, x + 0.14, ty, cw - 0.28, 1.1,
          { fontSize: 10.5, color: C.t2, valign: 'top' });
      }

      // Metric at card bottom
      if (s.metric_value) {
        const my = CON_Y + CON_H - 0.74;
        box(slide, x + 0.14, my, cw - 0.28, 0.012, C.border);
        txt(slide, s.metric_value, x + 0.14, my + 0.06, cw - 0.28, 0.34,
          { fontSize: 17, bold: true, color: C.t1 });
        if (s.metric_label) {
          txt(slide, s.metric_label, x + 0.14, my + 0.39, cw - 0.28, 0.24,
            { fontSize: 9.5, color: C.t2 });
        }
      }
    });
  }

  // CLOSING
  function buildClosing(slide, d, C) {
    box(slide, 0, 0, 0.12, H, C.accent);
    try { slide.addImage({ path: C.logo, x: 0.28, y: 0.32, w: 1.7, h: 0.48 }); } catch (e) {}
    txt(slide, 'Thank You', PAD, H * 0.28, W - PAD * 2, 1.6,
      { fontSize: 54, bold: true, color: C.t1, align: 'center' });
    if (d.subtitle) {
      txt(slide, d.subtitle, PAD, H * 0.57, W - PAD * 2, 0.52,
        { fontSize: 19, color: C.t2, align: 'center' });
    }
    if (d.contact) {
      txt(slide, d.contact, PAD, H * 0.72, W - PAD * 2, 0.42,
        { fontSize: 14, color: C.accent, align: 'center' });
    }
  }

  // ── Main ─────────────────────────────────────────────────────────────────

  async function generateDeck(deck) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title  = deck.title  || 'PnP Presentation';
    pptx.author = 'Plug and Play';

    const C     = THEME[deck.theme] || THEME.light;
    const PT    = deck.title || '';
    const logos = deck.logos || ['pnp'];

    for (const d of (deck.slides || [])) {
      const s = pptx.addSlide();
      s.background = { color: C.bg };
      const t = (d.type || '').toLowerCase();

      if      (t === 'cover-light' || t === 'cover-dark' || t === 'cover')
        buildCover(s, d, C);
      else if (t === 'section')
        buildSection(s, d, C);
      else if (t === 'index-numbered' || t === 'toc' || t === 'contents')
        buildIndex(s, d, C, PT, logos);
      else if (t === 'metrics-4' || t === 'metrics')
        buildMetrics4(s, d, C, PT, logos);
      else if (t === 'cards')
        buildCards(s, d, C, PT, logos);
      else if (t === 'rows')
        buildRows(s, d, C, PT, logos);
      else if (t === 'statement-photo-half' || t === 'statement')
        buildStatement(s, d, C, PT, logos);
      else if (t === 'startups-4' || t === 'startups')
        buildStartups4(s, d, C, PT, logos);
      else if (t === 'closing')
        buildClosing(s, d, C);
      else {
        // Unknown type — placeholder
        hdr(s, C, PT, logos);
        txt(s, '⚠ Unknown slide type: "' + d.type + '"',
          PAD, H / 2 - 0.3, CW, 0.6,
          { fontSize: 15, color: 'DD2222', align: 'center' });
      }
    }

    const filename = (deck.title || 'PnP-Deck')
      .replace(/[^\w\s\-]/g, '').replace(/\s+/g, '_').slice(0, 50) + '.pptx';

    await pptx.writeFile({ fileName: filename });

    document.body.innerHTML = [
      '<div style="font-family:Arial;padding:3rem;text-align:center;max-width:480px;margin:auto">',
      '<div style="font-size:2.5rem">✅</div>',
      '<p style="font-size:1.1rem;font-weight:bold;color:#1D2A57;margin:.5rem 0">',
      filename, '</p>',
      '<p style="color:#44527A;font-size:.9rem">Check your downloads folder.</p>',
      '</div>',
    ].join('');
  }

  // ── Auto-run ─────────────────────────────────────────────────────────────

  window.generateDeck = generateDeck;

  window.addEventListener('load', function () {
    if (window.PNP_DECK) {
      generateDeck(window.PNP_DECK).catch(function (e) {
        document.body.innerHTML =
          '<p style="color:red;padding:1.5rem;font-family:Arial">Error: ' + e.message + '</p>';
        console.error(e);
      });
    }
  });

}());
