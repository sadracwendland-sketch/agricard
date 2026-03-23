/* =====================================================
   AgriCard – Card Renderer v1.0
   Renderização pixel-perfect de slides PPTX no Canvas

   Responsabilidades:
   - Converter coordenadas EMU → px na escala do canvas
   - Substituir placeholders por dados reais
   - Preservar: fonte, tamanho, cor, alinhamento, posição
   - Substituição inteligente: shrink ou truncate
   - Suporte a imagens (logo_cliente, logo_marca)
   ===================================================== */

'use strict';

const SlideRenderer = {

  // Image cache to avoid repeated loading
  _imgCache: {},

  /* ═══════════════════════════════════════════════════════════════
     MAIN RENDER FUNCTION
  ═══════════════════════════════════════════════════════════════ */
  /**
   * Renderiza um slide PPTX no canvas com os dados substituídos.
   * @param {HTMLCanvasElement} canvas   - target canvas
   * @param {object}           slide    - parsed slide data from PptxStudio
   * @param {object}           data     - { placeholder: value }  resolved data
   * @param {object}           phMap    - { placeholder: fieldKey } mapping
   * @param {string}           overflow - 'shrink' | 'truncate'
   * @param {File}             [file]   - original PPTX file (for embedded images)
   * @param {number}           [scale]  - canvas scale factor (default: 1)
   */
  async renderSlide(canvas, slide, data, phMap, overflow = 'shrink', file = null, scale = 1) {
    if (!canvas || !slide) return;

    const slideW = slide.isPdf ? slide.width  : slide.width  / (914400 / 96);
    const slideH = slide.isPdf ? slide.height : slide.height / (914400 / 96);

    canvas.width  = Math.round(slideW * scale);
    canvas.height = Math.round(slideH * scale);

    const ctx    = canvas.getContext('2d');
    const scaleX = canvas.width  / (slide.isPdf ? slide.width  : slide.width  / (914400 / 96));
    const scaleY = canvas.height / (slide.isPdf ? slide.height : slide.height / (914400 / 96));

    // Background
    ctx.fillStyle = slide.bgColor || '#FFFFFF';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // If PDF slide, draw the thumbnail image as background
    if (slide.isPdf && slide.thumbDataUrl) {
      try {
        const img = await SlideRenderer._loadImage(slide.thumbDataUrl);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        // For PDF: no further element-level rendering (just overlay data)
        await SlideRenderer._renderDataOverlayOnPdf(ctx, slide, data, phMap, overflow, canvas.width, canvas.height);
        return;
      } catch (e) { /* fallback to blank */ }
    }

    // Pre-load embedded images from PPTX
    let zipImages = {};
    if (file && !slide.isPdf) {
      zipImages = await SlideRenderer._extractZipImages(file, slide);
    }

    // Draw elements in slide order
    for (const el of slide.elements) {
      if (el.type === 'image') {
        await this._drawImageElement(ctx, el, slide, scaleX, scaleY, zipImages);
      } else if (el.type === 'text') {
        this._drawTextElement(ctx, el, slide, scaleX, scaleY, data, phMap, overflow);
      }
    }
  },

  /* ─────────────────────────────────────────────────────────────
     Draw image element
  ───────────────────────────────────────────────────────────── */
  async _drawImageElement(ctx, el, slide, scaleX, scaleY, zipImages) {
    const emuToCanvasPx = (emu, scale) => emu * scale / (914400 / 96);
    const x = emuToCanvasPx(el.x, scaleX);
    const y = emuToCanvasPx(el.y, scaleY);
    const w = emuToCanvasPx(el.w, scaleX);
    const h = emuToCanvasPx(el.h, scaleY);

    if (!w || !h) return;

    // Check if this placeholder image should be replaced by data
    const isLogoClient = el.rawText?.includes('logo_cliente');
    const isLogoMarca  = el.rawText?.includes('logo_marca');

    let src = null;

    if (isLogoClient && zipImages['logo_cliente']) {
      src = zipImages['logo_cliente'];
    } else if (isLogoMarca && zipImages['logo_marca']) {
      src = zipImages['logo_marca'];
    } else if (el.rEmbed && zipImages[el.rEmbed]) {
      src = zipImages[el.rEmbed];
    }

    if (!src) return;

    try {
      const img = await this._loadImage(src);
      // Preserve aspect ratio, fit inside box
      const { drawX, drawY, drawW, drawH } = this._fitImage(img, x, y, w, h);
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.clip();
      ctx.drawImage(img, drawX, drawY, drawW, drawH);
      ctx.restore();
    } catch (e) { /* skip broken images */ }
  },

  _fitImage(img, x, y, boxW, boxH) {
    const imgAspect = img.width / img.height;
    const boxAspect = boxW / boxH;
    let drawW, drawH, drawX, drawY;

    if (imgAspect > boxAspect) {
      drawW = boxW;
      drawH = boxW / imgAspect;
      drawX = x;
      drawY = y + (boxH - drawH) / 2;
    } else {
      drawH = boxH;
      drawW = boxH * imgAspect;
      drawX = x + (boxW - drawW) / 2;
      drawY = y;
    }

    return { drawX, drawY, drawW, drawH };
  },

  /* ─────────────────────────────────────────────────────────────
     Draw text element with placeholder substitution
  ───────────────────────────────────────────────────────────── */
  _drawTextElement(ctx, el, slide, scaleX, scaleY, data, phMap, overflow) {
    const emuToCanvasPx = (emu, scale) => emu * scale / (914400 / 96);

    const x  = emuToCanvasPx(el.x, scaleX);
    const y  = emuToCanvasPx(el.y, scaleY);
    const bw = emuToCanvasPx(el.w, scaleX);
    const bh = emuToCanvasPx(el.h, scaleY);

    if (!bw || !bh) return;

    // Fill background
    if (el.fillColor && el.fillColor !== 'transparent') {
      ctx.fillStyle = el.fillColor;
      ctx.fillRect(x, y, bw, bh);
    }

    // Build paragraphs with placeholder substitution
    const paragraphs = this._buildParagraphs(el, data, phMap, scaleY);
    if (!paragraphs.length) return;

    // Calculate total text block height for vertical alignment
    const totalH = paragraphs.reduce((sum, p) => sum + p.lineHeight, 0);

    // Vertical alignment (default: top)
    const anchor = el.anchor || 't'; // t, ctr, b
    let curY;
    if (anchor === 'ctr') curY = y + (bh - totalH) / 2;
    else if (anchor === 'b') curY = y + bh - totalH - emuToCanvasPx(el.bIns || 45720, scaleY);
    else curY = y + emuToCanvasPx(el.tIns || 45720, scaleY);

    const lPad = emuToCanvasPx(el.lIns || 91440, scaleX);
    const rPad = emuToCanvasPx(el.rIns || 91440, scaleX);
    const maxW = bw - lPad - rPad;

    for (const para of paragraphs) {
      const { runs, lineHeight, align } = para;

      // Render each run in the paragraph
      let lineX;
      if (align === 'center')      lineX = x + bw / 2;
      else if (align === 'right')  lineX = x + bw - rPad;
      else                          lineX = x + lPad;

      // Draw runs sequentially (inline)
      this._drawInlineRuns(ctx, runs, lineX, curY, maxW, align, overflow, lineHeight);

      curY += lineHeight;
    }
  },

  _buildParagraphs(el, data, phMap, scaleY) {
    if (!el.textRuns?.length) return [];

    // Substitute placeholders in text runs
    const substituted = el.textRuns.map(run => {
      let text = run.text;
      // Replace all {{placeholder}} occurrences
      text = text.replace(/\{\{(\w+)\}\}/g, (match, ph) => {
        const fieldKey = phMap[ph.toLowerCase()] || ph.toLowerCase();
        return data[fieldKey] ?? data[ph.toLowerCase()] ?? match;
      });
      return { ...run, text };
    });

    // Group runs into paragraphs by '\n' boundaries
    const paragraphs = [];
    let currentRuns  = [];
    let align        = 'left';

    const flushParagraph = () => {
      if (currentRuns.length === 0) return;
      const maxFontSize = Math.max(...currentRuns.map(r => r.fontSizePt || 12));
      const lineHeight  = maxFontSize * 1.333 * scaleY * 1.25;
      paragraphs.push({ runs: currentRuns, lineHeight, align });
      currentRuns = [];
    };

    substituted.forEach(run => {
      align = run.textAlign || 'left';
      if (run.text.includes('\n')) {
        const parts = run.text.split('\n');
        parts.forEach((part, idx) => {
          if (part) currentRuns.push({ ...run, text: part });
          if (idx < parts.length - 1) flushParagraph();
        });
      } else {
        currentRuns.push(run);
      }
    });
    flushParagraph();

    return paragraphs;
  },

  _drawInlineRuns(ctx, runs, lineX, lineY, maxW, align, overflow, lineHeight) {
    if (!runs.length) return;

    // Measure total width of all runs at natural scale
    const runMeasures = runs.map(run => {
      const fsPx  = this._fontSizePx(run);
      const font  = this._fontStr(run, fsPx);
      ctx.font    = font;
      return { run, fsPx, font, w: ctx.measureText(run.text).width };
    });

    const totalW = runMeasures.reduce((s, m) => s + m.w, 0);

    // Overflow handling
    let scaleFactor = 1;
    if (totalW > maxW && maxW > 0) {
      if (overflow === 'shrink') {
        scaleFactor = maxW / totalW;
      } else {
        // truncate: handled per run below
      }
    }

    // Starting X for left-aligned baseline
    let runX;
    if (align === 'center')     runX = lineX - (totalW * scaleFactor) / 2;
    else if (align === 'right') runX = lineX - totalW * scaleFactor;
    else                         runX = lineX;

    const baselineY = lineY + lineHeight * 0.8;

    for (const { run, fsPx, font, w } of runMeasures) {
      const scaledFs = fsPx * scaleFactor;
      ctx.font = this._fontStr(run, scaledFs);
      ctx.fillStyle = run.color || '#000000';
      ctx.textAlign = 'left'; // we handle alignment manually

      let displayText = run.text;
      const scaledW   = w * scaleFactor;

      if (overflow === 'truncate' && scaledW > maxW) {
        while (displayText.length > 0 &&
               ctx.measureText(displayText + '…').width > maxW) {
          displayText = displayText.slice(0, -1);
        }
        displayText += '…';
      }

      ctx.fillText(displayText, runX, baselineY);
      runX += scaledW;
    }
  },

  _fontSizePx(run) {
    return (run.fontSizePt || 12) * 1.333;
  },

  _fontStr(run, fsPx) {
    const italic = run.italic ? 'italic ' : '';
    const bold   = run.bold   ? 'bold '   : '';
    const family = this._safeFont(run.fontFamily);
    return `${italic}${bold}${fsPx.toFixed(2)}px ${family}`;
  },

  /**
   * Mapeamento HelveticaNeueLT Pro → Barlow (Google Fonts)
   * Condicionados e variantes mapeados para equivalentes web disponíveis.
   */
  _safeFont(family) {
    if (!family || family === '+mj-lt' || family === '+mn-lt') {
      return '"Barlow", "Helvetica Neue", Arial, sans-serif';
    }
    // Condensed variants
    if (/HelveticaNeueLT.*(?:Cn|BlkEx|HvCn|BdCn|MdCn|LtCn|93|87|77|67|57|47)/i.test(family)) {
      return '"Barlow Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    // Semi-condensed / Extended
    if (/HelveticaNeueLT.*(?:Ex|SemiCn|53|63|73)/i.test(family)) {
      return '"Barlow Semi Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    // Any other HelveticaNeueLT
    if (family.includes('HelveticaNeueLT') || family.includes('Helvetica')) {
      return '"Barlow", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    // Clean up font name for web usage
    const cleaned = family.replace(/\+/g, ' ').replace(/,/g, ', ').trim();
    return `"${cleaned}", "Barlow", Arial, sans-serif`;
  },

  /* ─────────────────────────────────────────────────────────────
     PDF overlay: minimal data rendering on top of background
  ───────────────────────────────────────────────────────────── */
  async _renderDataOverlayOnPdf(ctx, slide, data, phMap, overflow, cw, ch) {
    // For PDFs (no element structure), create simple centered overlays
    // based on placeholder map. Admin should use PPTX for full control.
    const mapped = Object.entries(phMap).filter(([, v]) => v && data[v] !== undefined);

    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(0,0,0,0.75)';

    let y = ch * 0.7;
    const step = ch * 0.06;

    mapped.forEach(([ph, fieldKey]) => {
      const val = data[fieldKey] || '';
      if (!val) return;
      const fsPx = Math.max(14, ch * 0.028);
      ctx.font      = `bold ${fsPx}px Inter, Arial`;
      ctx.fillStyle = '#FFFFFF';
      // Shadow for readability
      ctx.shadowColor = 'rgba(0,0,0,0.9)';
      ctx.shadowBlur  = 4;
      ctx.fillText(String(val), cw / 2, y);
      ctx.shadowBlur = 0;
      y += step;
    });
  },

  /* ─────────────────────────────────────────────────────────────
     Extract embedded images from PPTX zip
  ───────────────────────────────────────────────────────────── */
  async _extractZipImages(file, slide) {
    const result = {};
    if (!file || !slide.rels) return result;

    try {
      const JSZip = window.JSZip;
      if (!JSZip) return result;

      const zip = await JSZip.loadAsync(await file.arrayBuffer());

      for (const [relId, relTarget] of Object.entries(slide.rels)) {
        if (!relTarget) continue;
        // Resolve path
        const imgPath = relTarget.startsWith('..')
          ? 'ppt/' + relTarget.replace(/^\.\.\//, '')
          : 'ppt/slides/' + relTarget;

        const imgFile = zip.file(imgPath) ||
                        zip.file(imgPath.replace('/slides', '')) ||
                        zip.file('ppt/media/' + relTarget.split('/').pop());

        if (imgFile) {
          try {
            const blob    = await imgFile.async('blob');
            const dataUrl = await this._blobToDataUrl(blob);
            result[relId] = dataUrl;
          } catch (e) { /* skip */ }
        }
      }
    } catch (e) {
      console.warn('Image extraction error:', e);
    }

    return result;
  },

  /* ─────────────────────────────────────────────────────────────
     Utilities
  ───────────────────────────────────────────────────────────── */
  _loadImage(src) {
    if (this._imgCache[src]) return Promise.resolve(this._imgCache[src]);
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload  = () => { this._imgCache[src] = img; resolve(img); };
      img.onerror = reject;
      img.src = src;
    });
  },

  _blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  },
};
