/* =================================================================
   AgriCard – PPTX Engine v2.0
   Motor completo de parsing e renderização pixel-perfect de PPTX.

   FLUXO:
   1. parsePptx(file) → extrai slide 1: shapes, textos, imagens, cores
   2. detectPlaceholders(shapes) → lista de {{campo}} encontrados
   3. renderToCanvas(shapes, data, canvas, opts) → renderiza pixel-perfect
   4. exportPNG / exportPDF → exportação em alta resolução

   DEPENDÊNCIAS CDN (carregadas sob demanda):
   - JSZip  3.10+   (https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js)
   ================================================================= */

const PptxEngine = {

  /* ─────────────────────────────────────────────────
     CONSTANTES DE CONVERSÃO
     EMU  = English Metric Unit (1 inch = 914400 EMU)
     Slide padrão PPTX = 9144000 × 5143500 EMU = 10×5.625 polegadas
  ───────────────────────────────────────────────── */
  EMU_PER_INCH: 914400,
  SLIDE_W_EMU:  9144000,
  SLIDE_H_EMU:  5143500,

  emuToPx(emu, dpi = 96) {
    return (emu / this.EMU_PER_INCH) * dpi;
  },

  /* ─────────────────────────────────────────────────
     LOAD JSZIP (lazy)
  ───────────────────────────────────────────────── */
  async _loadJSZip() {
    if (window.JSZip) return;
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  },

  /* ─────────────────────────────────────────────────
     PARSE PPTX
     Retorna objeto com:
       slideW, slideH  (EMU)
       shapes[]        (ver _parseShape)
       bgColor         (hex)
       bgImageData     (base64 se houver imagem de fundo)
  ───────────────────────────────────────────────── */
  async parsePptx(file) {
    await this._loadJSZip();
    const ab  = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(ab);

    // Dimensões do slide (presentation.xml)
    const presXml = await this._getText(zip, 'ppt/presentation.xml');
    const { slideW, slideH } = this._parseSlideDimensions(presXml);

    // Slide 1
    const slideXml = await this._getText(zip, 'ppt/slides/slide1.xml');
    const slideRels = await this._getText(zip, 'ppt/slides/_rels/slide1.xml.rels').catch(() => '');

    // Media map (rId → base64)
    const mediaMap = await this._buildMediaMap(zip, slideRels, 'ppt/slides/');

    // Theme (cores e fontes)
    const themeXml = await this._getText(zip, 'ppt/theme/theme1.xml').catch(() => '');

    // Layout/master (fundo)
    const layoutRels = await this._getText(zip, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels').catch(() => '');
    const layoutXml  = await this._getText(zip, 'ppt/slideLayouts/slideLayout1.xml').catch(() => '');
    const masterXml  = await this._getText(zip, 'ppt/slideMasters/slideMaster1.xml').catch(() => '');

    const shapes   = this._parseShapes(slideXml, mediaMap, slideW, slideH);
    const bgColor  = this._parseBgColor(slideXml, layoutXml, masterXml, themeXml);
    const bgImageData = this._parseBgImage(slideXml, layoutXml, mediaMap);

    return { slideW, slideH, shapes, bgColor, bgImageData, zip, mediaMap };
  },

  /* ─── Dimensões do slide ─── */
  _parseSlideDimensions(xml) {
    const m = xml.match(/<p:sldSz[^>]+cx="(\d+)"[^>]+cy="(\d+)"/);
    if (m) return { slideW: parseInt(m[1]), slideH: parseInt(m[2]) };
    return { slideW: this.SLIDE_W_EMU, slideH: this.SLIDE_H_EMU };
  },

  /* ─── Cor de fundo do slide ─── */
  _parseBgColor(slideXml, layoutXml, masterXml, themeXml) {
    for (const xml of [slideXml, layoutXml, masterXml]) {
      const bg = xml.match(/<p:bg>[\s\S]*?<\/p:bg>/);
      if (bg) {
        const solid = bg[0].match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/);
        if (solid) return '#' + solid[1];
        const sysClr = bg[0].match(/<a:sysClr[^>]+lastClr="([0-9A-Fa-f]{6})"/);
        if (sysClr) return '#' + sysClr[1];
      }
    }
    return '#FFFFFF';
  },

  _parseBgImage(slideXml, layoutXml, mediaMap) {
    for (const xml of [slideXml, layoutXml]) {
      const bgM = xml.match(/<p:bg>[\s\S]*?<\/p:bg>/);
      if (bgM) {
        const rId = bgM[0].match(/r:embed="(rId\d+)"/);
        if (rId && mediaMap[rId[1]]) return mediaMap[rId[1]];
      }
    }
    return null;
  },

  /* ─── Media Map ─── */
  async _buildMediaMap(zip, relsXml, basePath) {
    const map  = {};
    const rels = [...relsXml.matchAll(/Id="(rId\d+)"[^>]+Target="([^"]+)"/g)];
    for (const [, rId, target] of rels) {
      const fullPath = basePath + target.replace('../', '../').replace(/^\.\.\//, 'ppt/');
      const fixedPath = fullPath.replace('ppt/slides/../', 'ppt/');
      const f = zip.file(fixedPath);
      if (f) {
        try {
          const blob = await f.async('blob');
          map[rId]   = await this._blobToDataUrl(blob);
        } catch {}
      }
    }
    return map;
  },

  /* ─────────────────────────────────────────────────
     PARSE SHAPES
  ───────────────────────────────────────────────── */
  _parseShapes(xml, mediaMap, slideW, slideH) {
    const shapes = [];
    const spTree = xml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTree) return shapes;

    const tree = spTree[1];

    // Text shapes: <p:sp>
    const spMatches = [...tree.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    for (const [, sp] of spMatches) {
      const shape = this._parseTextShape(sp);
      if (shape) shapes.push(shape);
    }

    // Pictures: <p:pic>
    const picMatches = [...tree.matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)];
    for (const [, pic] of picMatches) {
      const shape = this._parsePicShape(pic, mediaMap);
      if (shape) shapes.push(shape);
    }

    // Connectors/lines: <p:cxnSp>
    const cxnMatches = [...tree.matchAll(/<p:cxnSp>([\s\S]*?)<\/p:cxnSp>/g)];
    for (const [, cxn] of cxnMatches) {
      const shape = this._parseConnector(cxn);
      if (shape) shapes.push(shape);
    }

    // Group shapes: <p:grpSp>
    const grpMatches = [...tree.matchAll(/<p:grpSp>([\s\S]*?)<\/p:grpSp>/g)];
    for (const [, grp] of grpMatches) {
      // Recursivo: extrai shapes do grupo
      const inner = this._parseShapes(`<p:spTree>${grp}</p:spTree>`, mediaMap, slideW, slideH);
      shapes.push(...inner);
    }

    return shapes;
  },

  /* ─── Text Shape ─── */
  _parseTextShape(sp) {
    const xfrm = this._parseXfrm(sp);
    if (!xfrm) return null;

    const paras = this._parseParas(sp);
    const solidFill = sp.match(/<p:sp>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/)
      || sp.match(/<p:spPr>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);

    return {
      type: 'text',
      x: xfrm.x, y: xfrm.y, w: xfrm.w, h: xfrm.h, rot: xfrm.rot,
      flipH: xfrm.flipH, flipV: xfrm.flipV,
      paras,
      fill: solidFill ? '#' + solidFill[1] : null,
      border: this._parseBorder(sp)
    };
  },

  /* ─── Picture Shape ─── */
  _parsePicShape(pic, mediaMap) {
    const xfrm = this._parseXfrm(pic);
    if (!xfrm) return null;
    const rId = (pic.match(/r:embed="(rId\d+)"/) || [])[1];
    const imgData = rId ? mediaMap[rId] : null;

    return {
      type: 'image',
      x: xfrm.x, y: xfrm.y, w: xfrm.w, h: xfrm.h, rot: xfrm.rot,
      flipH: xfrm.flipH, flipV: xfrm.flipV,
      imageData: imgData,
      rId,
      placeholder: this._extractPicPlaceholder(pic)
    };
  },

  /* ─── Connector/Line Shape ─── */
  _parseConnector(cxn) {
    const xfrm = this._parseXfrm(cxn);
    if (!xfrm) return null;
    const solidFill = cxn.match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    const lnMatch = cxn.match(/<a:ln\s+w="(\d+)"/);
    return {
      type: 'connector',
      x: xfrm.x, y: xfrm.y, w: xfrm.w, h: xfrm.h, rot: xfrm.rot,
      color: solidFill ? '#' + solidFill[1] : '#000000',
      lineWidth: lnMatch ? parseInt(lnMatch[1]) / 12700 : 1
    };
  },

  /* ─── Xfrm (posição/tamanho) ─── */
  _parseXfrm(xml) {
    const xfrm = xml.match(/<a:xfrm([^>]*)>([\s\S]*?)<\/a:xfrm>/);
    if (!xfrm) return null;
    const attrs = xfrm[1];
    const body  = xfrm[2];
    const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
    const ext = body.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!off || !ext) return null;
    return {
      x:     parseInt(off[1]),
      y:     parseInt(off[2]),
      w:     parseInt(ext[1]),
      h:     parseInt(ext[2]),
      rot:   parseFloat((attrs.match(/rot="(-?\d+)"/) || [,0])[1]) / 60000,
      flipH: /flipH="1"/.test(attrs),
      flipV: /flipV="1"/.test(attrs)
    };
  },

  /* ─── Border ─── */
  _parseBorder(sp) {
    const ln = sp.match(/<a:ln\s+([^>]*)>([\s\S]*?)<\/a:ln>/);
    if (!ln) return null;
    const wAttr = (ln[1].match(/\bw="(\d+)"/) || [])[1];
    const solid = ln[2].match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    const noFill = /<a:noFill/.test(ln[2]);
    if (noFill) return null;
    return {
      width: wAttr ? parseInt(wAttr) / 12700 : 1,
      color: solid ? '#' + solid[1] : '#000000'
    };
  },

  /* ─── Paragraphs & Runs ─── */
  _parseParas(sp) {
    const txBody = sp.match(/<a:txBody>([\s\S]*?)<\/a:txBody>/);
    if (!txBody) return [];

    // Body properties
    const bodyPr = txBody[1].match(/<a:bodyPr([^>]*)>/);
    const bodyAttrs = bodyPr ? bodyPr[1] : '';
    const anchor = (bodyAttrs.match(/anchor="(\w+)"/) || [,'t'])[1];
    const wrap    = !/wrap="none"/.test(bodyAttrs);
    const autofit = /<a:normAutofit/.test(txBody[1]);
    const noAutofit = /<a:noAutofit/.test(txBody[1]);
    const spAutoFit = /<a:spAutoFit/.test(txBody[1]);

    const paras = [];
    const paraMatches = [...txBody[1].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)];

    for (const [, p] of paraMatches) {
      // Paragraph-level props
      const pPr = p.match(/<a:pPr([^>]*)>/);
      const pAttrs = pPr ? pPr[1] : '';
      const align = (pAttrs.match(/algn="(\w+)"/) || [,'l'])[1];
      const indent = parseFloat((pAttrs.match(/indent="(-?\d+)"/) || [,0])[1]);
      const marL   = parseFloat((pAttrs.match(/marL="(-?\d+)"/) || [,0])[1]);
      const spcBef  = this._parseSpc(p.match(/<a:spcBef>([\s\S]*?)<\/a:spcBef>/) );
      const spcAft  = this._parseSpc(p.match(/<a:spcAft>([\s\S]*?)<\/a:spcAft>/) );
      const lnSpc   = this._parseLnSpc(p.match(/<a:lnSpc>([\s\S]*?)<\/a:lnSpc>/) );

      // Default run props from <a:pPr><a:defRPr>
      const defRpr  = p.match(/<a:defRPr([^>]*)>/);
      const defAttrs = defRpr ? defRpr[1] : '';

      const runs = [];
      const runMatches = [...p.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)];
      for (const [, r] of runMatches) {
        const run = this._parseRun(r, defAttrs);
        if (run) runs.push(run);
      }

      // Breaks
      const brMatches = [...p.matchAll(/<a:br>([\s\S]*?)<\/a:br>/g)];
      // Add line-break markers
      let rIdx = 0;
      const pXml = p;
      const tokens = [];
      let pos = 0;
      const allTags = [...pXml.matchAll(/<a:(r|br)>([\s\S]*?)<\/a:\1>/g)];
      for (const tag of allTags) {
        if (tag[1] === 'br') tokens.push({ type: 'break' });
        else {
          const run = this._parseRun(tag[2], defAttrs);
          if (run) tokens.push({ type: 'run', ...run });
        }
      }

      paras.push({
        runs: tokens,
        align,
        indent: indent / 12700,
        marL:   marL   / 12700,
        spcBef, spcAft, lnSpc,
        anchor, wrap, autofit, noAutofit, spAutoFit
      });
    }
    return paras;
  },

  _parseSpc(m) {
    if (!m) return 0;
    const pts = m[1].match(/<a:spcPts val="(\d+)"/);
    if (pts) return parseInt(pts[1]) / 100; // hundredths of a point → points
    const pct = m[1].match(/<a:spcPct val="(\d+)"/);
    if (pct) return { pct: parseInt(pct[1]) / 100000 };
    return 0;
  },

  _parseLnSpc(m) {
    if (!m) return 1;
    const pts = m[1].match(/<a:spcPts val="(\d+)"/);
    if (pts) return { pts: parseInt(pts[1]) / 100 };
    const pct = m[1].match(/<a:spcPct val="(\d+)"/);
    if (pct) return parseInt(pct[1]) / 100000;
    return 1;
  },

  /* ─── Run ─── */
  _parseRun(r, defAttrs = '') {
    const rPr = r.match(/<a:rPr([^>]*)\/?>/) || r.match(/<a:rPr([^>]*)>/);
    const attrs = rPr ? rPr[1] : defAttrs;

    const tMatch = r.match(/<a:t>([\s\S]*?)<\/a:t>/);
    const text   = tMatch ? this._decodeXml(tMatch[1]) : '';

    // Font size (hundredths of a point)
    const szM = attrs.match(/\bsz="(\d+)"/);
    const fontSize = szM ? parseInt(szM[1]) / 100 : null;

    // Bold / italic / underline / strike
    const bold      = /\bb="1"/.test(attrs) || /\bb="true"/.test(attrs);
    const italic    = /\bi="1"/.test(attrs) || /\bi="true"/.test(attrs);
    const underline = /\bu="sng"/.test(attrs) || /\bu="dbl"/.test(attrs);
    const strike    = /\bstrike="sngStrike"/.test(attrs) || /\bstrike="dblStrike"/.test(attrs);

    // Color
    const solidFill = r.match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    const color = solidFill ? '#' + solidFill[1] : null;

    // Font family
    const latin = r.match(/<a:latin typeface="([^"]+)"/);
    const fontFamily = latin ? latin[1] : null;

    // Char spacing (hundredths of a point)
    const spcM = attrs.match(/\bspc="(-?\d+)"/);
    const charSpacing = spcM ? parseInt(spcM[1]) / 100 : 0;

    return { text, fontSize, bold, italic, underline, strike, color, fontFamily, charSpacing };
  },

  /* ─── Picture placeholder nome ─── */
  _extractPicPlaceholder(pic) {
    const desc = pic.match(/<p:cNvPr[^>]+descr="([^"]+)"/);
    if (desc) {
      const m = desc[1].match(/\{\{(\w+)\}\}/);
      if (m) return m[1];
    }
    return null;
  },

  /* ─────────────────────────────────────────────────
     DETECT PLACEHOLDERS
     Retorna array de { field, type, shapeIdx }
  ───────────────────────────────────────────────── */
  detectPlaceholders(shapes) {
    const found = new Map();

    shapes.forEach((shape, idx) => {
      if (shape.type === 'text') {
        shape.paras.forEach(para => {
          para.runs.forEach(token => {
            if (token.type === 'break') return;
            const matches = [...(token.text || '').matchAll(/\{\{(\w+)\}\}/g)];
            matches.forEach(m => {
              if (!found.has(m[1])) {
                found.set(m[1], { field: m[1], type: 'text', shapeIdx: idx });
              }
            });
          });
        });
      }
      if (shape.type === 'image' && shape.placeholder) {
        if (!found.has(shape.placeholder)) {
          found.set(shape.placeholder, { field: shape.placeholder, type: 'image', shapeIdx: idx });
        }
      }
    });

    return [...found.values()];
  },

  /* ─────────────────────────────────────────────────
     RENDER TO CANVAS — pixel-perfect
     data = { field: value, ... }
     options = { overflowMode: 'shrink'|'truncate', dpi: 96 }
  ───────────────────────────────────────────────── */
  async renderToCanvas(parsed, data, canvas, options = {}) {
    const { slideW, slideH, shapes, bgColor, bgImageData } = parsed;
    const DPI     = options.dpi || 96;
    const W       = this.emuToPx(slideW, DPI);
    const H       = this.emuToPx(slideH, DPI);
    const scale   = W / slideW; // px/EMU

    canvas.width  = W;
    canvas.height = H;
    canvas.style.width  = W + 'px';
    canvas.style.height = H + 'px';

    const ctx = canvas.getContext('2d');
    await this._ensureFonts(shapes);

    // ── Background ──
    ctx.fillStyle = bgColor || '#FFFFFF';
    ctx.fillRect(0, 0, W, H);

    if (bgImageData) {
      try {
        const img = await this._loadImage(bgImageData);
        ctx.drawImage(img, 0, 0, W, H);
      } catch {}
    }

    // ── Sort by z-index (order in array = z-order) ──
    for (const shape of shapes) {
      ctx.save();

      const x = shape.x * scale;
      const y = shape.y * scale;
      const w = shape.w * scale;
      const h = shape.h * scale;

      // Rotation
      if (shape.rot) {
        ctx.translate(x + w / 2, y + h / 2);
        ctx.rotate((shape.rot * Math.PI) / 180);
        ctx.translate(-(x + w / 2), -(y + h / 2));
      }

      if (shape.type === 'text') {
        await this._drawTextShape(ctx, shape, data, x, y, w, h, scale, options);
      } else if (shape.type === 'image') {
        await this._drawImageShape(ctx, shape, data, x, y, w, h);
      } else if (shape.type === 'connector') {
        this._drawConnector(ctx, shape, x, y, w, h, scale);
      }

      ctx.restore();
    }
  },

  /* ─── Draw Text Shape ─── */
  async _drawTextShape(ctx, shape, data, x, y, w, h, scale, opts = {}) {
    // Fill
    if (shape.fill) {
      ctx.fillStyle = shape.fill;
      ctx.fillRect(x, y, w, h);
    }
    // Border
    if (shape.border) {
      ctx.strokeStyle = shape.border.color;
      ctx.lineWidth   = shape.border.width * scale / 12700 * 96 / 72;
      ctx.strokeRect(x, y, w, h);
    }

    if (!shape.paras || shape.paras.length === 0) return;

    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();

    const overflowMode = opts.overflowMode || 'shrink';
    const PAD = 5 * (scale * this.SLIDE_W_EMU / 9144000 / 96);

    // Build lines
    const lines = this._buildLines(shape.paras, data, w - PAD * 2, scale, overflowMode);

    // Vertical anchor
    const totalH = this._calcTotalH(lines);
    const anchor = shape.paras[0]?.anchor || 't';
    let curY = y + PAD;
    if (anchor === 'ctr') curY = y + (h - totalH) / 2;
    else if (anchor === 'b') curY = y + h - totalH - PAD;

    for (const line of lines) {
      if (line.break) { curY += line.height || 0; continue; }
      const lineW = line.runs.reduce((s, r) => s + r.measuredW, 0);
      let curX = x + PAD;
      if (line.align === 'ctr' || line.align === 'center') curX = x + (w - lineW) / 2;
      else if (line.align === 'r' || line.align === 'right') curX = x + w - lineW - PAD;

      for (const run of line.runs) {
        ctx.font         = this._buildFont(run);
        ctx.fillStyle    = run.color || '#000000';
        ctx.textBaseline = 'alphabetic';
        if (run.charSpacing) ctx.letterSpacing = run.charSpacing + 'px';
        ctx.fillText(run.text, curX, curY + run.ascent);
        if (run.underline) {
          ctx.beginPath();
          ctx.strokeStyle = run.color || '#000000';
          ctx.lineWidth   = Math.max(1, run.fontSize * 0.07);
          ctx.moveTo(curX, curY + run.ascent + 2);
          ctx.lineTo(curX + run.measuredW, curY + run.ascent + 2);
          ctx.stroke();
        }
        if (run.strike) {
          ctx.beginPath();
          ctx.strokeStyle = run.color || '#000000';
          ctx.lineWidth   = Math.max(1, run.fontSize * 0.07);
          const midY = curY + run.ascent - run.fontSize * 0.3;
          ctx.moveTo(curX, midY);
          ctx.lineTo(curX + run.measuredW, midY);
          ctx.stroke();
        }
        curX += run.measuredW;
      }
      curY += line.height;
    }

    ctx.restore();
  },

  /* ─── Build Lines (word-wrap + placeholder replace) ─── */
  _buildLines(paras, data, maxW, scale, overflowMode) {
    const lines = [];
    const SCALE_FACTOR = scale * (9144000 / this.SLIDE_W_EMU);

    for (const para of paras) {
      // Flatten runs, replacing placeholders
      const flatRuns = [];
      for (const token of para.runs) {
        if (token.type === 'break') { flatRuns.push({ type: 'break' }); continue; }
        const replaced = this._replacePlaceholders(token.text || '', data);
        flatRuns.push({ ...token, text: replaced });
      }

      // Build measured runs
      const measuredRuns = [];
      for (const run of flatRuns) {
        if (run.type === 'break') { measuredRuns.push({ type: 'break' }); continue; }
        const fontSize = (run.fontSize || 18) * SCALE_FACTOR;
        const mRun = { ...run, fontSize, measuredRuns: [] };
        measuredRuns.push(mRun);
      }

      // Word-wrap
      const paraLines = this._wrapPara(measuredRuns, para, maxW, overflowMode);
      lines.push(...paraLines);
    }

    return lines;
  },

  _wrapPara(runs, para, maxW, overflowMode) {
    const lines    = [];
    let   curLine  = { runs: [], align: para.align, height: 0 };
    let   curW     = 0;
    const tmpCanvas = document.createElement('canvas');
    const tmpCtx    = tmpCanvas.getContext('2d');

    for (const run of runs) {
      if (run.type === 'break') {
        const h = this._lineHeight(curLine.runs);
        curLine.height = h.height;
        lines.push(curLine);
        curLine = { runs: [], align: para.align, height: 0 };
        curW    = 0;
        continue;
      }

      // Measure and split by words
      const font    = this._buildFont(run);
      tmpCtx.font   = font;
      const words   = run.text.split(/(\s+)/);
      let   pending = '';

      for (const word of words) {
        const candidate = pending + word;
        const cW = tmpCtx.measureText(candidate).width;

        if (curW + cW > maxW && curLine.runs.length > 0 && pending !== '') {
          // Wrap: flush pending
          const measured = tmpCtx.measureText(pending);
          const tm       = tmpCtx.measureText(pending);
          const asc      = tm.actualBoundingBoxAscent  || run.fontSize * 0.8;
          const desc     = tm.actualBoundingBoxDescent || run.fontSize * 0.2;
          curLine.runs.push({ ...run, text: pending, measuredW: measured.width, ascent: asc, descent: desc });
          const h = this._lineHeight(curLine.runs);
          curLine.height = h.height;
          lines.push(curLine);
          curLine = { runs: [], align: para.align, height: 0 };
          curW    = 0;
          pending = word;
        } else {
          pending = candidate;
        }
      }

      if (pending) {
        const measured = tmpCtx.measureText(pending);
        const tm       = tmpCtx.measureText(pending);
        const asc      = tm.actualBoundingBoxAscent  || run.fontSize * 0.8;
        const desc     = tm.actualBoundingBoxDescent || run.fontSize * 0.2;

        let finalText = pending;
        if (overflowMode === 'truncate') {
          while (curW + tmpCtx.measureText(finalText).width > maxW && finalText.length > 1) {
            finalText = finalText.slice(0, -1);
          }
          if (finalText !== pending) finalText = finalText.slice(0, -1) + '…';
        }
        const w = tmpCtx.measureText(finalText).width;
        curLine.runs.push({ ...run, text: finalText, measuredW: w, ascent: asc, descent: desc });
        curW += w;
      }
    }

    if (curLine.runs.length > 0) {
      const h = this._lineHeight(curLine.runs);
      curLine.height = h.height;
      lines.push(curLine);
    }

    // Empty paragraph = one blank line
    if (lines.length === 0) {
      const dummyFontSize = para.runs[0]?.fontSize || 18;
      lines.push({ runs: [], align: para.align, height: dummyFontSize * 1.2 });
    }

    return lines;
  },

  _lineHeight(runs) {
    if (!runs.length) return { height: 14, ascent: 11, descent: 3 };
    let maxH = 0;
    for (const r of runs) {
      const h = (r.fontSize || 14) * 1.2;
      if (h > maxH) maxH = h;
    }
    return { height: maxH };
  },

  _calcTotalH(lines) {
    return lines.reduce((s, l) => s + (l.height || 0), 0);
  },

  _buildFont(run) {
    const size   = run.fontSize || 14;
    const weight = run.bold   ? '700' : '400';
    const style  = run.italic ? 'italic' : 'normal';
    const family = this._safeFont(run.fontFamily);
    return `${style} ${weight} ${size}px ${family}`;
  },

  /**
   * Mapeia fontes HelveticaNeueLT Pro → Barlow (Google Fonts)
   */
  _safeFont(family) {
    if (!family || family === '+mj-lt' || family === '+mn-lt') {
      return '"Barlow", "Helvetica Neue", Arial, sans-serif';
    }
    if (/HelveticaNeueLT.*(?:Cn|BlkEx|HvCn|BdCn|MdCn|LtCn|93|87|77|67|57|47)/i.test(family)) {
      return '"Barlow Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    if (/HelveticaNeueLT.*(?:Ex|SemiCn|53|63|73)/i.test(family)) {
      return '"Barlow Semi Condensed", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    if (family.includes('HelveticaNeueLT') || family.includes('Helvetica')) {
      return '"Barlow", "Helvetica Neue", Helvetica, Arial, sans-serif';
    }
    const cleaned = family.replace(/\+/g, ' ').trim();
    return `"${cleaned}", "Barlow", Arial, sans-serif`;
  },

  /* ─── Draw Image Shape ─── */
  async _drawImageShape(ctx, shape, data, x, y, w, h) {
    let imgSrc = shape.imageData;

    // Check if placeholder has data override
    if (shape.placeholder && data[shape.placeholder]) {
      imgSrc = data[shape.placeholder];
    }

    if (!imgSrc) return;

    try {
      const img = await this._loadImage(imgSrc);
      ctx.drawImage(img, x, y, w, h);
    } catch {}
  },

  /* ─── Draw Connector ─── */
  _drawConnector(ctx, shape, x, y, w, h, scale) {
    ctx.beginPath();
    ctx.strokeStyle = shape.color;
    ctx.lineWidth   = shape.lineWidth;
    ctx.moveTo(x, y);
    ctx.lineTo(x + w, y + h);
    ctx.stroke();
  },

  /* ─────────────────────────────────────────────────
     REPLACE PLACEHOLDERS
  ───────────────────────────────────────────────── */
  _replacePlaceholders(text, data) {
    return text.replace(/\{\{(\w+)\}\}/g, (match, field) => {
      return data[field] !== undefined ? String(data[field]) : match;
    });
  },

  /* ─────────────────────────────────────────────────
     EXPORT
  ───────────────────────────────────────────────── */
  async exportPNG(parsed, data, dpi = 300) {
    const canvas = document.createElement('canvas');
    await this.renderToCanvas(parsed, data, canvas, { dpi });
    return canvas.toDataURL('image/png', 1.0);
  },

  async exportJPEG(parsed, data, dpi = 300, quality = 0.97) {
    const canvas = document.createElement('canvas');
    await this.renderToCanvas(parsed, data, canvas, { dpi });
    return canvas.toDataURL('image/jpeg', quality);
  },

  async exportPDF(parsed, data, dpi = 150) {
    // Carrega jsPDF sob demanda
    if (!window.jspdf) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js');
    }
    const { jsPDF } = window.jspdf;
    const canvas = document.createElement('canvas');
    await this.renderToCanvas(parsed, data, canvas, { dpi });

    const W_mm = (parsed.slideW / this.EMU_PER_INCH) * 25.4;
    const H_mm = (parsed.slideH / this.EMU_PER_INCH) * 25.4;
    const ori  = W_mm > H_mm ? 'l' : 'p';

    const pdf = new jsPDF({ orientation: ori, unit: 'mm', format: [W_mm, H_mm] });
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, W_mm, H_mm);
    return pdf.output('bloburl');
  },

  /* ─────────────────────────────────────────────────
     BATCH PROCESSING
  ───────────────────────────────────────────────── */
  async batchExportPNG(parsed, dataArray, dpi = 150, onProgress) {
    const urls = [];
    for (let i = 0; i < dataArray.length; i++) {
      const url = await this.exportPNG(parsed, dataArray[i], dpi);
      urls.push(url);
      if (onProgress) onProgress(i + 1, dataArray.length);
    }
    return urls;
  },

  /* ─────────────────────────────────────────────────
     CSV / EXCEL PARSER
  ───────────────────────────────────────────────── */
  async parseCSV(file) {
    const text  = await file.text();
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) return [];
    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const vals = this._splitCsvLine(lines[i]);
      const row  = {};
      headers.forEach((h, j) => { row[h] = (vals[j] || '').replace(/^"|"$/g, '').trim(); });
      rows.push(row);
    }
    return rows;
  },

  async parseExcel(file) {
    if (!window.XLSX) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    }
    const ab  = await file.arrayBuffer();
    const wb  = XLSX.read(ab, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const arr = XLSX.utils.sheet_to_json(ws, { raw: false });
    return arr;
  },

  _splitCsvLine(line) {
    const result = [];
    let inQ = false, cur = '';
    for (const ch of line) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { result.push(cur); cur = ''; }
      else { cur += ch; }
    }
    result.push(cur);
    return result;
  },

  /* ─────────────────────────────────────────────────
     HELPERS
  ───────────────────────────────────────────────── */
  _getText(zip, path) {
    const f = zip.file(path);
    if (!f) return Promise.resolve('');
    return f.async('text');
  },

  _blobToDataUrl(blob) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => res(e.target.result);
      r.onerror = rej;
      r.readAsDataURL(blob);
    });
  },

  _loadImage(src) {
    return new Promise((res, rej) => {
      if (!src) return rej(new Error('no src'));
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload  = () => res(img);
      img.onerror = () => rej(new Error('img load error'));
      img.src = src;
    });
  },

  _loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src; s.onload = resolve;
      s.onerror = () => reject(new Error(`Failed: ${src}`));
      document.head.appendChild(s);
    });
  },

  async _ensureFonts(shapes) {
    if (!document.fonts) return;
    const families = new Set(['"Barlow"', '"Barlow Condensed"', '"Barlow Semi Condensed"', 'Inter', 'Arial']);
    for (const s of shapes) {
      if (s.type !== 'text') continue;
      for (const p of (s.paras || [])) {
        for (const r of (p.runs || [])) {
          if (r.fontFamily && r.fontFamily !== '+mj-lt' && r.fontFamily !== '+mn-lt') {
            families.add(this._safeFont(r.fontFamily).split(',')[0].trim().replace(/"/g, ''));
          }
        }
      }
    }
    try {
      const promises = [];
      for (const f of families) {
        ['400','700','800','900'].forEach(w => {
          promises.push(document.fonts.load(`${w} 16px "${f}"`).catch(() => {}));
        });
      }
      await Promise.all(promises);
    } catch {}
  },

  _decodeXml(str) {
    return str
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n)));
  }
};
