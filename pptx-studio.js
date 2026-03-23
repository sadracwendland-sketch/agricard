/* =====================================================
   AgriCard – PPTX Card Studio v1.0
   Sistema completo de geração de cards via PPTX

   FLUXO:
   1. Admin faz upload de PPTX → parser extrai slides
   2. Placeholders detectados: {{produtividade}}, {{fazenda}}, etc.
   3. Admin mapeia placeholders ↔ campos do formulário
   4. Usuário preenche dados (form / CSV upload)
   5. Canvas 2D renderiza pixel-perfect preservando layout
   6. Exportação: PNG alta-res, PDF ou lote via CSV

   PLACEHOLDERS SUPORTADOS:
   {{produtividade}}  {{unidade}}      {{fazenda}}
   {{produtor}}       {{cidade}}       {{estado}}
   {{variedade}}      {{cultura}}      {{safra}}
   {{data_plantio}}   {{data_colheita}} {{area}}
   {{tecnologia}}     {{populacao}}    {{notas}}
   {{logo_cliente}}   {{logo_marca}}
   ===================================================== */

'use strict';

const PptxStudio = {

  /* ──────── STATE ──────── */
  _currentFile:      null,    // File object
  _slides:           [],      // Array de SlideData parsed
  _selectedSlide:    null,    // SlideData selecionado
  _placeholderMap:   {},      // { placeholder: fieldId }
  _singleFormData:   {},      // dados do formulário manual
  _batchRows:        [],      // linhas do CSV/Excel
  _templateTitle:    '',
  _overflowMode:     'shrink', // 'shrink' | 'truncate'

  /* ──────── PLACEHOLDER REGISTRY ──────── */
  FIELD_REGISTRY: {
    produtividade:     { label: 'Produtividade',           icon: 'fa-chart-line',    example: '72,6' },
    produtividade_int: { label: 'Produtividade – Inteiro', icon: 'fa-chart-line',    example: '72',  derived: true },
    produtividade_dec: { label: 'Produtividade – Decimal', icon: 'fa-chart-line',    example: ',6',  derived: true },
    unidade:           { label: 'Unidade',                 icon: 'fa-balance-scale', example: 'sc/ha' },
    fazenda:        { label: 'Nome da Fazenda',  icon: 'fa-home',         example: 'Fazenda São João' },
    produtor:       { label: 'Nome do Produtor', icon: 'fa-user',         example: 'João da Silva' },
    cidade:         { label: 'Cidade',           icon: 'fa-map-marker-alt', example: 'Nova Lacerda' },
    estado:         { label: 'Estado (UF)',      icon: 'fa-flag',         example: 'MT' },
    variedade:      { label: 'Variedade',        icon: 'fa-leaf',         example: '79KA72' },
    cultura:        { label: 'Cultura',          icon: 'fa-seedling',     example: 'Soja' },
    safra:          { label: 'Safra / Estação',  icon: 'fa-calendar-alt', example: '2024/2025' },
    data_plantio:   { label: 'Data de Plantio',  icon: 'fa-calendar',     example: '15/10/2024' },
    data_colheita:  { label: 'Data de Colheita', icon: 'fa-calendar-check', example: '12/02/2025' },
    area:           { label: 'Área Colhida',     icon: 'fa-expand-arrows-alt', example: '25 ha' },
    tecnologia:     { label: 'Tecnologia',       icon: 'fa-flask',        example: 'Conkesta E3' },
    populacao:      { label: 'População',        icon: 'fa-layer-group',  example: '280.000 pl/ha' },
    notas:          { label: 'Observações',      icon: 'fa-sticky-note',  example: 'Irrigado' },
    logo_cliente:   { label: 'Logo Cliente (img)',  icon: 'fa-image',     example: '[URL imagem]' },
    logo_marca:     { label: 'Logo Marca (img)',    icon: 'fa-image',     example: '[URL imagem]' },
  },

  /* ═══════════════════════════════════════════════════════════════
     OPEN STUDIO MODAL
  ═══════════════════════════════════════════════════════════════ */
  openModal() {
    this._reset();
    const modal = document.getElementById('pptxStudioModal');
    if (modal) {
      modal.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
      this._setStep(1);
    }
  },

  closeModal() {
    document.getElementById('pptxStudioModal')?.classList.add('hidden');
    document.body.style.overflow = '';
  },

  _reset() {
    this._currentFile   = null;
    this._slides        = [];
    this._selectedSlide = null;
    this._placeholderMap = {};
    this._singleFormData = {};
    this._batchRows     = [];
    this._templateTitle = '';
    // Reset UI
    const dropzone = document.getElementById('studioDropzone');
    if (dropzone) dropzone.style.display = '';
    const slidesArea = document.getElementById('studioSlidesArea');
    if (slidesArea) slidesArea.style.display = 'none';
    const fileInput = document.getElementById('studioFileInput');
    if (fileInput) fileInput.value = '';
  },

  _setStep(step) {
    document.querySelectorAll('.studio-step').forEach(s => s.classList.remove('active'));
    const el = document.getElementById(`studioStep${step}`);
    if (el) el.classList.add('active');
    // update step indicators
    document.querySelectorAll('.step-indicator .step-dot').forEach((dot, idx) => {
      dot.classList.toggle('active', idx + 1 === step);
      dot.classList.toggle('done', idx + 1 < step);
    });
  },

  /* ═══════════════════════════════════════════════════════════════
     STEP 1 – UPLOAD & PARSE PPTX
  ═══════════════════════════════════════════════════════════════ */
  async handleUpload(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['pptx', 'pdf'].includes(ext)) {
      App.Toast.show('Use arquivos .pptx ou .pdf', 'error');
      return;
    }
    this._currentFile = file;

    // Show progress
    document.getElementById('studioDropzone').style.display = 'none';
    document.getElementById('studioParseProgress').style.display = 'flex';
    document.getElementById('studioParseStatus').textContent = 'Lendo arquivo...';

    try {
      if (ext === 'pptx') {
        await this._parsePptx(file);
      } else {
        await this._parsePdf(file);
      }
      this._renderSlideGrid();
      document.getElementById('studioParseProgress').style.display = 'none';
      document.getElementById('studioSlidesArea').style.display = '';
    } catch (err) {
      document.getElementById('studioParseProgress').style.display = 'none';
      document.getElementById('studioDropzone').style.display = '';
      App.Toast.show('Erro ao processar arquivo: ' + err.message, 'error');
      console.error(err);
    }
  },

  async _parsePptx(file) {
    const JSZip = await this._loadJSZip();
    const arrayBuf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuf);

    // Get slide list from presentation
    const presXml = await zip.file('ppt/presentation.xml')?.async('text');
    const slideCount = presXml ? (presXml.match(/<p:sldIdLst/g)?.length || 0) : 0;

    const slideFiles = Object.keys(zip.files)
      .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => {
        const na = parseInt(a.match(/\d+/)[0]);
        const nb = parseInt(b.match(/\d+/)[0]);
        return na - nb;
      });

    this._slides = [];
    const total = slideFiles.length;
    for (let i = 0; i < total; i++) {
      document.getElementById('studioParseStatus').textContent =
        `Processando slide ${i + 1} de ${total}...`;

      const slideXml = await zip.file(slideFiles[i]).async('text');
      const slideData = this._parseSlideXml(slideXml, i + 1);

      // Extract relationships to get images
      const relPath = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
      const relXml  = await zip.file(relPath)?.async('text') || '';
      slideData.rels = this._parseRels(relXml);

      // Try to get slide thumbnail image
      slideData.thumbDataUrl = await this._renderSlideToImage(slideData, zip, i);
      this._slides.push(slideData);
    }
  },

  _parseSlideXml(xml, index) {
    const parser = new DOMParser();
    const doc    = parser.parseFromString(xml, 'text/xml');

    const slide = {
      index,
      elements:     [],
      placeholders: [],
      bgColor:      '#FFFFFF',
      width:        9144000, // EMUs (default 10" = 9144000)
      height:       5143500,
    };

    // Background color
    const bgClr = doc.querySelector('bgPr solidFill srgbClr, bg solidFill srgbClr');
    if (bgClr) slide.bgColor = '#' + bgClr.getAttribute('val');

    // Dimensions from notes or default
    slide.width  = 9144000;
    slide.height = 5143500;

    // Parse all shape elements (sp)
    const shapes = doc.querySelectorAll('sp');
    shapes.forEach(sp => {
      const el = this._parseShape(sp);
      if (el) {
        slide.elements.push(el);
        // Check for placeholders
        const phs = el.rawText?.match(/\{\{(\w+)\}\}/g) || [];
        phs.forEach(ph => {
          const key = ph.replace(/[{}]/g, '').toLowerCase();
          if (!slide.placeholders.includes(key)) {
            slide.placeholders.push(key);
          }
        });
      }
    });

    // Parse picture elements (pic)
    const pics = doc.querySelectorAll('pic');
    pics.forEach(pic => {
      const el = this._parsePic(pic);
      if (el) slide.elements.push(el);
    });

    return slide;
  },

  _parseShape(sp) {
    // Position & size
    const xfrm  = sp.querySelector('spPr xfrm');
    if (!xfrm) return null;

    const off = xfrm.querySelector('off');
    const ext = xfrm.querySelector('ext');
    if (!off || !ext) return null;

    const x = parseInt(off.getAttribute('x') || 0);
    const y = parseInt(off.getAttribute('y') || 0);
    const w = parseInt(ext.getAttribute('cx') || 0);
    const h = parseInt(ext.getAttribute('cy') || 0);
    const rot = parseFloat(xfrm.getAttribute('rot') || 0) / 60000; // degrees

    // Text content
    const runs = sp.querySelectorAll('r');
    let rawText = '';
    const textRuns = [];

    runs.forEach(run => {
      const t = run.querySelector('t');
      if (!t) return;
      const text = t.textContent;
      rawText += text;

      // Font properties
      const rPr = run.querySelector('rPr');
      const pPr = run.closest('p')?.querySelector('pPr');

      // Font size in half-points → px (1pt = 1.333px at 96dpi)
      const szRaw = rPr?.getAttribute('sz') || '1800';
      const fontSizePt = parseInt(szRaw) / 100;

      // Bold/italic
      const bold   = rPr?.getAttribute('b') === '1';
      const italic = rPr?.getAttribute('i') === '1';

      // Font family
      const latinFont = rPr?.querySelector('latin');
      const fontFamily = latinFont?.getAttribute('typeface') || 'Inter, Arial';

      // Color
      let color = '#000000';
      const solidFill = rPr?.querySelector('solidFill srgbClr');
      if (solidFill) color = '#' + solidFill.getAttribute('val');
      const schemeFill = rPr?.querySelector('solidFill schemeClr');
      // fallback to common colors for scheme refs
      if (!solidFill && schemeFill) {
        const name = schemeFill.getAttribute('val');
        const schemeMap = { dk1: '#000000', lt1: '#FFFFFF', dk2: '#1F497D', lt2: '#EEECE1',
                            accent1: '#4F81BD', accent2: '#C0504D', accent3: '#9BBB59', accent4: '#8064A2',
                            accent5: '#4BACC6', accent6: '#F79646' };
        color = schemeMap[name] || '#333333';
      }

      // Paragraph alignment
      const algn = pPr?.getAttribute('algn') || 'l';
      const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
      const textAlign = alignMap[algn] || 'left';

      textRuns.push({ text, fontSizePt, bold, italic, fontFamily, color, textAlign });
    });

    // Paragraph-level alignment (fallback if no runs)
    let align = 'left';
    const pPrTop = sp.querySelector('pPr');
    if (pPrTop) {
      const algn = pPrTop.getAttribute('algn');
      const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
      align = alignMap[algn] || 'left';
    }

    // Text body margins
    const txBody = sp.querySelector('txBody bodyPr');
    const lIns  = parseInt(txBody?.getAttribute('lIns') || '91440') / 914.4;  // px
    const rIns  = parseInt(txBody?.getAttribute('rIns') || '91440') / 914.4;
    const tIns  = parseInt(txBody?.getAttribute('tIns') || '45720') / 914.4;
    const bIns  = parseInt(txBody?.getAttribute('bIns') || '45720') / 914.4;

    // Auto-fit
    const autoFit   = sp.querySelector('txBody bodyPr spAutoFit') !== null;
    const normAutoFit = sp.querySelector('txBody bodyPr normAutofit') !== null;

    // Fill
    let fillColor = 'transparent';
    const shapeFill = sp.querySelector('spPr solidFill srgbClr');
    if (shapeFill) fillColor = '#' + shapeFill.getAttribute('val');

    // Check if it's a placeholder shape (has phIdx or phType)
    const nvPr = sp.querySelector('nvPr ph');
    const isPhShape = !!nvPr;

    return {
      type:      'text',
      x, y, w, h, rot,
      rawText,
      textRuns,
      align,
      fillColor,
      lIns, rIns, tIns, bIns,
      autoFit, normAutoFit,
      isPhShape,
    };
  },

  _parsePic(pic) {
    const xfrm = pic.querySelector('spPr xfrm');
    if (!xfrm) return null;
    const off = xfrm.querySelector('off');
    const ext = xfrm.querySelector('ext');
    if (!off || !ext) return null;

    const blip = pic.querySelector('blip');
    const rEmbed = blip?.getAttribute('r:embed') || blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');

    return {
      type:    'image',
      x:       parseInt(off.getAttribute('x') || 0),
      y:       parseInt(off.getAttribute('y') || 0),
      w:       parseInt(ext.getAttribute('cx') || 0),
      h:       parseInt(ext.getAttribute('cy') || 0),
      rEmbed:  rEmbed || null,
      rawText: '',
    };
  },

  _parseRels(xml) {
    if (!xml) return {};
    const rels = {};
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    doc.querySelectorAll('Relationship').forEach(rel => {
      rels[rel.getAttribute('Id')] = rel.getAttribute('Target');
    });
    return rels;
  },

  // Render slide elements to a canvas image (thumbnail)
  async _renderSlideToImage(slide, zip, idx) {
    try {
      const THUMB_W = 480;
      const ASPECT  = slide.height / slide.width;
      const THUMB_H = Math.round(THUMB_W * ASPECT);

      const canvas  = document.createElement('canvas');
      canvas.width  = THUMB_W;
      canvas.height = THUMB_H;
      const ctx     = canvas.getContext('2d');

      const scaleX = THUMB_W / slide.width;
      const scaleY = THUMB_H / slide.height;

      // Background
      ctx.fillStyle = slide.bgColor || '#FFFFFF';
      ctx.fillRect(0, 0, THUMB_W, THUMB_H);

      // Draw elements in order
      for (const el of slide.elements) {
        const px = el.x * scaleX;
        const py = el.y * scaleY;
        const pw = el.w * scaleX;
        const ph = el.h * scaleY;

        if (el.type === 'text') {
          // Fill background
          if (el.fillColor && el.fillColor !== 'transparent') {
            ctx.fillStyle = el.fillColor;
            ctx.fillRect(px, py, pw, ph);
          }
          // Draw text runs
          this._drawTextRuns(ctx, el, px, py, pw, ph, scaleY, false);

        } else if (el.type === 'image' && el.rEmbed && zip) {
          const relPath = slide.rels?.[el.rEmbed];
          if (relPath) {
            const imgPath = 'ppt/slides/' + relPath.replace(/^\.*\//, '');
            const imgFile = zip.file(imgPath) || zip.file('ppt/' + relPath.replace(/^\.*\//, '').replace('slides/', ''));
            if (imgFile) {
              try {
                const blob    = await imgFile.async('blob');
                const dataUrl = await this._blobToDataUrl(blob);
                const img     = await this._loadImage(dataUrl);
                ctx.drawImage(img, px, py, pw, ph);
              } catch (e) { /* silently skip */ }
            }
          }
        }
      }

      return canvas.toDataURL('image/jpeg', 0.85);
    } catch (e) {
      console.warn('Thumb error:', e);
      return null;
    }
  },

  _drawTextRuns(ctx, el, px, py, pw, ph, scaleY, isPreview) {
    if (!el.textRuns?.length && !el.rawText) return;

    const runs = el.textRuns.length ? el.textRuns : [{
      text:       el.rawText,
      fontSizePt: 12,
      bold: false, italic: false,
      fontFamily: 'Inter, Arial',
      color:      '#333333',
      textAlign:  el.align || 'left',
    }];

    // group runs into paragraphs (split by newline)
    let curY = py + (el.tIns || 0) * scaleY;
    const lineH = (runs[0].fontSizePt || 12) * 1.333 * scaleY * 1.3;

    // Build paragraph groups
    const fullText = runs.map(r => r.text).join('');
    const lines = fullText.split('\n');
    let runIdx = 0;

    lines.forEach(line => {
      if (runIdx >= runs.length) runIdx = runs.length - 1;
      const run = runs[runIdx];
      runIdx++;

      const fsPx = (run.fontSizePt || 12) * 1.333 * scaleY;
      const font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${fsPx.toFixed(1)}px ${run.fontFamily || 'Inter, Arial'}`;
      ctx.font    = font;
      ctx.fillStyle = run.color || '#000000';

      const align = run.textAlign || el.align || 'left';
      ctx.textAlign = align === 'center' ? 'center' : align === 'right' ? 'right' : 'left';

      let drawX;
      if (align === 'center')  drawX = px + pw / 2;
      else if (align === 'right') drawX = px + pw - (el.rIns || 0) * scaleY;
      else drawX = px + (el.lIns || 0) * scaleY;

      // Overflow handling
      let displayText = line;
      const maxW = pw - (el.lIns + el.rIns || 0) * scaleY;
      if (ctx.measureText(line).width > maxW && maxW > 0) {
        if (this._overflowMode === 'truncate') {
          while (displayText.length > 0 && ctx.measureText(displayText + '…').width > maxW) {
            displayText = displayText.slice(0, -1);
          }
          displayText += '…';
        } else {
          // shrink: reduce scale
          const ratio = maxW / ctx.measureText(line).width;
          const newFs = fsPx * ratio;
          ctx.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${newFs.toFixed(1)}px ${run.fontFamily || 'Inter, Arial'}`;
        }
      }

      ctx.fillText(displayText, drawX, curY + fsPx);
      curY += lineH;
    });
  },

  async _parsePdf(file) {
    const pdfjsLib = await this._loadPdfJs();
    const arrayBuf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuf }).promise;

    this._slides = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      document.getElementById('studioParseStatus').textContent =
        `Processando página ${i} de ${pdf.numPages}...`;

      const page  = await pdf.getPage(i);
      const vp    = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas');
      canvas.width  = vp.width;
      canvas.height = vp.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;

      const dataUrl = canvas.toDataURL('image/jpeg', 0.85);
      this._slides.push({
        index:        i,
        elements:     [],
        placeholders: [],
        bgColor:      '#FFFFFF',
        width:        vp.width,
        height:       vp.height,
        thumbDataUrl: dataUrl,
        isPdf:        true,
      });
    }
  },

  /* ═══════════════════════════════════════════════════════════════
     STEP 1 UI – SLIDE GRID
  ═══════════════════════════════════════════════════════════════ */
  _renderSlideGrid() {
    const grid = document.getElementById('studioSlidesGrid');
    if (!grid) return;

    const allPh = new Set();
    this._slides.forEach(s => s.placeholders?.forEach(p => allPh.add(p)));

    grid.innerHTML = this._slides.map((slide, i) => {
      const phBadges = slide.placeholders?.map(p =>
        `<span class="ph-badge">${p}</span>`).join('') || '';
      const thumb = slide.thumbDataUrl
        ? `<img src="${slide.thumbDataUrl}" alt="Slide ${slide.index}" loading="lazy" />`
        : `<div class="slide-thumb-placeholder"><i class="fas fa-file-image"></i></div>`;
      return `
        <div class="studio-slide-card ${i === 0 ? 'selected' : ''}"
             data-idx="${i}" onclick="PptxStudio.selectSlide(${i})">
          <div class="slide-thumb">${thumb}</div>
          <div class="slide-info">
            <div class="slide-num">Slide ${slide.index}</div>
            <div class="slide-ph-list">${phBadges || '<span class="ph-none">Sem placeholders</span>'}</div>
          </div>
        </div>`;
    }).join('');

    // Auto-select first
    if (this._slides.length > 0) this.selectSlide(0);

    // Show placeholder summary
    const summaryEl = document.getElementById('studioPhSummary');
    if (summaryEl) {
      if (allPh.size > 0) {
        summaryEl.innerHTML = `<i class="fas fa-check-circle" style="color:var(--green)"></i>
          Detectados <strong>${allPh.size}</strong> placeholders: 
          ${[...allPh].map(p => `<code>{{${p}}}</code>`).join(', ')}`;
      } else {
        summaryEl.innerHTML = `<i class="fas fa-info-circle" style="color:var(--orange)"></i>
          Nenhum placeholder detectado. Você pode adicionar manualmente como <code>{{produtividade}}</code> no slide.`;
      }
    }
  },

  selectSlide(idx) {
    this._selectedSlide = this._slides[idx];
    document.querySelectorAll('.studio-slide-card').forEach((c, i) => {
      c.classList.toggle('selected', i === idx);
    });
    // Update "next" button state
    const btnNext = document.getElementById('studioNextStep1');
    if (btnNext) btnNext.disabled = false;
  },

  proceedToMapping() {
    if (!this._selectedSlide) {
      App.Toast.show('Selecione um slide primeiro.', 'warning');
      return;
    }
    this._buildMappingUI();
    this._setStep(2);
  },

  /* ═══════════════════════════════════════════════════════════════
     STEP 2 – PLACEHOLDER MAPPING
  ═══════════════════════════════════════════════════════════════ */
  _buildMappingUI() {
    const slide = this._selectedSlide;
    const container = document.getElementById('studioMappingContainer');
    if (!container) return;

    // Get all unique placeholders from slide
    const phs = slide.placeholders || [];

    // Also scan all text for {{...}} patterns
    const allTexts = slide.elements.map(e => e.rawText || '').join(' ');
    const detected = [...new Set([
      ...phs,
      ...(allTexts.match(/\{\{(\w+)\}\}/g) || []).map(p => p.replace(/[{}]/g, '').toLowerCase())
    ])];

    if (detected.length === 0) {
      container.innerHTML = `
        <div class="mapping-empty">
          <i class="fas fa-info-circle"></i>
          <p>Nenhum placeholder detectado no slide selecionado.</p>
          <p>Adicione placeholders como <code>{{produtividade}}</code> no seu PPTX e reimporte.</p>
          <p style="margin-top:12px"><strong>Placeholders disponíveis:</strong></p>
          <div class="ph-registry">
            ${Object.entries(this.FIELD_REGISTRY).map(([key, val]) =>
              `<code onclick="navigator.clipboard?.writeText('{{${key}}}')" title="Clique para copiar">{{${key}}}</code>`
            ).join('')}
          </div>
        </div>`;
      return;
    }

    // Initialize map with auto-detected matches
    detected.forEach(ph => {
      if (!this._placeholderMap[ph] && this.FIELD_REGISTRY[ph]) {
        this._placeholderMap[ph] = ph; // auto-map same name
      }
    });

    const fieldOptions = Object.entries(this.FIELD_REGISTRY).map(([key, val]) =>
      `<option value="${key}">${val.derived ? '↳ ' : ''}${val.label} ({{${key}}})</option>`
    ).join('');

    container.innerHTML = `
      <div class="mapping-header">
        <p class="mapping-desc">
          <i class="fas fa-link"></i>
          Mapeie cada placeholder do PPTX ao campo de dado correspondente.
          Placeholders com mesmo nome são mapeados automaticamente.
        </p>
      </div>
      <div class="mapping-grid">
        ${detected.map(ph => {
          const reg = this.FIELD_REGISTRY[ph];
          const isAuto = !!reg;
          const currentVal = this._placeholderMap[ph] || '';
          return `
            <div class="mapping-row ${isAuto ? 'auto-mapped' : ''}">
              <div class="mapping-ph">
                <code>{{${ph}}}</code>
                ${isAuto ? '<span class="auto-tag"><i class="fas fa-magic"></i> automático</span>' : ''}
              </div>
              <div class="mapping-arrow"><i class="fas fa-long-arrow-alt-right"></i></div>
              <div class="mapping-field">
                <select id="map_${ph}" onchange="PptxStudio._placeholderMap['${ph}'] = this.value">
                  <option value="">— Ignorar este campo —</option>
                  ${fieldOptions}
                </select>
              </div>
              ${reg ? `<div class="mapping-example"><small>Ex: ${reg.example}</small></div>` : ''}
            </div>`;
        }).join('')}
      </div>
      <div class="overflow-config">
        <label class="section-label"><i class="fas fa-text-width"></i> Comportamento quando texto excede a caixa:</label>
        <div class="radio-group">
          <label class="radio-label ${this._overflowMode === 'shrink' ? 'active' : ''}">
            <input type="radio" name="overflowMode" value="shrink"
              ${this._overflowMode === 'shrink' ? 'checked' : ''}
              onchange="PptxStudio._overflowMode='shrink';this.closest('.radio-group').querySelectorAll('.radio-label').forEach(l=>l.classList.toggle('active',l.querySelector('input').checked))">
            <i class="fas fa-compress-alt"></i> Reduzir fonte proporcionalmente
          </label>
          <label class="radio-label ${this._overflowMode === 'truncate' ? 'active' : ''}">
            <input type="radio" name="overflowMode" value="truncate"
              ${this._overflowMode === 'truncate' ? 'checked' : ''}
              onchange="PptxStudio._overflowMode='truncate';this.closest('.radio-group').querySelectorAll('.radio-label').forEach(l=>l.classList.toggle('active',l.querySelector('input').checked))">
            <i class="fas fa-ellipsis-h"></i> Truncar com "..."
          </label>
        </div>
      </div>`;

    // Set select values
    setTimeout(() => {
      detected.forEach(ph => {
        const sel = document.getElementById(`map_${ph}`);
        if (sel && this._placeholderMap[ph]) {
          sel.value = this._placeholderMap[ph];
        }
      });
    }, 50);
  },

  proceedToData() {
    this._buildDataEntryUI();
    this._setStep(3);
  },

  /* ═══════════════════════════════════════════════════════════════
     STEP 3 – DATA ENTRY (FORM or CSV)
  ═══════════════════════════════════════════════════════════════ */
  _buildDataEntryUI() {
    // Build form fields based on mapped placeholders
    const mapped = Object.entries(this._placeholderMap).filter(([, v]) => v);

    const formContainer = document.getElementById('studioDataForm');
    if (!formContainer) return;

    // Get unique target fields — excluindo os campos derivados (gerados automaticamente)
    const DERIVED = ['produtividade_int', 'produtividade_dec'];
    const fields = [...new Set(mapped.map(([, v]) => v))].filter(f => !DERIVED.includes(f));

    if (fields.length === 0) {
      formContainer.innerHTML = `<p class="info-msg">
        <i class="fas fa-info-circle"></i> Nenhum campo mapeado. Volte ao passo anterior.
      </p>`;
      return;
    }

    // Verifica se há placeholders derivados mapeados para exibir aviso informativo
    const hasDerived = mapped.some(([, v]) => DERIVED.includes(v));
    const derivedNote = hasDerived ? `
      <div class="derived-note">
        <i class="fas fa-cut"></i>
        <div>
          <strong>Separação automática ativa:</strong>
          Ao preencher <em>Produtividade</em> (ex: <code>187,1</code>), o sistema vai gerar
          automaticamente <code>{{produtividade_int}}</code> = <strong>187</strong>
          e <code>{{produtividade_dec}}</code> = <strong>,1</strong> — sem ação extra.
        </div>
      </div>` : '';

    formContainer.innerHTML = `
      ${derivedNote}
      <div class="data-form-grid">
        ${fields.map(fieldKey => {
          const reg = this.FIELD_REGISTRY[fieldKey] || { label: fieldKey, icon: 'fa-tag', example: '' };
          const isImage = fieldKey.startsWith('logo');
          return `
            <div class="form-group">
              <label><i class="fas ${reg.icon}"></i> ${reg.label}</label>
              ${isImage
                ? `<div class="logo-upload-row">
                     <input type="url" id="field_${fieldKey}" class="studio-field"
                       placeholder="URL da imagem ou..." value="${this._singleFormData[fieldKey] || ''}"
                       oninput="PptxStudio._singleFormData['${fieldKey}']=this.value;PptxStudio._schedulePreview()" />
                     <button type="button" class="btn-sm btn-secondary"
                       onclick="document.getElementById('logoUpload_${fieldKey}').click()">
                       <i class="fas fa-upload"></i>
                     </button>
                     <input type="file" id="logoUpload_${fieldKey}" accept="image/*" style="display:none"
                       onchange="PptxStudio._handleLogoUpload('${fieldKey}',this.files[0])" />
                   </div>`
                : `<input type="text" id="field_${fieldKey}" class="studio-field"
                     placeholder="${reg.example}"
                     value="${this._singleFormData[fieldKey] || ''}"
                     oninput="PptxStudio._singleFormData['${fieldKey}']=this.value;PptxStudio._schedulePreview()" />`
              }
            </div>`;
        }).join('')}
      </div>`;
  },

  async _handleLogoUpload(fieldKey, file) {
    if (!file) return;
    const dataUrl = await this._blobToDataUrl(file);
    this._singleFormData[fieldKey] = dataUrl;
    const input = document.getElementById(`field_${fieldKey}`);
    if (input) input.value = '(imagem carregada)';
    this._schedulePreview();
  },

  /* ─── CSV UPLOAD ─── */
  async handleCsvUpload(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['csv', 'xlsx', 'xls'].includes(ext)) {
      App.Toast.show('Use arquivos .csv, .xlsx ou .xls', 'error');
      return;
    }

    try {
      if (ext === 'csv') {
        const text = await file.text();
        this._batchRows = this._parseCsv(text);
      } else {
        this._batchRows = await this._parseExcel(file);
      }
      this._showCsvPreview();
    } catch (err) {
      App.Toast.show('Erro ao ler arquivo: ' + err.message, 'error');
      console.error(err);
    }
  },

  _parseCsv(text) {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return [];

    const headers = lines[0].split(/[,;]/).map(h => h.trim().replace(/^"|"$/g, '').toLowerCase());
    return lines.slice(1).map(line => {
      const values = line.split(/[,;]/).map(v => v.trim().replace(/^"|"$/g, ''));
      const row = {};
      headers.forEach((h, i) => { row[h] = values[i] || ''; });
      return row;
    });
  },

  async _parseExcel(file) {
    // Load SheetJS (XLSX library)
    if (!window.XLSX) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    }
    const arrayBuf = await file.arrayBuffer();
    const wb  = XLSX.read(arrayBuf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(row[i] || ''); });
      return obj;
    });
  },

  _showCsvPreview() {
    const area = document.getElementById('studioCsvPreview');
    if (!area) return;

    const rows  = this._batchRows;
    const limit = Math.min(rows.length, 5);
    const headers = rows.length ? Object.keys(rows[0]) : [];

    area.innerHTML = `
      <div class="csv-preview-header">
        <i class="fas fa-table"></i>
        <strong>${rows.length} registros carregados</strong> – exibindo os primeiros ${limit}
      </div>
      <div class="table-wrapper" style="max-height:180px;overflow-y:auto">
        <table class="data-table" style="font-size:12px">
          <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
          <tbody>
            ${rows.slice(0, limit).map(row =>
              `<tr>${headers.map(h => `<td>${row[h] || ''}</td>`).join('')}</tr>`
            ).join('')}
          </tbody>
        </table>
      </div>`;
  },

  /* ═══════════════════════════════════════════════════════════════
     STEP 4 – PREVIEW & GENERATE
  ═══════════════════════════════════════════════════════════════ */
  _previewTimer: null,
  _schedulePreview() {
    clearTimeout(this._previewTimer);
    this._previewTimer = setTimeout(() => this.renderPreview(), 400);
  },

  async renderPreview() {
    const slide = this._selectedSlide;
    if (!slide) return;

    const previewCanvas = document.getElementById('studioPreviewCanvas');
    if (!previewCanvas) return;

    const data = this._buildDataObject(this._singleFormData);
    await SlideRenderer.renderSlide(
      previewCanvas, slide, data,
      this._placeholderMap,
      this._overflowMode,
      this._currentFile
    );
  },

  _buildDataObject(formData) {
    // Merge placeholder → field → value
    const data = {};
    Object.entries(this._placeholderMap).forEach(([ph, fieldKey]) => {
      if (fieldKey && formData[fieldKey] !== undefined) {
        data[ph] = formData[fieldKey];
      }
    });

    // ── Derivar produtividade_int e produtividade_dec ──────────────
    // Fonte: campo "produtividade" resolvido em data OU diretamente em formData
    const prodRaw = data['produtividade'] ?? formData['produtividade'] ?? '';
    this._deriveProdutividade(prodRaw, data);

    return data;
  },

  /**
   * Divide produtividade no formato BR ("187,1") em:
   *   produtividade_int → "187"
   *   produtividade_dec → ",1"   (vazio se sem decimal)
   *
   * Escreve pelo nome canônico do fieldKey E pelo nome do placeholder
   * mapeado, garantindo que {{produtividade_int}} e {{meu_alias}}
   * funcionem independentemente.
   */
  _deriveProdutividade(raw, data) {
    const str = String(raw || '').trim();
    const [inteiro, decimal] = str.split(',');
    const prodInt = inteiro || '';
    const prodDec = decimal !== undefined ? ',' + decimal : '';

    // Escreve pelas chaves canônicas (fieldKey)
    data['produtividade_int'] = prodInt;
    data['produtividade_dec'] = prodDec;

    // Também escreve por qualquer placeholder mapeado para produtividade_int / _dec
    Object.entries(this._placeholderMap).forEach(([ph, fieldKey]) => {
      if (fieldKey === 'produtividade_int') data[ph] = prodInt;
      if (fieldKey === 'produtividade_dec') data[ph] = prodDec;
    });
  },

  /**
   * Garante derivação em linhas de lote (CSV/Excel) após _mapCsvRow().
   */
  _deriveProdutividadeFromRow(rowMapped) {
    const str = String(rowMapped['produtividade'] || '').trim();
    const [inteiro, decimal] = str.split(',');
    rowMapped['produtividade_int'] = inteiro || '';
    rowMapped['produtividade_dec'] = decimal !== undefined ? ',' + decimal : '';
    return rowMapped;
  },

  proceedToPreview() {
    this._setStep(4);
    this.renderPreview();
  },

  /* ═══════════════════════════════════════════════════════════════
     GENERATE & EXPORT
  ═══════════════════════════════════════════════════════════════ */
  async generateSingle(format = 'png') {
    const slide = this._selectedSlide;
    if (!slide) { App.Toast.show('Nenhum slide selecionado.', 'error'); return; }

    const btn = document.getElementById('btnStudioGenerate');
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="loading"></span> Gerando...'; }

    try {
      const data = this._buildDataObject(this._singleFormData);
      const canvas = document.createElement('canvas');
      // High-res: 300 DPI equivalent (≈ 3x screen resolution)
      const SCALE = 3;
      canvas.width  = Math.round(slide.width  / (914400 / 96) * SCALE); // EMU → px at 96dpi × scale
      canvas.height = Math.round(slide.height / (914400 / 96) * SCALE);

      await SlideRenderer.renderSlide(
        canvas, slide, data,
        this._placeholderMap,
        this._overflowMode,
        this._currentFile,
        SCALE
      );

      const filename = this._buildFilename(data, 'card');

      if (format === 'png') {
        this._downloadCanvas(canvas, filename + '.png', 'image/png', 1.0);
      } else if (format === 'jpeg') {
        this._downloadCanvas(canvas, filename + '.jpg', 'image/jpeg', 0.95);
      } else if (format === 'pdf') {
        await this._exportPdf(canvas, filename + '.pdf');
      }

      App.Toast.show('Card gerado com sucesso!', 'success');
    } catch (err) {
      App.Toast.show('Erro ao gerar: ' + err.message, 'error');
      console.error(err);
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-download"></i> Baixar'; }
    }
  },

  async generateBatch() {
    if (!this._batchRows.length) {
      App.Toast.show('Carregue um arquivo CSV/Excel primeiro.', 'warning');
      return;
    }

    const slide = this._selectedSlide;
    if (!slide) { App.Toast.show('Nenhum slide selecionado.', 'error'); return; }

    const btn = document.getElementById('btnBatchGenerate');
    if (btn) { btn.disabled = true; }

    const progressBar  = document.getElementById('batchProgressBar');
    const progressText = document.getElementById('batchProgressText');
    const batchArea    = document.getElementById('batchProgressArea');
    if (batchArea) batchArea.style.display = '';

    const total   = this._batchRows.length;
    const canvases = [];
    const SCALE    = 3;

    for (let i = 0; i < total; i++) {
      const rowRaw = this._batchRows[i];
      // Map CSV column names → field keys (try exact match and common synonyms)
      let rowMapped = this._mapCsvRow(rowRaw);
      // Garante produtividade_int e produtividade_dec para o lote
      rowMapped = this._deriveProdutividadeFromRow(rowMapped);
      const data = this._buildDataObject(rowMapped);

      const canvas = document.createElement('canvas');
      canvas.width  = Math.round(slide.width  / (914400 / 96) * SCALE);
      canvas.height = Math.round(slide.height / (914400 / 96) * SCALE);

      await SlideRenderer.renderSlide(
        canvas, slide, data,
        this._placeholderMap,
        this._overflowMode,
        this._currentFile,
        SCALE
      );

      canvases.push({ canvas, data, index: i + 1 });

      // Update progress
      const pct = Math.round(((i + 1) / total) * 100);
      if (progressBar)  progressBar.style.width = pct + '%';
      if (progressText) progressText.textContent = `Processando ${i + 1} / ${total}...`;

      // Yield to UI
      await new Promise(r => setTimeout(r, 10));
    }

    if (progressText) progressText.textContent = 'Empacotando ZIP...';

    // Export ZIP
    await this._exportBatchZip(canvases);
    if (batchArea) batchArea.style.display = 'none';
    if (btn) btn.disabled = false;
    App.Toast.show(`${total} cards gerados com sucesso!`, 'success');
  },

  _mapCsvRow(row) {
    // Map CSV column names to field keys
    const synonyms = {
      produtividade: ['produtividade', 'productivity', 'prod', 'producao'],
      unidade:       ['unidade', 'unit', 'un'],
      fazenda:       ['fazenda', 'farm', 'nome_fazenda'],
      produtor:      ['produtor', 'producer', 'nome_produtor', 'nome'],
      cidade:        ['cidade', 'city', 'municipio'],
      estado:        ['estado', 'state', 'uf'],
      variedade:     ['variedade', 'variety', 'cultivar'],
      cultura:       ['cultura', 'culture', 'crop'],
      safra:         ['safra', 'season', 'ano_safra'],
      data_plantio:  ['data_plantio', 'planting_date', 'plantio'],
      data_colheita: ['data_colheita', 'harvest_date', 'colheita'],
      area:          ['area', 'area_colhida', 'ha'],
      tecnologia:    ['tecnologia', 'technology', 'tech'],
      populacao:     ['populacao', 'population', 'pop'],
      notas:         ['notas', 'notes', 'obs', 'observacoes'],
    };

    const mapped = {};
    Object.entries(synonyms).forEach(([fieldKey, keys]) => {
      for (const k of keys) {
        const found = Object.keys(row).find(rk => rk.toLowerCase().replace(/\s+/g, '_') === k);
        if (found !== undefined) {
          mapped[fieldKey] = row[found];
          break;
        }
      }
    });

    return mapped;
  },

  _downloadCanvas(canvas, filename, mimeType, quality) {
    canvas.toBlob(blob => {
      const url = URL.createObjectURL(blob);
      const a   = document.createElement('a');
      a.href     = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    }, mimeType, quality);
  },

  async _exportPdf(canvas, filename) {
    // Load jsPDF
    if (!window.jspdf) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js');
    }
    const { jsPDF } = window.jspdf;
    const imgData  = canvas.toDataURL('image/jpeg', 0.95);
    const pdf = new jsPDF({
      orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
      unit:        'px',
      format:      [canvas.width, canvas.height],
      compress:    true,
    });
    pdf.addImage(imgData, 'JPEG', 0, 0, canvas.width, canvas.height);
    pdf.save(filename);
  },

  async _exportBatchZip(canvases) {
    const JSZip = await this._loadJSZip();
    const zip   = new JSZip();
    const folder = zip.folder('cards');

    for (const { canvas, data, index } of canvases) {
      const blob = await new Promise(res => canvas.toBlob(res, 'image/jpeg', 0.95));
      const name = this._buildFilename(data, `card_${String(index).padStart(3, '0')}`) + '.jpg';
      folder.file(name, blob);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    const url  = URL.createObjectURL(zipBlob);
    const link = document.createElement('a');
    link.href     = url;
    link.download = 'cards_lote.zip';
    link.click();
    URL.revokeObjectURL(url);
  },

  _buildFilename(data, prefix) {
    const parts = [
      prefix,
      data['variedade'] || data['fazenda'] || '',
      data['cidade']    || '',
      data['safra']     || new Date().getFullYear(),
    ].filter(Boolean).join('_').replace(/[^a-z0-9_\-\.]/gi, '_').replace(/__+/g, '_');
    return parts.substring(0, 80);
  },

  /* ═══════════════════════════════════════════════════════════════
     HELPERS
  ═══════════════════════════════════════════════════════════════ */
  async _loadJSZip() {
    if (window.JSZip) return window.JSZip;
    await this._loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    return window.JSZip;
  },

  async _loadPdfJs() {
    if (window.pdfjsLib) return window.pdfjsLib;
    await this._loadScript('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js');
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    return window.pdfjsLib;
  },

  _loadScript(url) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${url}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src  = url;
      s.onload = resolve;
      s.onerror = () => reject(new Error('Failed to load: ' + url));
      document.head.appendChild(s);
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

  _loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload  = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });
  },

  /* ═══════════════════════════════════════════════════════════════
     CSV TEMPLATE DOWNLOAD
  ═══════════════════════════════════════════════════════════════ */
  downloadCsvTemplate() {
    const headers = Object.keys(this.FIELD_REGISTRY).filter(k => !k.startsWith('logo'));
    const example = {
      produtividade: '72.6',
      unidade:       'sc/ha',
      fazenda:       'Fazenda São João',
      produtor:      'João da Silva',
      cidade:        'Nova Lacerda',
      estado:        'MT',
      variedade:     '79KA72',
      cultura:       'Soja',
      safra:         '2024/2025',
      data_plantio:  '15/10/2024',
      data_colheita: '12/02/2025',
      area:          '25 ha',
      tecnologia:    'Conkesta E3',
      populacao:     '280.000 pl/ha',
      notas:         'Irrigado',
    };

    const csv = [
      headers.join(';'),
      headers.map(h => example[h] || '').join(';'),
      headers.map(h => example[h] ? example[h] + '_2' : '').join(';'),
    ].join('\n');

    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href     = url;
    link.download = 'modelo_cards_lote.csv';
    link.click();
    URL.revokeObjectURL(url);
    App.Toast.show('Modelo CSV baixado!', 'success');
  },
};
