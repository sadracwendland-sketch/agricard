/* =====================================================
   AgriCard Stine - File Import Module
   Suporte a upload de PPTX e PDF como modelo de card
   ===================================================== */

const FileImporter = {
  _slides: [],          // array de { dataUrl, index, label }
  _currentFile: null,
  _mode: null,          // 'pdf' | 'pptx'

  // ===================================================
  // OPEN IMPORT MODAL
  // ===================================================
  openModal() {
    this._slides = [];
    this._currentFile = null;
    this._mode = null;
    const modal = document.getElementById('importFileModal');
    if (modal) {
      modal.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
      this._resetUI();
    }
  },

  closeModal() {
    const modal = document.getElementById('importFileModal');
    if (modal) modal.classList.add('hidden');
    document.body.style.overflow = '';
  },

  _resetUI() {
    const dropzone   = document.getElementById('importDropzone');
    const progress   = document.getElementById('importProgress');
    const slidesArea = document.getElementById('importSlidesArea');
    const actionArea = document.getElementById('importActionArea');
    if (dropzone)   dropzone.style.display   = '';
    if (progress)   progress.style.display   = 'none';
    if (slidesArea) slidesArea.style.display  = 'none';
    if (actionArea) actionArea.style.display  = 'none';
    const fileInput = document.getElementById('importFileInput');
    if (fileInput) fileInput.value = '';
    const grid = document.getElementById('importSlidesGrid');
    if (grid) grid.innerHTML = '';
    const info = document.getElementById('importFileInfo');
    if (info) info.textContent = '';
  },

  // ===================================================
  // FILE SELECTED
  // ===================================================
  async handleFile(file) {
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!['pdf', 'pptx'].includes(ext)) {
      App.Toast.show('Formato não suportado. Use PDF ou PPTX.', 'error');
      return;
    }

    this._currentFile = file;
    this._mode = ext;

    const info = document.getElementById('importFileInfo');
    if (info) info.textContent = `📄 ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;

    document.getElementById('importDropzone').style.display = 'none';
    this._showProgress('Lendo arquivo...');

    try {
      if (ext === 'pdf') {
        await this._processPDF(file);
      } else {
        await this._processPPTX(file);
      }
    } catch (err) {
      console.error('FileImporter error:', err);
      App.Toast.show('Erro ao processar arquivo: ' + (err.message || err), 'error');
      this._resetUI();
    }
  },

  // ===================================================
  // PROCESS PDF — usa PDF.js
  // ===================================================
  async _processPDF(file) {
    this._showProgress('Carregando PDF.js...');

    // Carrega PDF.js dinamicamente se ainda não carregado
    if (!window.pdfjsLib) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js');
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    }

    this._showProgress('Processando páginas do PDF...');

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    this._slides = [];
    const total = pdf.numPages;

    for (let i = 1; i <= total; i++) {
      this._showProgress(`Renderizando página ${i} de ${total}...`);
      const page   = await pdf.getPage(i);
      const scale  = 1.5;
      const vp     = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width  = vp.width;
      canvas.height = vp.height;
      const ctx = canvas.getContext('2d');
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      this._slides.push({
        dataUrl: canvas.toDataURL('image/jpeg', 0.88),
        index: i,
        label: `Página ${i}`
      });
    }

    this._renderSlides();
  },

  // ===================================================
  // PROCESS PPTX — usa JSZip para extrair thumbnails
  // ===================================================
  async _processPPTX(file) {
    this._showProgress('Carregando JSZip...');

    if (!window.JSZip) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    }

    this._showProgress('Descompactando PPTX...');

    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    this._slides = [];

    // Estratégia 1: thumbnails oficiais (pasta ppt/media ou docProps/thumbnails)
    const thumbFolder = zip.folder('ppt/media');
    const docThumb    = zip.file('docProps/thumbnail.jpeg') ||
                        zip.file('docProps/thumbnail.jpg')  ||
                        zip.file('docProps/thumbnail.png');

    // Estratégia 2: imagens de fundo dos slides
    const mediaFiles = [];
    zip.forEach((path, f) => {
      if (!f.dir && path.startsWith('ppt/media/') &&
          /\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(path)) {
        mediaFiles.push({ path, file: f });
      }
    });

    // Estratégia 3: renderizar cada slide como SVG/canvas usando pptx2svg (fallback)
    // Tenta extrair relações de slides para ordenar as imagens
    const slideRels = {};
    zip.forEach((path, f) => {
      const m = path.match(/ppt\/slides\/_rels\/slide(\d+)\.xml\.rels$/);
      if (m) slideRels[parseInt(m[1])] = f;
    });

    // Coleta slides em ordem
    const slideFiles = [];
    zip.forEach((path, f) => {
      const m = path.match(/ppt\/slides\/slide(\d+)\.xml$/);
      if (m) slideFiles.push({ num: parseInt(m[1]), path, file: f });
    });
    slideFiles.sort((a, b) => a.num - b.num);

    if (slideFiles.length === 0) {
      throw new Error('Não foi possível encontrar slides no arquivo PPTX.');
    }

    this._showProgress(`Processando ${slideFiles.length} slides...`);

    // Para cada slide, tenta encontrar a imagem principal
    for (let si = 0; si < slideFiles.length; si++) {
      const sf    = slideFiles[si];
      const relF  = slideRels[sf.num];
      let imgDataUrl = null;

      if (relF) {
        const relXml = await relF.async('text');
        // Pega todas as imagens referenciadas neste slide
        const imgRefs = [...relXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)].map(m => m[1]);

        for (const imgRef of imgRefs) {
          const imgFile = zip.file(`ppt/media/${imgRef}`);
          if (imgFile && /\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(imgRef)) {
            const blob = await imgFile.async('blob');
            imgDataUrl = await this._blobToDataUrl(blob);
            break; // usa primeira imagem do slide
          }
        }
      }

      // Fallback: renderiza o slide como canvas via SVG simplificado
      if (!imgDataUrl) {
        imgDataUrl = await this._renderSlideAsCanvas(zip, sf);
      }

      if (imgDataUrl) {
        this._slides.push({
          dataUrl: imgDataUrl,
          index: sf.num,
          label: `Slide ${sf.num}`
        });
      }
    }

    // Se não conseguiu nenhuma imagem dos slides, usa thumbnail geral
    if (this._slides.length === 0 && docThumb) {
      const blob = await docThumb.async('blob');
      const url  = await this._blobToDataUrl(blob);
      this._slides.push({ dataUrl: url, index: 1, label: 'Capa do arquivo' });
    }

    if (this._slides.length === 0) {
      throw new Error('Nenhuma imagem encontrada nos slides. Tente exportar como PDF.');
    }

    this._renderSlides();
  },

  // Renderiza um slide PPTX como canvas baseado nas formas/textos (simplificado)
  async _renderSlideAsCanvas(zip, slideFile) {
    try {
      const slideXml = await slideFile.file.async('text');

      // Extrai informações básicas de cor de fundo
      const bgColorMatch = slideXml.match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
      const bgColor = bgColorMatch ? '#' + bgColorMatch[1] : '#1a3a1a';

      // Cria canvas com cor de fundo do slide
      const canvas = document.createElement('canvas');
      canvas.width  = 960;
      canvas.height = 540;
      const ctx = canvas.getContext('2d');

      // Fundo
      ctx.fillStyle = bgColor;
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Overlay gradiente sutil
      const grad = ctx.createLinearGradient(0, 0, canvas.width, canvas.height);
      grad.addColorStop(0, 'rgba(0,0,0,0.2)');
      grad.addColorStop(1, 'rgba(0,0,0,0.0)');
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Textos simples extraídos do XML
      const texts = [...slideXml.matchAll(/<a:t>([^<]{3,})<\/a:t>/g)]
        .map(m => m[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').trim())
        .filter(t => t.length > 2)
        .slice(0, 10);

      ctx.fillStyle = 'rgba(255,255,255,0.9)';
      ctx.textAlign = 'center';

      let y = 140;
      texts.forEach((text, i) => {
        const size = i === 0 ? 42 : i <= 2 ? 28 : 18;
        ctx.font = `${i <= 1 ? '700' : '400'} ${size}px Arial`;
        ctx.fillStyle = i === 0 ? 'rgba(255,255,255,1)' : 'rgba(255,255,255,0.8)';
        ctx.fillText(text.substring(0, 60), canvas.width / 2, y);
        y += size + 14;
        if (y > canvas.height - 60) return;
      });

      // Label do slide
      ctx.font = '12px Arial';
      ctx.fillStyle = 'rgba(255,255,255,0.4)';
      ctx.textAlign = 'right';
      ctx.fillText(`Slide ${slideFile.num}`, canvas.width - 20, canvas.height - 20);

      return canvas.toDataURL('image/jpeg', 0.85);
    } catch (e) {
      return null;
    }
  },

  // ===================================================
  // RENDER SLIDES GRID
  // ===================================================
  _renderSlides() {
    const slidesArea = document.getElementById('importSlidesArea');
    const actionArea = document.getElementById('importActionArea');
    const grid       = document.getElementById('importSlidesGrid');
    const info       = document.getElementById('importSlidesInfo');

    if (!grid) return;

    document.getElementById('importProgress').style.display = 'none';
    if (slidesArea) slidesArea.style.display = '';
    if (actionArea) actionArea.style.display = '';
    if (info) info.textContent = `${this._slides.length} ${this._mode === 'pdf' ? 'página(s)' : 'slide(s)'} encontrado(s). Clique para selecionar o modelo desejado.`;

    grid.innerHTML = this._slides.map((s, i) => `
      <div class="import-slide-card" id="importSlide_${i}" onclick="FileImporter.selectSlide(${i})">
        <div class="import-slide-thumb">
          <img src="${s.dataUrl}" alt="${s.label}" loading="lazy" />
          <div class="import-slide-overlay">
            <i class="fas fa-check-circle"></i>
          </div>
        </div>
        <div class="import-slide-label">${s.label}</div>
      </div>
    `).join('');
  },

  // ===================================================
  // SELECT SLIDE
  // ===================================================
  _selectedSlideIndex: null,

  selectSlide(index) {
    this._selectedSlideIndex = index;
    // Marca visualmente
    document.querySelectorAll('.import-slide-card').forEach((el, i) => {
      el.classList.toggle('selected', i === index);
    });
    const useBtn = document.getElementById('btnUseSelectedSlide');
    if (useBtn) useBtn.disabled = false;
  },

  // ===================================================
  // USE SELECTED SLIDE AS BACKGROUND
  // ===================================================
  useSelected() {
    if (this._selectedSlideIndex === null || !this._slides[this._selectedSlideIndex]) {
      App.Toast.show('Selecione um slide/página primeiro.', 'error');
      return;
    }

    const slide = this._slides[this._selectedSlideIndex];
    const dataUrl = slide.dataUrl;

    // Preenche o campo de URL de fundo do template
    const bgInput = document.getElementById('tplBgImageUrl');
    if (bgInput) {
      bgInput.value = dataUrl;
    }

    // Atualiza o nome do template com sugestão
    const nameInput = document.getElementById('tplName');
    if (nameInput && !nameInput.value) {
      const fileName = this._currentFile?.name?.replace(/\.(pptx|pdf)$/i, '') || 'Importado';
      nameInput.value = `${fileName} – ${slide.label}`;
    }

    // Atualiza o preview do template
    TemplatesManager.updatePreview();

    // Fecha o modal de importação
    this.closeModal();

    App.Toast.show(`✅ ${slide.label} aplicado como fundo do template!`, 'success');
  },

  // ===================================================
  // USE ALL AS SEPARATE TEMPLATES
  // ===================================================
  async useAllAsSeparate() {
    if (this._slides.length === 0) {
      App.Toast.show('Nenhum slide carregado.', 'error');
      return;
    }

    const baseName = this._currentFile?.name?.replace(/\.(pptx|pdf)$/i,'') || 'Importado';
    const btn = document.getElementById('btnUseAllSlides');
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="loading"></span> Criando...'; }

    let created = 0;
    for (const s of this._slides) {
      try {
        await API.createTemplate({
          name:             `${baseName} – ${s.label}`,
          description:      `Importado de ${this._currentFile?.name || 'arquivo'}`,
          layout_type:      'imported',
          header_color:     '#2E7D32',
          header_text:      'RESULTADOS DE PRODUTIVIDADE',
          slogan:           'NÃO É SORTE! É STINE',
          footer_logo_text: 'STINE',
          bg_image_url:     s.dataUrl,
          badge_label:      '',
          show_ranking_badge: false,
          active:           false,
          sort_order:       99
        });
        created++;
      } catch (e) {
        console.error('Erro ao criar template:', e);
      }
    }

    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-layer-group"></i> Criar Todos como Templates'; }

    CardGenerator._templates = [];
    this.closeModal();
    await TemplatesManager.loadTemplates();
    App.Toast.show(`✅ ${created} template(s) criado(s) com sucesso!`, 'success');
  },

  // ===================================================
  // DRAG & DROP
  // ===================================================
  setupDrop() {
    const zone = document.getElementById('importDropzone');
    if (!zone) return;

    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('dragover');
      const file = e.dataTransfer?.files?.[0];
      if (file) this.handleFile(file);
    });
  },

  // ===================================================
  // HELPERS
  // ===================================================
  _showProgress(msg) {
    const progress = document.getElementById('importProgress');
    const text     = document.getElementById('importProgressText');
    if (progress) progress.style.display = '';
    if (text)     text.textContent = msg;
  },

  _blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  },

  _loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src;
      s.onload  = resolve;
      s.onerror = () => reject(new Error(`Falha ao carregar: ${src}`));
      document.head.appendChild(s);
    });
  }
};
