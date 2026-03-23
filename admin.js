/* =====================================================
   AgriCard Stine – Admin Module v4.0
   ===================================================== */

const Admin = {
  userFilter: 'all',
  _allRecords: [],
  _allRecordsRaw: [],

  async init() {
    await this.loadDashboard();
    this.setupFilterButtons();
  },

  /* ═══════════════════════════════════════════════════
     DASHBOARD
  ═══════════════════════════════════════════════════ */
  async loadDashboard() {
    try {
      const [usersRes, recordsRes, varietiesRes] = await Promise.all([
        API.getUsers(),
        API.getRecords(),
        API.getVarieties()
      ]);

      const users     = usersRes.data     || [];
      const records   = recordsRes.data   || [];
      const varieties = varietiesRes.data || [];

      const approved = users.filter(u => u.status === 'approved' && u.role !== 'admin');
      const pending  = users.filter(u => u.status === 'pending');

      document.getElementById('statTotalUsers').textContent     = approved.length;
      document.getElementById('statPendingUsers').textContent   = pending.length;
      document.getElementById('statTotalRecords').textContent   = records.length;
      document.getElementById('statTotalVarieties').textContent = varieties.length;

      const badge = document.getElementById('navBadgePending');
      if (badge) {
        badge.style.display = pending.length ? '' : 'none';
        badge.textContent   = pending.length;
      }

      const alertBox    = document.getElementById('pendingAlert');
      const pendingList = document.getElementById('pendingUsersList');
      if (pending.length > 0) {
        alertBox.style.display = 'block';
        pendingList.innerHTML = pending.map(u => `
          <div class="pending-user-card">
            <div class="pending-user-info">
              <strong>${this.esc(u.name)}</strong>
              <span>${this.esc(u.email)} · ${this.esc(u.company || 'Sem empresa')}</span>
            </div>
            <div class="pending-user-actions">
              <button class="action-btn action-btn-green" onclick="Admin.approveUser('${u.id}')">
                <i class="fas fa-check"></i> Aprovar
              </button>
              <button class="action-btn action-btn-red" onclick="Admin.rejectUser('${u.id}')">
                <i class="fas fa-times"></i> Rejeitar
              </button>
            </div>
          </div>
        `).join('');
      } else {
        alertBox.style.display = 'none';
      }

      const recentEl = document.getElementById('recentRecordsList');
      const recent   = records
        .slice().sort((a, b) => (b.created_at || 0) - (a.created_at || 0))
        .slice(0, 8);

      if (recent.length === 0) {
        recentEl.innerHTML = `<div class="empty-state">
          <i class="fas fa-chart-bar"></i>
          <p>Nenhum registro de produtividade ainda</p>
        </div>`;
      } else {
        recentEl.innerHTML = `<table class="data-table">
          <thead><tr>
            <th>Variedade</th><th>Produtor</th><th>Cidade/UF</th>
            <th>Produtividade</th><th>Safra</th><th>Ações</th>
          </tr></thead>
          <tbody>
          ${recent.map(r => `
            <tr>
              <td>
                <strong>${this.esc(r.variety_name || '-')}</strong>
                <br><small style="color:var(--gray-500)">${this.esc(r.brand || '')}</small>
              </td>
              <td>${this.esc(r.producer_name || '-')}</td>
              <td>${this.esc(r.city || '-')}/${this.esc(r.state || '-')}</td>
              <td>
                <strong style="color:var(--green);font-size:15px">
                  ${parseFloat(r.productivity || 0).toLocaleString('pt-BR', {minimumFractionDigits:1, maximumFractionDigits:1})}
                </strong>
                <small>${this.esc(r.unit || '')}</small>
              </td>
              <td>${this.esc(r.season || '-')}</td>
              <td>
                <button class="action-btn action-btn-blue" onclick="Admin.previewRecord('${r.id}')">
                  <i class="fas fa-image"></i> Card
                </button>
              </td>
            </tr>
          `).join('')}
          </tbody>
        </table>`;
      }

    } catch (err) {
      console.error('Admin.loadDashboard error:', err);
    }
  },

  /* ═══════════════════════════════════════════════════
     USERS
  ═══════════════════════════════════════════════════ */
  async loadUsers() {
    try {
      const res = await API.getUsers();
      const all = (res.data || []).filter(u => u.role !== 'admin');
      this.renderUsers(all, this.userFilter);
    } catch (err) {
      console.error(err);
    }
  },

  renderUsers(users, filter) {
    const filtered  = filter === 'all' ? users : users.filter(u => u.status === filter);
    const container = document.getElementById('usersTableContainer');

    if (filtered.length === 0) {
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-users"></i>
        <p>Nenhum usuário encontrado</p>
      </div>`;
      return;
    }

    container.innerHTML = `<table class="data-table">
      <thead><tr>
        <th>Nome</th><th>E-mail</th><th>Empresa</th>
        <th>Região</th><th>Status</th><th>Ações</th>
      </tr></thead>
      <tbody>
      ${filtered.map(u => `
        <tr>
          <td><strong>${this.esc(u.name)}</strong></td>
          <td>${this.esc(u.email)}</td>
          <td>${this.esc(u.company || '-')}</td>
          <td>${this.esc(u.region || '-')}</td>
          <td>${this.statusBadge(u.status)}</td>
          <td>
            <div class="table-actions">
              ${u.status !== 'approved'
                ? `<button class="action-btn action-btn-green" onclick="Admin.approveUser('${u.id}')"><i class="fas fa-check"></i> Aprovar</button>`
                : ''}
              ${u.status !== 'rejected'
                ? `<button class="action-btn action-btn-red" onclick="Admin.rejectUser('${u.id}')"><i class="fas fa-times"></i> Rejeitar</button>`
                : ''}
              <button class="action-btn action-btn-red" onclick="Admin.deleteUser('${u.id}')">
                <i class="fas fa-trash"></i>
              </button>
            </div>
          </td>
        </tr>
      `).join('')}
      </tbody>
    </table>`;
  },

  setupFilterButtons() {
    document.querySelectorAll('#userFilterTabs .filter-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#userFilterTabs .filter-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        this.userFilter = btn.dataset.filter;
        this.loadUsers();
      });
    });
  },

  async approveUser(id) {
    try {
      await API.updateUser(id, { status: 'approved' });
      App.Toast.show('Usuário aprovado com sucesso!', 'success');
      await Promise.all([this.loadDashboard(), this.loadUsers()]);
    } catch {
      App.Toast.show('Erro ao aprovar usuário.', 'error');
    }
  },

  async rejectUser(id) {
    App.confirm('Rejeitar este usuário?', async () => {
      try {
        await API.updateUser(id, { status: 'rejected' });
        App.Toast.show('Usuário rejeitado.', 'warning');
        await Promise.all([this.loadDashboard(), this.loadUsers()]);
      } catch {
        App.Toast.show('Erro ao rejeitar.', 'error');
      }
    });
  },

  async deleteUser(id) {
    App.confirm('Excluir este usuário permanentemente?', async () => {
      try {
        await API.deleteUser(id);
        App.Toast.show('Usuário excluído.', 'info');
        await this.loadUsers();
      } catch {
        App.Toast.show('Erro ao excluir.', 'error');
      }
    });
  },

  /* ═══════════════════════════════════════════════════
     VARIETIES
  ═══════════════════════════════════════════════════ */
  async loadVarieties() {
    const grid = document.getElementById('varietiesGrid');
    if (!grid) return;

    grid.innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:24px;color:#888">
      <span class="loading"></span> Carregando variedades…
    </div>`;

    try {
      const res       = await API.getVarieties();
      const varieties = res.data || [];

      if (varieties.length === 0) {
        grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
          <i class="fas fa-leaf"></i>
          <p>Nenhuma variedade cadastrada. Clique em "Nova Variedade" para começar.</p>
        </div>`;
        return;
      }

      grid.innerHTML = varieties.map(v => {
        const hasModel  = !!v.template_image;
        const hasCoords = !!v.field_coords;
        const color     = v.primary_color || '#2E7D32';

        return `
        <div class="variety-card" style="border-top-color:${color}">

          <!-- Thumbnail do modelo -->
          ${hasModel ? `
            <div class="variety-card-template-thumb" onclick="Admin.openVarietyModal('${v.id}')">
              <img src="${this._thumbSrc(v.template_image)}" alt="modelo"
                style="width:100%;height:100%;object-fit:cover;display:block"
                onerror="this.parentElement.innerHTML='<div style=\'display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#888;font-size:12px\'><i class=\'fas fa-image\' style=\'font-size:24px;margin-bottom:4px\'></i>Modelo</div>'" />
              <div class="variety-card-thumb-overlay">
                <i class="fas fa-pencil-alt"></i> Editar
              </div>
            </div>
          ` : `
            <div class="variety-card-no-template" onclick="Admin.openVarietyModal('${v.id}')">
              <i class="fas fa-upload" style="font-size:24px;margin-bottom:6px;color:#bbb"></i>
              <span style="font-size:12px;color:#aaa">Adicionar modelo</span>
            </div>
          `}

          <div class="variety-card-header">
            <div>
              <div class="variety-brand">${this.esc(v.brand || '')} · ${this.esc(v.culture || '')}</div>
              <div class="variety-code">${this.esc(v.name)}</div>
            </div>
            <div style="width:18px;height:18px;border-radius:50%;background:${color};flex-shrink:0"></div>
          </div>

          ${v.technology ? `<span class="variety-tech">${this.esc(v.technology)}</span>` : ''}
          ${v.maturity_group ? `<span class="variety-tech" style="background:var(--blue-light);color:var(--blue-mid);margin-left:4px">GM ${this.esc(v.maturity_group)}</span>` : ''}

          <div class="variety-template-status" style="margin-top:8px;font-size:11px">
            ${hasModel
              ? `<span class="tpl-status-ok"><i class="fas fa-check-circle"></i> Modelo vinculado</span>`
              : `<span class="tpl-status-missing"><i class="fas fa-exclamation-triangle"></i> Sem modelo de card</span>`}
            ${(() => {
              if (!v.pptx_elements) return hasModel
                ? `<span class="tpl-status-missing" style="margin-left:4px"><i class="fas fa-crosshairs"></i> Sem layout PPTX</span>`
                : '';
              try {
                const els = JSON.parse(v.pptx_elements);
                const phs = els.filter(e => e.placeholder);
                return `<span class="tpl-status-ok" style="margin-left:4px">
                  <i class="fas fa-magic"></i> ${phs.length} placeholder(s) PPTX</span>`;
              } catch { return ''; }
            })()}
          </div>

          ${v.description ? `<p style="font-size:11px;color:var(--gray-500);margin-top:8px">${this.esc(v.description)}</p>` : ''}

          <div class="variety-actions">
            <button class="action-btn action-btn-blue" onclick="Admin.openVarietyModal('${v.id}')">
              <i class="fas fa-edit"></i> Editar
            </button>
            ${hasModel ? `
            <button class="action-btn action-btn-orange" onclick="Admin.openVarietyCalibrator('${v.id}')" title="Calibrar posições dos campos">
              <i class="fas fa-crosshairs"></i> Calibrar
            </button>
            ` : ''}
            <button class="action-btn action-btn-red" onclick="Admin.deleteVariety('${v.id}')">
              <i class="fas fa-trash"></i>
            </button>
          </div>
        </div>
      `}).join('');

    } catch (err) {
      console.error(err);
      grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
        <i class="fas fa-exclamation-circle"></i>
        <p>Erro ao carregar variedades</p>
      </div>`;
    }
  },

  /* Retorna uma versão reduzida do src para uso no img[src] sem explodir o HTML */
  _thumbSrc(src) {
    if (!src) return '';
    // base64 data URL — usa como está (o browser consegue renderizar)
    if (src.startsWith('data:')) return src;
    return this.esc(src);
  },

  /* ─── MODAL DE VARIEDADE ─── */
  openVarietyModal(id = null) {
    const modal = document.getElementById('varietyModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';

    this._resetTemplateUpload();
    this._currentVarietyId = id;

    if (id) {
      API.getVarieties().then(res => {
        const v = (res.data || []).find(x => x.id === id);
        if (!v) return;

        this._setField('varName',        v.name          || '');
        this._setField('varBrand',       v.brand         || '');
        this._setField('varTechnology',  v.technology    || '');
        this._setField('varCulture',     v.culture       || '');
        this._setField('varMaturity',    v.maturity_group || '');
        this._setField('varLogoUrl',     v.logo_url      || '');
        this._setField('varDescription', v.description   || '');

        const colorEl = document.getElementById('varColor');
        const hexEl   = document.getElementById('varColorHex');
        if (colorEl) colorEl.value       = v.primary_color || '#2E7D32';
        if (hexEl)   hexEl.textContent   = v.primary_color || '#2E7D32';

        const titleEl = document.getElementById('varietyModalTitle');
        if (titleEl) titleEl.innerHTML = '<i class="fas fa-edit"></i> Editar Variedade';

        const saveBtn = document.getElementById('btnSaveVariety');
        if (saveBtn) saveBtn.onclick = () => Admin.saveVariety(id);

        this._currentTemplateImage = v.template_image || null;
        if (v.template_image) {
          this._showTemplatePreview(v.template_image, v.name ? `Modelo: ${v.name}` : 'Modelo atual');
        }

        // Mostra status dos placeholders PPTX se já processado
        const status = document.getElementById('varTemplateStatus');
        if (status && v.pptx_elements) {
          try {
            const els = JSON.parse(v.pptx_elements);
            const phs = [...new Set(els.filter(e => e.placeholder).map(e => `{{${e.placeholder}}}` ))];
            if (phs.length > 0) {
              status.innerHTML = `<i class="fas fa-check-circle" style="color:var(--green)"></i>
                ${phs.length} placeholder(s): <strong>${phs.join(', ')}</strong>`;
            }
          } catch { /* ignore */ }
        }
      });
    } else {
      ['varName','varBrand','varTechnology','varCulture','varMaturity','varLogoUrl','varDescription']
        .forEach(fid => this._setField(fid, ''));
      const colorEl = document.getElementById('varColor');
      const hexEl   = document.getElementById('varColorHex');
      if (colorEl) colorEl.value     = '#2E7D32';
      if (hexEl)   hexEl.textContent = '#2E7D32';

      const titleEl = document.getElementById('varietyModalTitle');
      if (titleEl) titleEl.innerHTML = '<i class="fas fa-leaf"></i> Nova Variedade';

      const saveBtn = document.getElementById('btnSaveVariety');
      if (saveBtn) saveBtn.onclick = () => Admin.saveVariety(null);

      this._currentTemplateImage = null;
    }
  },

  _setField(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value;
  },

  _currentVarietyId:      null,
  _currentTemplateImage:  null,
  _newTemplateImageData:  null,
  _newPptxElements:       null,   // JSON string com elementos extraídos do PPTX
  _newPptxSlideW:         null,   // largura do slide em EMU
  _newPptxSlideH:         null,   // altura do slide em EMU
  _newLogoData:           null,   // base64 do logo extraído do PPTX

  _resetTemplateUpload() {
    this._newTemplateImageData = null;
    this._newPptxElements      = null;
    this._newPptxSlideW        = null;
    this._newPptxSlideH        = null;
    this._newLogoData          = null;
    const dropzone   = document.getElementById('varTemplateDropzone');
    const preview    = document.getElementById('varTemplatePreview');
    const status     = document.getElementById('varTemplateStatus');
    const inp        = document.getElementById('varTemplateFileInput');
    const logoPreview = document.getElementById('varLogoPreview');
    const logoImg    = document.getElementById('varLogoPreviewImg');
    if (dropzone)    dropzone.style.display    = '';
    if (preview)     { preview.style.display   = 'none'; const img = preview.querySelector('img'); if(img) img.src=''; }
    if (status)      status.textContent         = '';
    if (inp)         inp.value                  = '';
    if (logoPreview) logoPreview.style.display  = 'none';
    if (logoImg)     logoImg.src                = '';
  },

  _showTemplatePreview(src, label = '') {
    const dropzone = document.getElementById('varTemplateDropzone');
    const preview  = document.getElementById('varTemplatePreview');
    if (!preview) return;
    const img = preview.querySelector('img');
    if (img)   img.src = src;
    const lbl = preview.querySelector('.var-tpl-label');
    if (lbl)   lbl.textContent = label || 'Modelo carregado';
    if (dropzone) dropzone.style.display = 'none';
    preview.style.display = '';
  },

  /* ─── UPLOAD DE ARQUIVO DE MODELO ─── */
  async handleTemplateFile(file) {
    if (!file) return;
    const ext    = file.name.split('.').pop().toLowerCase();
    const status = document.getElementById('varTemplateStatus');
    const setStatus = t => { if (status) status.textContent = t; };

    if (['png','jpg','jpeg','webp','gif'].includes(ext)) {
      setStatus('Carregando imagem…');
      const reader = new FileReader();
      reader.onload = e => {
        this._newTemplateImageData = e.target.result;
        this._showTemplatePreview(e.target.result, file.name);
        setStatus(`✅ ${file.name} pronto`);
        CardGenerator._invalidateCache(this._currentVarietyId);
      };
      reader.readAsDataURL(file);

    } else if (ext === 'pdf') {
      setStatus('Processando PDF…');
      try {
        const dataUrl = await this._pdfFirstPage(file);
        this._newTemplateImageData = dataUrl;
        this._showTemplatePreview(dataUrl, file.name + ' (página 1)');
        setStatus(`✅ ${file.name} pronto`);
        CardGenerator._invalidateCache(this._currentVarietyId);
      } catch (e) {
        setStatus('❌ Erro ao processar PDF: ' + e.message);
        App.Toast.show('Erro ao processar PDF: ' + e.message, 'error');
      }

    } else if (ext === 'pptx') {
      setStatus('⏳ Processando PPTX… (renderizando slide completo com logo preservada)');
      try {
        // PptxParser v9: renderiza o slide completo (fundo + logos + ícones) como PNG
        // Logo já estará composta no template_image — sem reinjeção posterior
        const parsed = await PptxParser.parseFile(file);

        // fullSlideImage = PNG do slide completo (fundo + logos em transparência correta)
        let previewSrc = parsed.fullSlideImage || parsed.bgImageData || parsed.thumbnailData;
        if (!previewSrc) throw new Error('Não foi possível renderizar o slide do PPTX.');

        // O template_image é o PNG completo (já inclui logo com transparência)
        // NÃO comprimir para JPEG se houver transparência.
        // Se o PNG for muito grande (>1.5MB base64), usa JPEG como fallback aceitável.
        setStatus('🖼 Comprimindo imagem do slide…');
        previewSrc = await this._compressImagePng(previewSrc, 1080);
        if (previewSrc.length > 1500000) {
          previewSrc = await this._compressImagePng(previewSrc, 900);
        }
        if (previewSrc.length > 1500000) {
          previewSrc = await this._compressImagePng(previewSrc, 720);
        }
        // Último recurso: converte para JPEG se ainda > 1.5MB
        // (slides com muitos gradientes ficam grandes mesmo em PNG)
        if (previewSrc.length > 1500000) {
          console.warn('[Admin] PNG ainda muito grande após 3 compressões — convertendo para JPEG');
          previewSrc = await this._compressImage(previewSrc, 1080, 0.88);
        }

        // Salva o slide completo como template_image
        this._newTemplateImageData = previewSrc;
        this._newLogoData          = null;  // Logo já está embutida no template_image

        // Salva elementos PPTX (apenas texto)
        this._newPptxElements = JSON.stringify(parsed.elements);
        this._newPptxSlideW   = parsed.slideW;
        this._newPptxSlideH   = parsed.slideH;

        // Esconde preview do logo (não há mais logo separado)
        const logoPreview = document.getElementById('varLogoPreview');
        if (logoPreview) logoPreview.style.display = 'none';

        this._showTemplatePreview(previewSrc, `${file.name} (slide completo renderizado)`);

        // Conta placeholders de texto detectados
        const allPhs = [];
        parsed.elements.forEach(e => {
          if (e.placeholder) allPhs.push(`{{${e.placeholder}}}`);
          if (e.placeholders) e.placeholders.forEach(p => allPhs.push(`{{${p}}}`));
        });
        const phList  = [...new Set(allPhs)].join(', ');
        const phCount = [...new Set(allPhs)].length;

        if (phCount > 0) {
          setStatus(`✅ ${file.name} — ${phCount} placeholder(s) de texto: ${phList} | 🖼 Logo e imagens preservadas no layout`);
        } else {
          setStatus(`⚠️ Slide renderizado, mas nenhum placeholder de texto detectado. Verifique os nomes das caixas no PPTX (use {{campo}}).`);
        }

        CardGenerator._invalidateCache(this._currentVarietyId);
      } catch (e) {
        setStatus('❌ Erro ao processar PPTX: ' + e.message);
        App.Toast.show('Erro ao processar PPTX: ' + e.message, 'error');
        console.error(e);
      }

    } else {
      App.Toast.show('Formato não suportado. Use PNG, JPG, PDF ou PPTX.', 'error');
    }
  },

  /* ─── PDF → imagem via PDF.js ─── */
  async _pdfFirstPage(file) {
    if (!window.pdfjsLib) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js');
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    }
    const ab   = await file.arrayBuffer();
    const pdf  = await window.pdfjsLib.getDocument({ data: ab }).promise;
    const page = await pdf.getPage(1);
    const vp   = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement('canvas');
    canvas.width  = vp.width;
    canvas.height = vp.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    return canvas.toDataURL('image/jpeg', 0.92);
  },

  /* ─── PPTX → imagem via JSZip ─── */
  async _pptxFirstSlide(file) {
    if (!window.JSZip) {
      await this._loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    }
    const ab  = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(ab);

    // Tenta thumbnail oficial do PPTX
    for (const name of ['docProps/thumbnail.jpeg','docProps/thumbnail.jpg','docProps/thumbnail.png']) {
      const f = zip.file(name);
      if (f) return this._blobToDataUrl(await f.async('blob'));
    }

    // Tenta primeira imagem do slide 1
    const rel1 = zip.file('ppt/slides/_rels/slide1.xml.rels');
    if (rel1) {
      const xml = await rel1.async('text');
      const m   = xml.match(/Target="\.\.\/media\/([^"]+\.(png|jpg|jpeg|gif|bmp|webp))"/i);
      if (m) {
        const imgFile = zip.file(`ppt/media/${m[1]}`);
        if (imgFile) return this._blobToDataUrl(await imgFile.async('blob'));
      }
    }

    // Fallback canvas com textos do slide 1
    const slide1 = zip.file('ppt/slides/slide1.xml');
    if (slide1) {
      const xml   = await slide1.async('text');
      const bg    = (xml.match(/<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/) || [])[1];
      const texts = [...xml.matchAll(/<a:t>([^<]{2,})<\/a:t>/g)]
        .map(m => m[1].replace(/&amp;/g,'&').trim()).filter(Boolean).slice(0, 8);
      const c = document.createElement('canvas');
      c.width = 720; c.height = 1280;
      const x = c.getContext('2d');
      x.fillStyle = bg ? '#' + bg : '#1a3a1a';
      x.fillRect(0, 0, c.width, c.height);
      x.fillStyle = 'white'; x.textAlign = 'center';
      let y = 240;
      texts.forEach((t, i) => {
        const sz = i === 0 ? 64 : i <= 2 ? 40 : 24;
        x.font = `${i <= 1 ? '900' : '400'} ${sz}px Arial`;
        x.fillText(t.substring(0, 40), c.width / 2, y);
        y += sz + 20;
      });
      return c.toDataURL('image/jpeg', 0.92);
    }

    throw new Error('Nenhuma imagem encontrada no PPTX. Tente exportar como PDF primeiro.');
  },

  _blobToDataUrl(blob) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => res(e.target.result);
      r.onerror = rej;
      r.readAsDataURL(blob);
    });
  },

  /**
   * Comprime uma imagem base64 para JPEG com largura máxima e qualidade definidas.
   * Garante que o resultado fique abaixo de ~700KB de base64 (≈512KB de dados).
   * @param {string} dataUrl - base64 original (png/jpeg/etc)
   * @param {number} maxW    - largura máxima em px (default 1080)
   * @param {number} quality - qualidade JPEG 0-1 (default 0.82)
   * @returns {Promise<string>} - base64 JPEG comprimida
   */
  async _compressImage(dataUrl, maxW = 1080, quality = 0.82) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        let w = img.naturalWidth;
        let h = img.naturalHeight;
        // Redimensiona se necessário
        if (w > maxW) {
          h = Math.round(h * maxW / w);
          w = maxW;
        }
        const c = document.createElement('canvas');
        c.width  = w;
        c.height = h;
        const ctx = c.getContext('2d');
        ctx.drawImage(img, 0, 0, w, h);
        const result = c.toDataURL('image/jpeg', quality);
        resolve(result);
      };
      img.onerror = reject;
      img.src = dataUrl;
    });
  },

  /**
   * Redimensiona imagem mantendo PNG (preserva transparência/alpha).
   * Diferente de _compressImage que converte para JPEG e perde transparência.
   * @param {string} dataUrl - base64 original (qualquer formato)
   * @param {number} maxW    - largura máxima em px
   * @returns {Promise<string>} - base64 PNG redimensionado
   */
  async _compressImagePng(dataUrl, maxW = 1080) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        let w = img.naturalWidth;
        let h = img.naturalHeight;
        if (w > maxW) {
          h = Math.round(h * maxW / w);
          w = maxW;
        }
        const c = document.createElement('canvas');
        c.width  = w;
        c.height = h;
        const ctx = c.getContext('2d');
        // Limpa com transparência (não preto)
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        // PNG preserva canal alpha
        resolve(c.toDataURL('image/png'));
      };
      img.onerror = reject;
      img.src = dataUrl;
    });
  },

  _loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error(`Falha ao carregar: ${src}`));
      document.head.appendChild(s);
    });
  },

  closeVarietyModal() {
    const modal = document.getElementById('varietyModal');
    if (modal) modal.classList.add('hidden');
    document.body.style.overflow = '';
  },

  async saveVariety(id) {
    // Logo URL: não salva se for base64 ou placeholder interno
    const rawLogoUrl = (document.getElementById('varLogoUrl')?.value || '').trim();
    const logoUrlClean = (rawLogoUrl.startsWith('data:') || rawLogoUrl === '(logo extraído do PPTX)')
      ? '' : rawLogoUrl;

    const data = {
      name:           (document.getElementById('varName')?.value        || '').trim(),
      brand:          (document.getElementById('varBrand')?.value       || '').trim(),
      technology:     (document.getElementById('varTechnology')?.value  || '').trim(),
      culture:        (document.getElementById('varCulture')?.value     || '').trim(),
      maturity_group: (document.getElementById('varMaturity')?.value    || '').trim(),
      primary_color:  document.getElementById('varColor')?.value         || '#2E7D32',
      logo_url:       logoUrlClean,
      description:    (document.getElementById('varDescription')?.value || '').trim()
    };

    if (this._newTemplateImageData) {
      // v9: template_image = slide completo (fundo + logos fixos como PNG)
      data.template_image = this._newTemplateImageData;
      // Elementos PPTX (apenas texto — sem logo_image separado)
      if (this._newPptxElements !== null) {
        data.pptx_elements = this._newPptxElements;
        data.pptx_slide_w  = this._newPptxSlideW;
        data.pptx_slide_h  = this._newPptxSlideH;
      }
      // v9: pptx_logo não é mais usado — logo está embutida no template_image
      // Limpa campo legado se existia
      data.pptx_logo = '';
    } else if (id && this._currentTemplateImage) {
      data.template_image = this._currentTemplateImage;
    }

    if (!data.name || !data.brand || !data.culture) {
      App.Toast.show('Preencha Nome, Marca e Cultura.', 'error');
      return;
    }
    if (!id && !data.template_image) {
      App.Toast.show('Faça o upload do modelo de card para esta variedade.', 'error');
      return;
    }

    const btn = document.getElementById('btnSaveVariety');
    const origText = btn?.innerHTML;
    if (btn) { btn.innerHTML = '<span class="loading"></span> Salvando…'; btn.disabled = true; }

    try {
      // Log sizes para debug
      const templateSize = data.template_image ? Math.round(data.template_image.length / 1024) : 0;
      const elemsSize    = data.pptx_elements  ? Math.round(data.pptx_elements.length  / 1024) : 0;
      console.log(`[saveVariety v9] template_image (slide completo)=${templateSize}KB | pptx_elements=${elemsSize}KB`);

      if (id) {
        await API.updateVariety(id, data);
        App.Toast.show('Variedade atualizada!', 'success');
      } else {
        await API.createVariety(data);
        App.Toast.show('Variedade cadastrada!', 'success');
      }
      this.closeVarietyModal();
      CardGenerator._invalidateCache(id);
      await this.loadVarieties();
    } catch (err) {
      console.error('[saveVariety] Erro:', err);
      App.Toast.show('Erro ao salvar variedade: ' + (err.message || 'Tente novamente.'), 'error');
    } finally {
      if (btn) { btn.innerHTML = origText; btn.disabled = false; }
    }
  },

  async deleteVariety(id) {
    App.confirm('Excluir esta variedade? Esta ação não pode ser desfeita.', async () => {
      try {
        await API.deleteVariety(id);
        App.Toast.show('Variedade excluída.', 'info');
        CardGenerator._invalidateCache(id);
        await this.loadVarieties();
      } catch {
        App.Toast.show('Erro ao excluir.', 'error');
      }
    });
  },

  /* ─── CALIBRADOR ─── */
  async openVarietyCalibrator(id) {
    try {
      const res = await API.getVarieties();
      const v   = (res.data || []).find(x => x.id === id);
      if (!v) { App.Toast.show('Variedade não encontrada.', 'error'); return; }

      CardGenerator.openCalibrator(v, async (newCoords) => {
        try {
          await API.updateVariety(id, { field_coords: JSON.stringify(newCoords) });
          CardGenerator._invalidateCache(id);
          App.Toast.show('Calibração salva!', 'success');
          await this.loadVarieties();
        } catch {
          App.Toast.show('Erro ao salvar calibração.', 'error');
        }
      });
    } catch {
      App.Toast.show('Erro ao abrir calibrador.', 'error');
    }
  },

  /* ═══════════════════════════════════════════════════
     ALL RECORDS
  ═══════════════════════════════════════════════════ */
  async loadAllRecords() {
    try {
      const res = await API.getRecords();
      this._allRecordsRaw = (res.data || []).sort((a, b) => (b.created_at || 0) - (a.created_at || 0));
      this.renderAllRecords(this._allRecordsRaw);
    } catch (err) {
      console.error(err);
    }
  },

  filterRecords(q) {
    if (!q) { this.renderAllRecords(this._allRecordsRaw); return; }
    const lower = q.toLowerCase();
    this.renderAllRecords(
      this._allRecordsRaw.filter(r =>
        [r.variety_name, r.producer_name, r.farm_name, r.city, r.state, r.brand, r.season]
          .some(v => v && v.toLowerCase().includes(lower))
      )
    );
  },

  renderAllRecords(records) {
    const container = document.getElementById('allRecordsContainer');

    if (!records || records.length === 0) {
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-chart-bar"></i>
        <p>Nenhum registro encontrado</p>
      </div>`;
      return;
    }

    container.innerHTML = `<table class="data-table">
      <thead><tr>
        <th>Usuário</th><th>Variedade</th><th>Produtor</th><th>Fazenda</th>
        <th>Cidade/UF</th><th>Produtividade</th><th>Safra</th><th>Termo</th><th>Status</th><th>Ações</th>
      </tr></thead>
      <tbody>
      ${records.map(r => `
        <tr>
          <td style="font-size:12px">${this.esc(r.user_name || '-')}</td>
          <td>
            <strong>${this.esc(r.variety_name || '-')}</strong>
            <br><small style="color:var(--gray-500)">${this.esc(r.brand || '')}${r.technology ? ' · ' + this.esc(r.technology) : ''}</small>
          </td>
          <td>${this.esc(r.producer_name || '-')}</td>
          <td>${this.esc(r.farm_name || '-')}</td>
          <td>${this.esc(r.city || '-')}/${this.esc(r.state || '-')}</td>
          <td>
            <strong style="color:var(--green);font-size:15px">
              ${parseFloat(r.productivity || 0).toLocaleString('pt-BR', {minimumFractionDigits:1, maximumFractionDigits:1})}
            </strong>
            <small>${this.esc(r.unit || '')}</small>
          </td>
          <td>${this.esc(r.season || '-')}</td>
          <td>
            ${r.termo_filename
              ? `<span class="badge badge-published" title="${this.esc(r.termo_nome_padronizado||r.termo_filename)}">
                   <i class="fas fa-check"></i> Sim
                 </span>`
              : `<span class="badge badge-draft"><i class="fas fa-times"></i> Não</span>`
            }
          </td>
          <td>${this.statusBadgeRecord(r.status)}</td>
          <td>
            ${typeof AccessControl !== 'undefined'
              ? AccessControl.renderAdminRecordActions(r.id)
              : `<div class="table-actions">
                   <button class="action-btn action-btn-blue" onclick="Admin.previewRecord('${r.id}')">
                     <i class="fas fa-image"></i> Card
                   </button>
                   <button class="action-btn action-btn-red" onclick="Admin.deleteRecord('${r.id}')">
                     <i class="fas fa-trash"></i>
                   </button>
                 </div>`
            }
          </td>
        </tr>
      `).join('')}
      </tbody>
    </table>`;
  },

  async previewRecord(id) {
    try {
      // Usa registro já carregado em memória para evitar re-fetch de payload grande
      let record = (this._allRecordsRaw || []).find(r => r.id === id);

      if (!record) {
        // Fallback: busca da lista completa
        try {
          const res = await API.getRecords();
          record = (res.data || []).find(r => r.id === id);
        } catch {}
      }

      if (!record) {
        // Último recurso: GET individual
        record = await API.getRecord(id);
      }

      if (!record) {
        App.Toast.show('Registro não encontrado.', 'error');
        return;
      }

      try {
        const vRes = await API.getVarieties();
        const v    = (vRes.data || []).find(x => x.id === record.variety_id);
        if (v) {
          record._color         = v.primary_color  || '#2E7D32';
          record._templateImage = v.template_image || null;
        }
      } catch {}

      CardGenerator.openPreview(record);
    } catch (err) {
      console.error('[Admin.previewRecord]', err);
      App.Toast.show('Erro ao carregar registro: ' + (err.message || 'tente novamente.'), 'error');
    }
  },

  async deleteRecord(id) {
    App.confirm('Excluir este registro permanentemente?', async () => {
      try {
        await API.deleteRecord(id);
        App.Toast.show('Registro excluído.', 'info');
        await this.loadAllRecords();
      } catch {
        App.Toast.show('Erro ao excluir.', 'error');
      }
    });
  },

  /* ═══════════════════════════════════════════════════
     CARDS GALLERY (admin)
  ═══════════════════════════════════════════════════ */
  async loadCardsGallery() {
    const grid = document.getElementById('cardsGalleryGrid');
    if (!grid) return;
    try {
      const res     = await API.getRecords();
      const records = (res.data || [])
        .filter(r => r.status === 'published')
        .sort((a, b) => (b.created_at || 0) - (a.created_at || 0));

      if (records.length === 0) {
        grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
          <i class="fas fa-images"></i>
          <p>Nenhum card publicado ainda</p>
        </div>`;
        return;
      }

      grid.innerHTML = records.map(r => `
        <div class="gallery-card">
          <div class="gallery-card-thumb" style="background:linear-gradient(135deg,#0d2b0d,${r._color || '#2E7D32'})">
            <div style="padding:10px;color:white;text-align:center;font-size:11px;font-weight:700;margin-top:20px">
              <div style="font-size:28px;font-weight:900">
                ${parseFloat(r.productivity||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}
              </div>
              <div>${this.esc(r.unit || 'sc/ha')}</div>
              <div style="margin-top:6px;opacity:.8">${this.esc(r.variety_name||'-')}</div>
            </div>
          </div>
          <div class="gallery-card-info">
            <div class="gallery-card-title">${this.esc(r.variety_name || '-')}</div>
            <div class="gallery-card-sub">${this.esc(r.producer_name||'-')} · ${this.esc(r.city||'-')}/${this.esc(r.state||'-')}</div>
          </div>
          <div class="gallery-card-actions">
            <button class="btn-primary btn-sm" style="flex:1" onclick="Admin.previewRecord('${r.id}')">
              <i class="fas fa-image"></i> Visualizar
            </button>
          </div>
        </div>
      `).join('');

    } catch (err) {
      console.error(err);
    }
  },

  /* ═══════════════════════════════════════════════════
     HELPERS
  ═══════════════════════════════════════════════════ */
  esc(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  },

  statusBadge(status) {
    const m = {
      pending:  '<span class="badge badge-pending"><i class="fas fa-clock"></i> Pendente</span>',
      approved: '<span class="badge badge-approved"><i class="fas fa-check"></i> Aprovado</span>',
      rejected: '<span class="badge badge-rejected"><i class="fas fa-times"></i> Rejeitado</span>'
    };
    return m[status] || `<span class="badge">${status}</span>`;
  },

  statusBadgeRecord(status) {
    const m = {
      draft:     '<span class="badge badge-draft"><i class="fas fa-edit"></i> Rascunho</span>',
      published: '<span class="badge badge-published"><i class="fas fa-check"></i> Publicado</span>'
    };
    return m[status] || `<span class="badge">${status}</span>`;
  },

  /* ═══════════════════════════════════════════════════
     TERMOS DE AUTORIZAÇÃO
  ═══════════════════════════════════════════════════ */
  _allTermos: [],

  async loadTermos() {
    const container = document.getElementById('termosContainer');
    if (!container) return;
    container.innerHTML = `<div style="padding:24px;text-align:center;color:#888">
      <span class="loading"></span> Carregando termos...
    </div>`;

    try {
      const res = await API.getRecords();
      this._allTermos = (res.data || [])
        .filter(r => r.termo_filename)
        .sort((a, b) => (b.created_at || 0) - (a.created_at || 0));

      const countEl = document.getElementById('termosCount');
      if (countEl) countEl.textContent = `${this._allTermos.length} termo(s) encontrado(s)`;

      this.renderTermos(this._allTermos);
    } catch (err) {
      console.error(err);
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-exclamation-circle"></i>
        <p>Erro ao carregar termos</p>
      </div>`;
    }
  },

  filterTermos(q) {
    if (!q) { this.renderTermos(this._allTermos); return; }
    const lower = q.toLowerCase();
    this.renderTermos(
      this._allTermos.filter(r =>
        [r.producer_name, r.variety_name, r.city, r.state, r.termo_filename, r.termo_nome_padronizado]
          .some(v => v && v.toLowerCase().includes(lower))
      )
    );
  },

  renderTermos(records) {
    const container = document.getElementById('termosContainer');
    if (!container) return;

    if (!records || records.length === 0) {
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-file-contract"></i>
        <p>Nenhum termo encontrado</p>
      </div>`;
      return;
    }

    container.innerHTML = `<table class="data-table">
      <thead><tr>
        <th>Produtor</th><th>Variedade</th><th>Cidade/UF</th>
        <th>Arquivo</th><th>OneDrive</th><th>Data</th><th>Ações</th>
      </tr></thead>
      <tbody>
      ${records.map(r => `
        <tr>
          <td><strong>${this.esc(r.producer_name || '-')}</strong></td>
          <td>${this.esc(r.variety_name || '-')}<br><small>${this.esc(r.brand||'')} · ${this.esc(r.season||'')}</small></td>
          <td>${this.esc(r.city||'-')}/${this.esc(r.state||'-')}</td>
          <td>
            <span title="${this.esc(r.termo_nome_padronizado || r.termo_filename)}" style="font-size:11px">
              <i class="fas fa-file-alt"></i>
              ${this.esc((r.termo_nome_padronizado || r.termo_filename || '').substring(0, 35))}
            </span>
          </td>
          <td>
            ${r.termo_onedrive_path
              ? `<a href="${this.esc(r.termo_onedrive_path)}" target="_blank" class="badge badge-published" style="text-decoration:none">
                   <i class="fab fa-microsoft"></i> Ver
                 </a>`
              : `<span class="badge badge-draft"><i class="fas fa-times"></i> Não enviado</span>`
            }
          </td>
          <td style="font-size:11px">
            ${r.created_at ? new Date(r.created_at).toLocaleDateString('pt-BR') : '-'}
          </td>
          <td>
            <div class="table-actions">
              ${r.termo_file
                ? `<button class="action-btn action-btn-blue" onclick="Admin.viewTermo('${r.id}')" title="Visualizar termo">
                     <i class="fas fa-eye"></i>
                   </button>`
                : ''
              }
              <button class="action-btn action-btn-blue" onclick="Admin.previewRecord('${r.id}')" title="Ver card">
                <i class="fas fa-image"></i>
              </button>
            </div>
          </td>
        </tr>
      `).join('')}
      </tbody>
    </table>`;
  },

  async viewTermo(recordId) {
    try {
      const record = await API.getRecord(recordId);
      if (!record.termo_file) {
        App.Toast.show('Arquivo do termo não encontrado neste registro.', 'warning');
        return;
      }

      // Detecta tipo do arquivo
      const dataUrl  = record.termo_file;
      const isPdf    = dataUrl.includes('application/pdf') || dataUrl.includes(';base64,JVB');

      if (isPdf) {
        // Abre PDF em nova aba
        const blob = this._dataUrlToBlob(dataUrl);
        const url  = URL.createObjectURL(blob);
        window.open(url, '_blank');
        setTimeout(() => URL.revokeObjectURL(url), 30000);
      } else {
        // Exibe imagem em modal simples
        const w = window.open('', '_blank', 'width=800,height=1000');
        w.document.write(`<!DOCTYPE html><html><head><title>Termo</title>
          <style>body{margin:0;background:#1a1a1a;display:flex;align-items:center;justify-content:center;min-height:100vh}
          img{max-width:100%;max-height:100vh;object-fit:contain}</style></head>
          <body><img src="${dataUrl}" alt="Termo de Autorização" /></body></html>`);
        w.document.close();
      }
    } catch (err) {
      App.Toast.show('Erro ao abrir termo.', 'error');
      console.error(err);
    }
  },

  _dataUrlToBlob(dataUrl) {
    const [header, data] = dataUrl.split(',');
    const mime = (header.match(/:(.*?);/) || ['','application/octet-stream'])[1];
    const binary = atob(data);
    const array  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) array[i] = binary.charCodeAt(i);
    return new Blob([array], { type: mime });
  },

  exportTermosCSV() {
    const records = this._allTermos;
    if (!records || records.length === 0) {
      App.Toast.show('Nenhum termo para exportar.', 'warning');
      return;
    }

    const headers = ['Produtor','Fazenda','Variedade','Marca','Tecnologia','Cultura','Safra',
      'Cidade','UF','Produtividade','Unidade','Arquivo Termo','Caminho OneDrive','Data Criação'];

    const rows = records.map(r => [
      r.producer_name || '',
      r.farm_name || '',
      r.variety_name || '',
      r.brand || '',
      r.technology || '',
      r.culture || '',
      r.season || '',
      r.city || '',
      r.state || '',
      r.productivity || '',
      r.unit || '',
      r.termo_nome_padronizado || r.termo_filename || '',
      r.termo_onedrive_path || '',
      r.created_at ? new Date(r.created_at).toLocaleDateString('pt-BR') : ''
    ]);

    this._downloadCSV('termos_autorizacao.csv', headers, rows);
  },

  /* ═══════════════════════════════════════════════════
     AUDITORIA
  ═══════════════════════════════════════════════════ */
  _allAuditLogs: [],

  async loadAuditLogs() {
    const container = document.getElementById('auditContainer');
    if (!container) return;
    container.innerHTML = `<div style="padding:24px;text-align:center;color:#888">
      <span class="loading"></span> Carregando logs...
    </div>`;

    try {
      const res = await fetch('tables/audit_logs?limit=500&sort=created_at');
      const data = await res.json();
      this._allAuditLogs = (data.data || []).sort((a, b) => (b.created_at || 0) - (a.created_at || 0));

      const countEl = document.getElementById('auditCount');
      if (countEl) countEl.textContent = `${this._allAuditLogs.length} evento(s)`;

      this.renderAuditLogs(this._allAuditLogs);
    } catch (err) {
      console.error(err);
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-history"></i>
        <p>Erro ao carregar logs de auditoria</p>
      </div>`;
    }
  },

  filterAuditLogs(q) {
    if (!q) { this.renderAuditLogs(this._allAuditLogs); return; }
    const lower = q.toLowerCase();
    this.renderAuditLogs(
      this._allAuditLogs.filter(r =>
        [r.user_name, r.action, r.producer_name, r.variety_name, r.culture, r.city]
          .some(v => v && v.toLowerCase().includes(lower))
      )
    );
  },

  renderAuditLogs(logs) {
    const container = document.getElementById('auditContainer');
    if (!container) return;

    if (!logs || logs.length === 0) {
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-history"></i>
        <p>Nenhum log encontrado</p>
      </div>`;
      return;
    }

    const actionLabels = {
      'record_created': '<span class="badge badge-published"><i class="fas fa-plus"></i> Criação</span>',
      'card_generated': '<span class="badge" style="background:#2563eb;color:#fff"><i class="fas fa-image"></i> Card Gerado</span>',
      'card_downloaded': '<span class="badge" style="background:#0891b2;color:#fff"><i class="fas fa-download"></i> Download</span>',
      'termo_uploaded': '<span class="badge" style="background:#7c3aed;color:#fff"><i class="fas fa-file-alt"></i> Termo Enviado</span>'
    };

    container.innerHTML = `<table class="data-table">
      <thead><tr>
        <th>Data/Hora</th><th>Usuário</th><th>Ação</th><th>Produtor</th>
        <th>Variedade</th><th>Cidade/UF</th><th>Card</th><th>Termo</th>
      </tr></thead>
      <tbody>
      ${logs.map(log => `
        <tr>
          <td style="font-size:11px;white-space:nowrap">
            ${log.created_at ? new Date(log.created_at).toLocaleString('pt-BR') : '-'}
          </td>
          <td>
            <strong>${this.esc(log.user_name || '-')}</strong>
          </td>
          <td>${actionLabels[log.action] || `<span class="badge">${this.esc(log.action||'-')}</span>`}</td>
          <td>${this.esc(log.producer_name || '-')}</td>
          <td>${this.esc(log.variety_name || '-')}<br><small>${this.esc(log.culture||'')} · ${this.esc(log.season||'')}</small></td>
          <td>${this.esc(log.city||'-')}/${this.esc(log.state||'-')}</td>
          <td style="font-size:11px">
            ${log.card_onedrive_path
              ? `<a href="${this.esc(log.card_onedrive_path)}" target="_blank" title="${this.esc(log.card_filename||'')}">
                   <i class="fab fa-microsoft"></i> OneDrive
                 </a>`
              : log.card_filename
                ? `<span title="${this.esc(log.card_filename)}"><i class="fas fa-file-image"></i> ${this.esc(log.card_filename.substring(0,20))}</span>`
                : '-'
            }
          </td>
          <td style="font-size:11px">
            ${log.termo_onedrive_path
              ? `<a href="${this.esc(log.termo_onedrive_path)}" target="_blank">
                   <i class="fab fa-microsoft"></i> OneDrive
                 </a>`
              : log.termo_filename
                ? `<i class="fas fa-file-alt"></i> ${this.esc(log.termo_filename.substring(0,20))}`
                : '-'
            }
          </td>
        </tr>
      `).join('')}
      </tbody>
    </table>`;
  },

  exportAuditCSV() {
    const logs = this._allAuditLogs;
    if (!logs || logs.length === 0) {
      App.Toast.show('Nenhum log para exportar.', 'warning');
      return;
    }

    const headers = ['Data/Hora','Usuário','Ação','Produtor','Variedade','Cultura','Safra',
      'Cidade','UF','Produtividade','Unidade','Card (arquivo)','Card (OneDrive)','Termo (arquivo)','Termo (OneDrive)'];

    const rows = logs.map(log => [
      log.created_at ? new Date(log.created_at).toLocaleString('pt-BR') : '',
      log.user_name || '',
      log.action || '',
      log.producer_name || '',
      log.variety_name || '',
      log.culture || '',
      log.season || '',
      log.city || '',
      log.state || '',
      log.productivity || '',
      log.unit || '',
      log.card_filename || '',
      log.card_onedrive_path || '',
      log.termo_filename || '',
      log.termo_onedrive_path || ''
    ]);

    this._downloadCSV('auditoria_agricard.csv', headers, rows);
  },

  exportAllRecordsCSV() {
    const records = this._allRecordsRaw;
    if (!records || records.length === 0) {
      App.Toast.show('Nenhum registro para exportar.', 'warning');
      return;
    }

    const headers = ['Usuário','Produtor','Fazenda','Variedade','Marca','Tecnologia','Cultura','Safra',
      'Data Plantio','Data Colheita','Produtividade','Unidade','Área','Cidade','UF',
      'Status','Termo','Card (OneDrive)','Data Criação'];

    const rows = records.map(r => [
      r.user_name || '',
      r.producer_name || '',
      r.farm_name || '',
      r.variety_name || '',
      r.brand || '',
      r.technology || '',
      r.culture || '',
      r.season || '',
      r.planting_date || '',
      r.harvest_date || '',
      r.productivity || '',
      r.unit || '',
      r.area || '',
      r.city || '',
      r.state || '',
      r.status || '',
      r.termo_nome_padronizado || r.termo_filename || '',
      r.card_onedrive_path || '',
      r.created_at ? new Date(r.created_at).toLocaleDateString('pt-BR') : ''
    ]);

    this._downloadCSV('registros_agricard.csv', headers, rows);
  },

  _downloadCSV(filename, headers, rows) {
    const BOM = '\uFEFF'; // UTF-8 BOM for Excel
    const escape = v => {
      const str = String(v || '');
      return str.includes(',') || str.includes('"') || str.includes('\n')
        ? `"${str.replace(/"/g, '""')}"`
        : str;
    };
    const csv = BOM +
      headers.map(escape).join(',') + '\r\n' +
      rows.map(row => row.map(escape).join(',')).join('\r\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href     = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
};
