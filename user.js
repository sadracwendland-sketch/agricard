/* =====================================================
   AgriCard Stine – User Module v3.0
   - Controle de acesso por perfil
   - Upload obrigatório de Termo de Autorização
   - Fluxo pós-geração: preview + redirect
   - Integração com OneDrive (via CardGenerator)
   ===================================================== */

'use strict';

const User = {
  currentEditId:   null,
  _allMyRecords:   [],
  _termoFile:      null,   // File objeto do termo selecionado
  _termoBase64:    null,   // base64 do termo

  async init() {
    await this.loadDashboard();
    await this.loadVarietiesSelect();
    this.setupFormEvents();
    // Aplica restrições de UI após inicialização
    if (typeof AccessControl !== 'undefined') AccessControl.applyUI();
  },

  // ===================================================
  // DASHBOARD
  // ===================================================
  async loadDashboard() {
    const userId = Auth.currentUser?.id;
    if (!userId) return;

    const greet = document.getElementById('userWelcome');
    if (greet) {
      const hour = new Date().getHours();
      const g    = hour < 12 ? 'Bom dia' : hour < 18 ? 'Boa tarde' : 'Boa noite';
      greet.textContent = `${g}, ${Auth.currentUser.name?.split(' ')[0] || 'Usuário'}! `;
    }

    try {
      const res     = await API.getRecords();
      const all     = res.data || [];
      const records = all.filter(r => r.user_id === userId);
      const sorted  = records.sort((a, b) => (b.created_at || 0) - (a.created_at || 0));

      // Popula _allMyRecords para uso no previewRecord sem re-fetch
      this._allMyRecords = sorted;

      document.getElementById('userStatRecords').textContent   = records.length;
      document.getElementById('userStatCards').textContent     = records.filter(r => r.status === 'published').length;
      document.getElementById('userStatDownloads').textContent = records.filter(r => r.status === 'published').length;

      const banner = document.getElementById('quickStartBanner');
      if (banner) banner.style.display = records.length === 0 ? 'flex' : 'none';

      const container = document.getElementById('userRecentRecords');
      const recent    = sorted.slice(0, 6);

      if (recent.length === 0) {
        container.innerHTML = `<div class="empty-state">
          <i class="fas fa-seedling"></i>
          <p>Nenhum registro ainda.
            <a href="#" onclick="App.navigate('user-new-record')">Crie seu primeiro card!</a>
          </p>
        </div>`;
        return;
      }

      const isAdmin = typeof AccessControl !== 'undefined' && AccessControl.isAdmin();

      container.innerHTML = `<table class="data-table">
        <thead><tr>
          <th>Variedade</th><th>Produtor</th><th>Cidade/UF</th>
          <th>Produtividade</th><th>Status</th><th>Ações</th>
        </tr></thead>
        <tbody>
        ${recent.map(r => `
          <tr>
            <td>
              <strong>${this.esc(r.variety_name || '-')}</strong>
              <br><small style="color:var(--gray-500)">${this.esc(r.brand||'')} · Safra ${this.esc(r.season||'-')}</small>
            </td>
            <td>${this.esc(r.producer_name || '-')}</td>
            <td>${this.esc(r.city||'-')}/${this.esc(r.state||'-')}</td>
            <td>
              <strong style="color:var(--green);font-size:15px">
                ${parseFloat(r.productivity||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}
              </strong>
              <small>${this.esc(r.unit||'')}</small>
            </td>
            <td>${this.statusBadge(r.status)}</td>
            <td>
              ${typeof AccessControl !== 'undefined'
                ? AccessControl.renderUserRecordActions(r.id)
                : `<button class="action-btn action-btn-blue" onclick="User.previewRecord('${r.id}')">
                    <i class="fas fa-image"></i> Card</button>`
              }
            </td>
          </tr>
        `).join('')}
        </tbody>
      </table>`;

    } catch (err) {
      console.error('User.loadDashboard error:', err);
    }
  },

  // ===================================================
  // VARIETIES SELECT
  // ===================================================
  async loadVarietiesSelect() {
    try {
      const res      = await API.getVarieties();
      const varieties = (res.data || []).sort((a, b) =>
        `${a.brand} ${a.name}`.localeCompare(`${b.brand} ${b.name}`)
      );
      const sel = document.getElementById('recVariety');
      if (!sel) return;
      sel.innerHTML = '<option value="">Selecione a variedade...</option>';

      const brands = [...new Set(varieties.map(v => v.brand))];
      brands.forEach(brand => {
        const group = document.createElement('optgroup');
        group.label = brand;
        varieties.filter(v => v.brand === brand).forEach(v => {
          const opt = document.createElement('option');
          opt.value = v.id;
          opt.textContent = `${v.name} (${v.culture})`;
          opt.dataset.technology    = v.technology    || '';
          opt.dataset.culture       = v.culture       || '';
          opt.dataset.brand         = v.brand         || '';
          opt.dataset.name          = v.name          || '';
          opt.dataset.color         = v.primary_color || '#2E7D32';
          opt.dataset.maturityGroup = v.maturity_group || '';
          group.appendChild(opt);
        });
        sel.appendChild(group);
      });
    } catch (err) {
      console.error(err);
    }
  },

  // ===================================================
  // FORM EVENTS
  // ===================================================
  setupFormEvents() {
    // Variety change
    document.getElementById('recVariety')?.addEventListener('change', e => {
      const opt = e.target.options[e.target.selectedIndex];
      if (opt.value) {
        document.getElementById('recTechnology').value = opt.dataset.technology || '';
        document.getElementById('recCulture').value    = opt.dataset.culture    || '';
      } else {
        document.getElementById('recTechnology').value = '';
        document.getElementById('recCulture').value    = '';
      }
    });

    // State uppercase
    document.getElementById('recState')?.addEventListener('input', e => {
      e.target.value = e.target.value.toUpperCase().slice(0, 2);
    });

    // Productivity format
    document.getElementById('recProductivity')?.addEventListener('input', e => {
      e.target.value = e.target.value.replace(/[^\d,\.]/g, '');
    });

    // Termo upload
    const termoInput = document.getElementById('recTermo');
    if (termoInput) {
      termoInput.addEventListener('change', e => this._handleTermoFile(e.target.files[0]));
    }

    // Drag & drop na zona do termo
    const termoZone = document.getElementById('termoDropzone');
    if (termoZone) {
      termoZone.addEventListener('dragover', e => { e.preventDefault(); termoZone.classList.add('dragover'); });
      termoZone.addEventListener('dragleave', () => termoZone.classList.remove('dragover'));
      termoZone.addEventListener('drop', e => {
        e.preventDefault();
        termoZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) { termoInput.files = e.dataTransfer.files; this._handleTermoFile(file); }
      });
      termoZone.addEventListener('click', () => termoInput?.click());
    }

    // Buttons
    document.getElementById('btnSaveDraft')?.addEventListener('click',      () => this.saveRecord('draft'));
    document.getElementById('btnSaveAndPreview')?.addEventListener('click', () => this.saveRecord('published', true));
  },

  // ===================================================
  // TERMO DE AUTORIZAÇÃO
  // ===================================================
  async _handleTermoFile(file) {
    if (!file) return;

    const validation = typeof AccessControl !== 'undefined'
      ? AccessControl.validateTermo(file)
      : { valid: true };

    if (!validation.valid) {
      App.Toast.show(validation.error, 'error');
      document.getElementById('recTermo').value = '';
      this._resetTermo();
      return;
    }

    this._termoFile = file;

    // Converte para base64
    try {
      this._termoBase64 = await this._readFileBase64(file);
    } catch {
      App.Toast.show('Erro ao ler arquivo do termo.', 'error');
      return;
    }

    // Atualiza UI
    const statusEl = document.getElementById('termoStatus');
    const nameEl   = document.getElementById('termoFileName');
    const sizeEl   = document.getElementById('termoFileSize');

    if (statusEl) statusEl.classList.remove('hidden');
    if (nameEl)   nameEl.textContent = file.name;
    if (sizeEl)   sizeEl.textContent = `${(file.size / 1024).toFixed(1)} KB`;

    const zone = document.getElementById('termoDropzone');
    if (zone) zone.classList.add('has-file');
  },

  _resetTermo() {
    this._termoFile   = null;
    this._termoBase64 = null;
    const statusEl = document.getElementById('termoStatus');
    if (statusEl) statusEl.classList.add('hidden');
    const zone = document.getElementById('termoDropzone');
    if (zone) zone.classList.remove('has-file');
  },

  _readFileBase64(file) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload  = e => res(e.target.result);
      r.onerror = () => rej(new Error('Falha ao ler arquivo'));
      r.readAsDataURL(file);
    });
  },

  // ===================================================
  // SAVE RECORD
  // ===================================================
  async saveRecord(status = 'draft', openPreview = false) {
    // Bloqueia edição para usuário comum (segurança)
    if (this.currentEditId && typeof AccessControl !== 'undefined' && !AccessControl.can('can_edit_card')) {
      App.Toast.show('Você não tem permissão para editar registros.', 'error');
      return;
    }

    const sel = document.getElementById('recVariety');
    const opt = sel?.options[sel.selectedIndex];

    const fields = {
      user_id:          Auth.currentUser?.id,
      user_name:        Auth.currentUser?.name || '',
      variety_id:       sel?.value || '',
      variety_name:     opt?.dataset.name  || '',
      brand:            opt?.dataset.brand || '',
      technology:       document.getElementById('recTechnology')?.value.trim() || '',
      culture:          document.getElementById('recCulture')?.value.trim() || '',
      season:           document.getElementById('recSeason')?.value.trim() || '',
      planting_date:    document.getElementById('recPlantingDate')?.value.trim() || '',
      harvest_date:     document.getElementById('recHarvestDate')?.value.trim() || '',
      productivity:     (document.getElementById('recProductivity')?.value.trim() || '').replace('.', ','),
      unit:             document.getElementById('recUnit')?.value || 'sc/ha',
      area:             document.getElementById('recArea')?.value.trim() || '',
      city:             document.getElementById('recCity')?.value.trim() || '',
      state:            document.getElementById('recState')?.value.trim().toUpperCase() || '',
      producer_name:    document.getElementById('recProducer')?.value.trim() || '',
      farm_name:        document.getElementById('recFarm')?.value.trim() || '',
      plant_population: document.getElementById('recPopulation')?.value.trim() || '',
      notes:            document.getElementById('recNotes')?.value.trim() || '',
      lgpd_accepted:    true,
      status
    };

    // Validação obrigatória
    const required = [
      ['variety_id',    'Selecione uma variedade'],
      ['planting_date', 'Informe a data de plantio'],
      ['harvest_date',  'Informe a data de colheita'],
      ['productivity',  'Informe a produtividade (ex: 110,7)'],
      ['area',          'Informe a área colhida'],
      ['city',          'Informe a cidade'],
      ['state',         'Informe o estado (UF)'],
      ['producer_name', 'Informe o nome do produtor'],
      ['farm_name',     'Informe o nome da fazenda']
    ];

    for (const [field, msg] of required) {
      if (!fields[field]) {
        App.Toast.show(msg + '.', 'error');
        return;
      }
    }

    // Validação do termo de autorização (obrigatório apenas na criação)
    if (!this.currentEditId && !this._termoBase64) {
      App.Toast.show('O Termo de Autorização é obrigatório. Faça o upload do arquivo.', 'error');
      document.getElementById('termoDropzone')?.scrollIntoView({ behavior: 'smooth' });
      return;
    }

    // Inclui dados do termo se existirem
    if (this._termoBase64) {
      const nomeTermoPadronizado = typeof OneDrive !== 'undefined'
        ? OneDrive.buildFilename(fields, this._termoFile?.name?.split('.').pop() || 'pdf')
        : this._buildFilename(fields, this._termoFile?.name?.split('.').pop() || 'pdf');

      fields.termo_file              = this._termoBase64;
      fields.termo_filename          = this._termoFile?.name || '';
      fields.termo_nome_padronizado  = nomeTermoPadronizado;
    }

    // Show loading
    const previewBtn = document.getElementById('btnSaveAndPreview');
    const draftBtn   = document.getElementById('btnSaveDraft');
    if (previewBtn) previewBtn.innerHTML = '<span class="loading"></span> Salvando...';
    if (previewBtn) previewBtn.disabled = true;
    if (draftBtn)   draftBtn.disabled   = true;

    try {
      let record;
      if (this.currentEditId) {
        record = await API.updateRecord(this.currentEditId, fields);
        App.Toast.show('Registro atualizado com sucesso!', 'success');
        this.currentEditId = null;
        const titleEl = document.getElementById('formTitle');
        if (titleEl) titleEl.innerHTML = '<i class="fas fa-plus-circle"></i> Novo Registro de Produtividade';
      } else {
        record = await API.createRecord(fields);

        // Log de auditoria
        if (typeof OneDrive !== 'undefined') {
          OneDrive.log('record_created', record, {
            termo_filename: fields.termo_nome_padronizado
          });
        }

        App.Toast.show('Registro salvo com sucesso!', 'success');
      }

      this.clearForm();

      if (openPreview && record) {
        // Aguarda um momento e abre o card
        setTimeout(() => {
          const colorOpt = document.getElementById('recVariety');
          if (colorOpt?.options[colorOpt.selectedIndex]) {
            record._color = colorOpt.options[colorOpt.selectedIndex]?.dataset?.color || '#2E7D32';
          }
          CardGenerator.openPreview(record);
        }, 300);
      }

      await this.loadDashboard();

    } catch (err) {
      App.Toast.show('Erro ao salvar. Verifique os dados.', 'error');
      console.error(err);
    } finally {
      if (previewBtn) previewBtn.innerHTML = '<i class="fas fa-eye"></i> Salvar e Gerar Card';
      if (previewBtn) previewBtn.disabled = false;
      if (draftBtn)   draftBtn.disabled   = false;
    }
  },

  clearForm() {
    const textIds = [
      'recTechnology','recCulture','recSeason',
      'recPlantingDate','recHarvestDate','recProductivity',
      'recArea','recCity','recState',
      'recProducer','recFarm','recPopulation','recNotes'
    ];
    textIds.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    const sel = document.getElementById('recVariety');
    if (sel) sel.selectedIndex = 0;
    const unitSel = document.getElementById('recUnit');
    if (unitSel) unitSel.value = 'sc/ha';

    // Reset termo
    this._resetTermo();
    const termoInput = document.getElementById('recTermo');
    if (termoInput) termoInput.value = '';

    this.currentEditId = null;
  },

  cancelEdit() {
    this.currentEditId = null;
    this.clearForm();
    const titleEl = document.getElementById('formTitle');
    if (titleEl) titleEl.innerHTML = '<i class="fas fa-plus-circle"></i> Novo Registro de Produtividade';
    App.navigate('user-dashboard');
  },

  // ===================================================
  // MY RECORDS
  // ===================================================
  async loadMyRecords() {
    const userId = Auth.currentUser?.id;
    try {
      const res = await API.getRecords();
      this._allMyRecords = (res.data || [])
        .filter(r => r.user_id === userId)
        .sort((a, b) => (b.created_at || 0) - (a.created_at || 0));
      this.renderMyRecords(this._allMyRecords);
    } catch (err) {
      console.error(err);
    }
  },

  filterMyRecords(q) {
    if (!q) { this.renderMyRecords(this._allMyRecords); return; }
    const lower = q.toLowerCase();
    this.renderMyRecords(
      this._allMyRecords.filter(r =>
        [r.variety_name, r.producer_name, r.farm_name, r.city, r.state, r.season]
          .some(v => v && v.toLowerCase().includes(lower))
      )
    );
  },

  renderMyRecords(records) {
    const container = document.getElementById('myRecordsContainer');
    if (!container) return;

    if (!records || records.length === 0) {
      container.innerHTML = `<div class="empty-state">
        <i class="fas fa-inbox"></i>
        <p>Nenhum registro encontrado.</p>
      </div>`;
      return;
    }

    container.innerHTML = `<table class="data-table">
      <thead><tr>
        <th>Variedade</th><th>Produtor</th><th>Fazenda</th>
        <th>Cidade/UF</th><th>Produtividade</th><th>Colheita</th><th>Termo</th><th>Status</th><th>Ações</th>
      </tr></thead>
      <tbody>
      ${records.map(r => `
        <tr>
          <td>
            <strong>${this.esc(r.variety_name || '-')}</strong>
            <br><small style="color:var(--gray-500)">${this.esc(r.brand||'')} · Safra ${this.esc(r.season||'-')}</small>
          </td>
          <td>${this.esc(r.producer_name || '-')}</td>
          <td>${this.esc(r.farm_name || '-')}</td>
          <td>${this.esc(r.city||'-')}/${this.esc(r.state||'-')}</td>
          <td>
            <strong style="color:var(--green);font-size:15px">
              ${parseFloat(r.productivity||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}
            </strong>
            <small>${this.esc(r.unit||'')}</small>
          </td>
          <td>${this.esc(r.harvest_date||'-')}</td>
          <td>
            ${r.termo_filename
              ? `<span class="badge badge-published" title="${this.esc(r.termo_nome_padronizado||r.termo_filename)}">
                   <i class="fas fa-check"></i> Enviado
                 </span>`
              : `<span class="badge badge-draft"><i class="fas fa-times"></i> Pendente</span>`
            }
          </td>
          <td>${this.statusBadge(r.status)}</td>
          <td>
            ${typeof AccessControl !== 'undefined'
              ? AccessControl.renderUserRecordActions(r.id)
              : `<button class="action-btn action-btn-blue" onclick="User.previewRecord('${r.id}')">
                   <i class="fas fa-image"></i> Card</button>`
            }
          </td>
        </tr>
      `).join('')}
      </tbody>
    </table>`;
  },

  // ===================================================
  // MY CARDS GALLERY
  // ===================================================
  async loadMyCards() {
    const userId = Auth.currentUser?.id;
    const grid   = document.getElementById('myCardsGrid');
    if (!grid) return;

    try {
      const res = await API.getRecords();
      const pub = (res.data || [])
        .filter(r => r.user_id === userId && r.status === 'published')
        .sort((a, b) => (b.created_at || 0) - (a.created_at || 0));

      if (pub.length === 0) {
        grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
          <i class="fas fa-images"></i>
          <p>Nenhum card publicado ainda. Salve um registro com "Salvar e Gerar Card".</p>
        </div>`;
        return;
      }

      grid.innerHTML = pub.map(r => `
        <div class="gallery-card" onclick="User.previewRecord('${r.id}')">
          <div class="gallery-card-thumb" style="
            background:linear-gradient(135deg, #0a1a0a, ${r._color||'#2E7D32'});
            display:flex; align-items:center; justify-content:center; flex-direction:column;
            color:white; text-align:center; gap:4px; padding:16px;
          ">
            <div style="font-size:11px;font-weight:700;letter-spacing:.1em;opacity:.7;text-transform:uppercase">
              ${this.esc(r.brand||'STINE')}
            </div>
            <div style="font-size:22px;font-weight:900;letter-spacing:-.5px">
              ${this.esc(r.variety_name||'-')}
            </div>
            <div style="font-size:36px;font-weight:900;line-height:1">
              ${parseFloat(r.productivity||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}
            </div>
            <div style="font-size:12px;opacity:.8">${this.esc(r.unit||'sc/ha')}</div>
          </div>
          <div class="gallery-card-info">
            <div class="gallery-card-title">${this.esc(r.variety_name||'-')}</div>
            <div class="gallery-card-sub">${this.esc(r.city||'-')}/${this.esc(r.state||'-')} · ${this.esc(r.season||'-')}</div>
            ${r.card_filename
              ? `<div style="font-size:10px;color:var(--green);margin-top:4px">
                   <i class="fas fa-cloud-upload-alt"></i> ${this.esc(r.card_filename)}
                 </div>`
              : ''
            }
          </div>
          <div class="gallery-card-actions">
            <button class="btn-primary btn-sm" style="flex:1"
              onclick="event.stopPropagation();User.previewRecord('${r.id}')">
              <i class="fas fa-download"></i> Baixar Card
            </button>
          </div>
        </div>
      `).join('');

    } catch (err) {
      console.error(err);
    }
  },

  // ===================================================
  // ACTIONS
  // ===================================================
  async previewRecord(id) {
    try {
      // Tenta usar o registro já carregado em memória (evita re-fetch de base64 grande)
      let record = this._allMyRecords.find(r => r.id === id);

      if (!record) {
        // Fallback: busca da lista completa (sem termo_file para evitar payload enorme)
        try {
          const res = await API.getRecords();
          record = (res.data || []).find(r => r.id === id);
        } catch {}
      }

      if (!record) {
        // Último recurso: busca individual
        record = await API.getRecord(id);
      }

      if (!record) {
        App.Toast.show('Registro não encontrado.', 'error');
        return;
      }

      // Busca dados da variedade
      try {
        const vRes = await API.getVarieties();
        const v    = (vRes.data || []).find(x => x.id === record.variety_id);
        if (v) {
          record._color         = v.primary_color  || '#2E7D32';
          record._templateImage = v.template_image || null;
        }
      } catch (vErr) {
        console.warn('[previewRecord] Falha ao buscar variedade:', vErr);
      }

      CardGenerator.openPreview(record);
    } catch (err) {
      console.error('[previewRecord] Erro:', err);
      App.Toast.show('Erro ao carregar registro: ' + (err.message || 'tente novamente.'), 'error');
    }
  },

  async editRecord(id) {
    // Bloqueia edição para usuário comum
    if (typeof AccessControl !== 'undefined' && !AccessControl.can('can_edit_card')) {
      App.Toast.show('Você não tem permissão para editar registros.', 'error');
      return;
    }

    try {
      const record = await API.getRecord(id);
      this.currentEditId = id;
      App.navigate('user-new-record');
      await this.loadVarietiesSelect();

      const titleEl = document.getElementById('formTitle');
      if (titleEl) titleEl.innerHTML = '<i class="fas fa-edit"></i> Editando Registro';

      const fields = {
        recVariety:      record.variety_id,
        recTechnology:   record.technology,
        recCulture:      record.culture,
        recSeason:       record.season,
        recPlantingDate: record.planting_date,
        recHarvestDate:  record.harvest_date,
        recProductivity: record.productivity,
        recUnit:         record.unit     || 'sc/ha',
        recArea:         record.area,
        recCity:         record.city,
        recState:        record.state,
        recProducer:     record.producer_name,
        recFarm:         record.farm_name,
        recPopulation:   record.plant_population,
        recNotes:        record.notes
      };

      Object.entries(fields).forEach(([id, val]) => {
        const el = document.getElementById(id);
        if (el) el.value = val || '';
      });

      App.Toast.show('Modo de edição ativado.', 'info');
    } catch {
      App.Toast.show('Erro ao carregar registro para edição.', 'error');
    }
  },

  async deleteRecord(id) {
    // Bloqueia exclusão para usuário comum
    if (typeof AccessControl !== 'undefined' && !AccessControl.can('can_delete_card')) {
      App.Toast.show('Você não tem permissão para excluir registros.', 'error');
      return;
    }

    App.confirm('Excluir este registro permanentemente?', async () => {
      try {
        await API.deleteRecord(id);
        App.Toast.show('Registro excluído.', 'info');
        await this.loadDashboard();
        await this.loadMyRecords();
      } catch {
        App.Toast.show('Erro ao excluir.', 'error');
      }
    });
  },

  // ===================================================
  // SUCCESS SCREEN — mostrada após geração do card
  // ===================================================
  showSuccessScreen(record, canvasDataUrl) {
    // Navega para a tela de cards
    App.navigate('user-cards');
    this.loadMyCards();

    // Mostra mensagem de sucesso com preview
    const grid = document.getElementById('myCardsGrid');
    if (grid && canvasDataUrl) {
      const successHtml = `
        <div class="card-success-banner" style="
          grid-column:1/-1;
          background:linear-gradient(135deg,#1a4a1a,#2e7d32);
          border-radius:16px; padding:24px; color:white;
          display:flex; align-items:center; gap:20px; margin-bottom:16px;
          box-shadow:0 8px 32px rgba(0,128,0,.3);
        ">
          <img src="${canvasDataUrl}" alt="Card gerado"
            style="width:100px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,.3);flex-shrink:0" />
          <div style="flex:1">
            <h3 style="margin:0 0 8px;font-size:18px">
              <i class="fas fa-check-circle"></i> Card gerado com sucesso!
            </h3>
            <p style="margin:0;font-size:13px;opacity:.9">
              ${this.esc(record.variety_name||'')} · ${this.esc(record.city||'')}/${this.esc(record.state||'')}
              · Safra ${this.esc(record.season||'-')}
            </p>
            ${record.card_onedrive_path
              ? `<p style="margin:4px 0 0;font-size:11px;opacity:.7">
                   <i class="fas fa-cloud"></i> Enviado ao OneDrive: ${this.esc(record.card_onedrive_path)}
                 </p>`
              : ''
            }
          </div>
          <a href="${canvasDataUrl}" download="${this.esc(record.card_filename || 'card.png')}"
            class="btn-primary" style="flex-shrink:0;white-space:nowrap">
            <i class="fas fa-download"></i> Baixar Card
          </a>
        </div>`;

      // Insere no topo da grid
      grid.insertAdjacentHTML('afterbegin', successHtml);
      grid.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    App.Toast.show('✅ Card gerado com sucesso!', 'success');
  },

  // ===================================================
  // HELPERS
  // ===================================================
  esc(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  },

  statusBadge(status) {
    const map = {
      draft:        '<span class="badge badge-draft"><i class="fas fa-edit"></i> Rascunho</span>',
      published:    '<span class="badge badge-published"><i class="fas fa-check"></i> Publicado</span>',
      term_pending: '<span class="badge" style="background:#f59e0b;color:#fff"><i class="fas fa-clock"></i> Aguardando Termo</span>'
    };
    return map[status] || `<span class="badge">${status}</span>`;
  },

  _buildFilename(record, ext) {
    const s = str => (str||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/[^a-zA-Z0-9\s]/g,'').trim().replace(/\s+/g,'_');
    const date = new Date().toISOString().slice(0,10);
    return [s(record.producer_name), s(record.variety_name), s(record.city), date].join('_') + '.' + ext;
  }
};
