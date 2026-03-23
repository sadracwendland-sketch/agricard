/* =====================================================
   AgriCard Stine - Templates Manager Module
   Gerencia os layouts de card disponíveis
   ===================================================== */

const TemplatesManager = {
  _templates: [],
  _editingId: null,

  // ===================================================
  // INIT
  // ===================================================
  async init() {
    await this.loadTemplates();
    this.setupEvents();
  },

  // ===================================================
  // SEED DEFAULT TEMPLATES (chamado na primeira vez)
  // ===================================================
  async seedDefaults() {
    const defaults = [
      {
        name: 'Stine Clássico Verde',
        description: 'Layout padrão com fundo de lavoura e painel branco central. Identidade visual STINE.',
        layout_type: 'classic',
        header_color: '#2E7D32',
        header_text: 'RESULTADOS DE PRODUTIVIDADE',
        slogan: 'NÃO É SORTE! É STINE',
        footer_logo_text: 'STINE',
        bg_image_url: 'https://images.unsplash.com/photo-1625246333195-78d9c38ad449?w=600&q=80',
        show_ranking_badge: false,
        badge_label: '',
        active: true,
        sort_order: 1
      },
      {
        name: 'Stine Dark Premium',
        description: 'Layout escuro com gradiente azul-verde. Visual premium para redes sociais.',
        layout_type: 'dark',
        header_color: '#1B5E20',
        header_text: 'RECORDES DE PRODUTIVIDADE',
        slogan: 'NÃO É SORTE! É STINE',
        footer_logo_text: 'STINE',
        bg_image_url: 'https://images.unsplash.com/photo-1574943320219-553eb213f72d?w=600&q=80',
        show_ranking_badge: false,
        badge_label: '',
        active: false,
        sort_order: 2
      },
      {
        name: 'Stine Campo Aberto',
        description: 'Fundo panorâmico de campo com overlay verde suave. Ideal para destaque de variedades.',
        layout_type: 'panoramic',
        header_color: '#388E3C',
        header_text: 'RESULTADO DE CAMPO',
        slogan: 'NÃO É SORTE! É STINE',
        footer_logo_text: 'STINE',
        bg_image_url: 'https://images.unsplash.com/photo-1500382017468-9049fed747ef?w=600&q=80',
        show_ranking_badge: false,
        badge_label: '',
        active: false,
        sort_order: 3
      },
      {
        name: 'Stine Ranking #1',
        description: 'Layout com medalha/ranking destacado. Perfeito para celebrar primeiro lugar em produtividade.',
        layout_type: 'ranking',
        header_color: '#F57F17',
        header_text: 'CAMPEÃO DE PRODUTIVIDADE',
        slogan: 'NÃO É SORTE! É STINE',
        footer_logo_text: 'STINE',
        bg_image_url: 'https://images.unsplash.com/photo-1625246333195-78d9c38ad449?w=600&q=80',
        show_ranking_badge: true,
        badge_label: '🏆 MELHOR DO MUNICÍPIO',
        active: false,
        sort_order: 4
      }
    ];

    for (const tpl of defaults) {
      try {
        await API.createTemplate(tpl);
      } catch (e) { /* ignora se já existir */ }
    }
  },

  // ===================================================
  // LOAD TEMPLATES
  // ===================================================
  async loadTemplates() {
    try {
      const res = await API.getTemplates();
      this._templates = (res.data || []).sort((a,b) => (a.sort_order||99) - (b.sort_order||99));

      // Se não há templates, cria os padrões
      if (this._templates.length === 0) {
        await this.seedDefaults();
        const res2 = await API.getTemplates();
        this._templates = (res2.data || []).sort((a,b) => (a.sort_order||99) - (b.sort_order||99));
      }

      this.renderTemplates();
      return this._templates;
    } catch (err) {
      console.error('TemplatesManager.loadTemplates error:', err);
      return [];
    }
  },

  // ===================================================
  // RENDER TEMPLATES GRID
  // ===================================================
  renderTemplates() {
    const container = document.getElementById('templatesGrid');
    if (!container) return;

    if (this._templates.length === 0) {
      container.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
        <i class="fas fa-palette"></i>
        <p>Nenhum template cadastrado. Clique em "Novo Template" para começar.</p>
      </div>`;
      return;
    }

    container.innerHTML = this._templates.map(t => `
      <div class="template-card ${t.active ? 'template-active' : ''}">
        <!-- Miniatura do layout -->
        <div class="template-thumb" style="background:linear-gradient(160deg, ${t.header_color||'#2E7D32'} 0%, #0d2b0d 100%)">
          <div class="tpl-thumb-inner">
            <div class="tpl-thumb-header" style="background:${t.header_color||'#2E7D32'}">
              ${this.esc(t.header_text || 'RESULTADOS DE PRODUTIVIDADE')}
            </div>
            <div class="tpl-thumb-body">
              ${t.show_ranking_badge ? '<div class="tpl-thumb-medal">🏆</div>' : ''}
              <div class="tpl-thumb-variety">VARIEDADE XX</div>
              <div class="tpl-thumb-prod" style="color:${t.header_color||'#2E7D32'}">110,5</div>
              <div class="tpl-thumb-unit">sc/ha</div>
              <div class="tpl-thumb-location">📍 Cidade/UF</div>
              ${t.badge_label ? `<div class="tpl-thumb-badge" style="background:${t.header_color||'#2E7D32'}">${this.esc(t.badge_label)}</div>` : ''}
            </div>
            <div class="tpl-thumb-footer" style="background:#111">
              <span style="color:#aaa;font-size:7px">${this.esc(t.slogan||'NÃO É SORTE! É STINE')}</span>
            </div>
          </div>
          ${t.bg_image_url ? `<img src="${this.esc(t.bg_image_url)}" alt="bg" class="tpl-thumb-bg" crossorigin="anonymous" onerror="this.style.display='none'"/>` : ''}
        </div>

        <!-- Info -->
        <div class="template-info">
          <div class="template-name-row">
            <span class="template-name">${this.esc(t.name)}</span>
            ${t.active ? '<span class="template-active-badge"><i class="fas fa-check-circle"></i> Ativo</span>' : ''}
          </div>
          <p class="template-desc">${this.esc(t.description || '')}</p>
          <div class="template-meta">
            <span class="tpl-meta-item" style="background:${t.header_color||'#2E7D32'}20;color:${t.header_color||'#2E7D32'}">
              <i class="fas fa-palette"></i> ${this.getLayoutName(t.layout_type)}
            </span>
            ${t.show_ranking_badge ? '<span class="tpl-meta-item" style="background:#fff3e0;color:#e65100"><i class="fas fa-medal"></i> Ranking</span>' : ''}
          </div>
        </div>

        <!-- Actions -->
        <div class="template-actions">
          ${!t.active ? `
            <button class="action-btn action-btn-green" onclick="TemplatesManager.setActive('${t.id}')">
              <i class="fas fa-check"></i> Ativar
            </button>
          ` : '<button class="action-btn" disabled style="opacity:.5;cursor:not-allowed"><i class="fas fa-check-circle"></i> Ativo</button>'}
          <button class="action-btn action-btn-blue" onclick="TemplatesManager.openModal('${t.id}')">
            <i class="fas fa-edit"></i> Editar
          </button>
          <button class="action-btn action-btn-red" onclick="TemplatesManager.deleteTemplate('${t.id}')">
            <i class="fas fa-trash"></i>
          </button>
        </div>
      </div>
    `).join('');
  },

  getLayoutName(type) {
    const map = { classic: 'Clássico', dark: 'Dark', panoramic: 'Panorâmico', ranking: 'Ranking', minimal: 'Minimalista', bold: 'Negrito' };
    return map[type] || (type || 'Padrão');
  },

  // ===================================================
  // SET ACTIVE TEMPLATE
  // ===================================================
  async setActive(id) {
    try {
      // Desativa todos
      for (const t of this._templates) {
        if (t.active && t.id !== id) {
          await API.updateTemplate(t.id, { active: false });
        }
      }
      // Ativa o escolhido
      await API.updateTemplate(id, { active: true });
      App.Toast.show('Template ativado com sucesso!', 'success');
      await this.loadTemplates();
      // Também recarrega os templates do CardGenerator
      CardGenerator._templates = [];
    } catch (err) {
      App.Toast.show('Erro ao ativar template.', 'error');
    }
  },

  // ===================================================
  // OPEN TEMPLATE MODAL
  // ===================================================
  async openModal(id = null) {
    this._editingId = id;
    const modal = document.getElementById('templateModal');
    if (!modal) return;

    // Limpa formulário
    this.clearForm();

    if (id) {
      const t = this._templates.find(x => x.id === id);
      if (t) {
        document.getElementById('tplName').value           = t.name || '';
        document.getElementById('tplDescription').value    = t.description || '';
        document.getElementById('tplLayoutType').value     = t.layout_type || 'classic';
        document.getElementById('tplHeaderColor').value    = t.header_color || '#2E7D32';
        document.getElementById('tplHeaderColorHex').textContent = t.header_color || '#2E7D32';
        document.getElementById('tplHeaderText').value     = t.header_text || '';
        document.getElementById('tplSlogan').value         = t.slogan || '';
        document.getElementById('tplFooterLogoText').value = t.footer_logo_text || '';
        document.getElementById('tplBgImageUrl').value     = t.bg_image_url || '';
        document.getElementById('tplBadgeLabel').value     = t.badge_label || '';
        document.getElementById('tplShowRanking').checked  = !!t.show_ranking_badge;
        document.getElementById('tplSortOrder').value      = t.sort_order || 1;
        document.getElementById('tplActive').checked       = !!t.active;
        document.getElementById('templateModalTitle').innerHTML = '<i class="fas fa-edit"></i> Editar Template';
      }
    } else {
      document.getElementById('templateModalTitle').innerHTML = '<i class="fas fa-plus"></i> Novo Template de Card';
    }

    // Preview em tempo real
    this.updatePreview();
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  },

  closeModal() {
    const modal = document.getElementById('templateModal');
    if (modal) modal.classList.add('hidden');
    document.body.style.overflow = '';
    this._editingId = null;
  },

  clearForm() {
    ['tplName','tplDescription','tplHeaderText','tplSlogan','tplFooterLogoText','tplBgImageUrl','tplBadgeLabel']
      .forEach(id => { const el = document.getElementById(id); if(el) el.value = ''; });
    const ltype = document.getElementById('tplLayoutType'); if(ltype) ltype.value = 'classic';
    const color = document.getElementById('tplHeaderColor'); if(color) color.value = '#2E7D32';
    const hex   = document.getElementById('tplHeaderColorHex'); if(hex) hex.textContent = '#2E7D32';
    const rank  = document.getElementById('tplShowRanking'); if(rank) rank.checked = false;
    const order = document.getElementById('tplSortOrder'); if(order) order.value = this._templates.length + 1;
    const act   = document.getElementById('tplActive'); if(act) act.checked = false;
  },

  // ===================================================
  // SETUP EVENTS
  // ===================================================
  setupEvents() {
    // Color picker live update
    const colorPicker = document.getElementById('tplHeaderColor');
    if (colorPicker) {
      colorPicker.addEventListener('input', (e) => {
        const hex = document.getElementById('tplHeaderColorHex');
        if (hex) hex.textContent = e.target.value;
        this.updatePreview();
      });
    }

    // Outros campos que disparam preview
    ['tplHeaderText','tplSlogan','tplFooterLogoText','tplLayoutType','tplBadgeLabel'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('input', () => this.updatePreview());
    });

    const rankCb = document.getElementById('tplShowRanking');
    if (rankCb) rankCb.addEventListener('change', () => this.updatePreview());

    // BG Image preview
    const bgInput = document.getElementById('tplBgImageUrl');
    if (bgInput) bgInput.addEventListener('blur', () => this.updatePreview());
  },

  // ===================================================
  // LIVE PREVIEW NO MODAL
  // ===================================================
  updatePreview() {
    const previewEl = document.getElementById('tplLivePreview');
    if (!previewEl) return;

    const headerColor = document.getElementById('tplHeaderColor')?.value || '#2E7D32';
    const headerText  = document.getElementById('tplHeaderText')?.value  || 'RESULTADOS DE PRODUTIVIDADE';
    const slogan      = document.getElementById('tplSlogan')?.value      || 'NÃO É SORTE! É STINE';
    const footerText  = document.getElementById('tplFooterLogoText')?.value || 'STINE';
    const bgUrl       = document.getElementById('tplBgImageUrl')?.value  || '';
    const badgeLabel  = document.getElementById('tplBadgeLabel')?.value  || '';
    const showRanking = document.getElementById('tplShowRanking')?.checked || false;

    previewEl.innerHTML = `
      <div class="stine-card stine-card-preview">
        <div class="stine-bg">
          ${bgUrl ? `<img src="${this.esc(bgUrl)}" alt="bg" style="width:100%;height:100%;object-fit:cover" crossorigin="anonymous" onerror="this.style.display='none'" />` : `<div style="width:100%;height:100%;background:linear-gradient(160deg,#0d2b0d,${headerColor})"></div>`}
          <div class="stine-bg-overlay-top"></div>
          <div class="stine-bg-overlay-bottom"></div>
        </div>

        <div class="stine-header" style="background:${headerColor}">
          <div class="stine-header-title">${this.esc(headerText)}</div>
          <div class="stine-header-sub">Soja &nbsp;|&nbsp; Safra 25/26</div>
        </div>

        ${showRanking ? `
          <div class="stine-ranking-badge">
            <div class="stine-ranking-medal">
              <div class="medal-ribbon-left"></div>
              <div class="medal-ribbon-right"></div>
              <div class="medal-circle"><span class="medal-number">1</span></div>
            </div>
          </div>` : ''}

        <div class="stine-panel">
          <div class="stine-logo-row">
            <div class="stine-logo">
              <div class="stine-logo-leaf">
                <svg viewBox="0 0 24 24" fill="${headerColor}">
                  <path d="M17 8C8 10 5.9 16.17 3.82 21c2.22-.9 4.77-1.5 6.18-2 2.5-1 6-4.5 6-4.5S14.5 17 10 19.5c2 0 4-1 5.5-2S19 13 17 8z"/>
                </svg>
              </div>
              <span class="stine-logo-text">STINE</span><sup class="stine-logo-reg">®</sup>
            </div>
            <div class="stine-tech-badge">
              <span class="tech-name">Conkesta E3</span>
              <span class="tech-crop">SOJA</span>
            </div>
          </div>
          <div class="stine-variety-code" style="color:${headerColor}">79KA72</div>
          <div class="stine-dates-row">
            <div class="stine-date-item"><span>10/10/2025</span></div>
            <div class="dates-dot">•</div>
            <div class="stine-date-item"><span>15/02/2026</span></div>
          </div>
          <div class="stine-productivity">
            <span class="prod-number" style="color:${headerColor}">110,5</span>
            <span class="prod-unit">sc/ha</span>
          </div>
          <div class="stine-location-ribbon">
            <svg class="pin-icon" viewBox="0 0 24 24">
              <path fill="#EF5350" d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5S10.62 6.5 12 6.5s2.5 1.12 2.5 2.5S13.38 11.5 12 11.5z"/>
            </svg>
            <span class="location-text">Nova Lacerda/MT</span>
          </div>
          <div class="stine-producer-block">
            <div class="producer-main">João da Silva</div>
            <div class="producer-farm">Fazenda São João</div>
            <div class="producer-area">25 ha | sequeiro</div>
          </div>
          ${badgeLabel ? `<div class="stine-obs-badge" style="background:${headerColor}">${this.esc(badgeLabel)}</div>` : '<div style="height:8px"></div>'}
        </div>

        <div class="stine-footer">
          <div class="footer-slogan"><span class="slogan-nao">NÃO É SORTE!</span></div>
          <div class="footer-brand">
            <span class="footer-e">É</span>
            <div class="footer-logo-wrap">
              <div class="footer-leaf">
                <svg viewBox="0 0 24 24" fill="white"><path d="M17 8C8 10 5.9 16.17 3.82 21c2.22-.9 4.77-1.5 6.18-2 2.5-1 6-4.5 6-4.5S14.5 17 10 19.5c2 0 4-1 5.5-2S19 13 17 8z"/></svg>
              </div>
              <span class="footer-brand-text">${this.esc(footerText)}</span>
              <sup class="footer-reg">®</sup>
            </div>
          </div>
        </div>
      </div>
    `;
  },

  // ===================================================
  // SAVE TEMPLATE
  // ===================================================
  async saveTemplate() {
    const data = {
      name:               document.getElementById('tplName')?.value.trim() || '',
      description:        document.getElementById('tplDescription')?.value.trim() || '',
      layout_type:        document.getElementById('tplLayoutType')?.value || 'classic',
      header_color:       document.getElementById('tplHeaderColor')?.value || '#2E7D32',
      header_text:        document.getElementById('tplHeaderText')?.value.trim() || 'RESULTADOS DE PRODUTIVIDADE',
      slogan:             document.getElementById('tplSlogan')?.value.trim() || 'NÃO É SORTE! É STINE',
      footer_logo_text:   document.getElementById('tplFooterLogoText')?.value.trim() || 'STINE',
      bg_image_url:       document.getElementById('tplBgImageUrl')?.value.trim() || '',
      badge_label:        document.getElementById('tplBadgeLabel')?.value.trim() || '',
      show_ranking_badge: document.getElementById('tplShowRanking')?.checked || false,
      sort_order:         parseInt(document.getElementById('tplSortOrder')?.value) || 99,
      active:             document.getElementById('tplActive')?.checked || false
    };

    if (!data.name) {
      App.Toast.show('Informe o nome do template.', 'error');
      return;
    }

    try {
      // Se marcado como ativo, desativa os demais
      if (data.active) {
        for (const t of this._templates) {
          if (t.active && t.id !== this._editingId) {
            await API.updateTemplate(t.id, { active: false });
          }
        }
      }

      if (this._editingId) {
        await API.updateTemplate(this._editingId, data);
        App.Toast.show('Template atualizado!', 'success');
      } else {
        await API.createTemplate(data);
        App.Toast.show('Template criado!', 'success');
      }

      this.closeModal();
      await this.loadTemplates();
      // Recarrega templates no CardGenerator
      CardGenerator._templates = [];

    } catch (err) {
      console.error(err);
      App.Toast.show('Erro ao salvar template.', 'error');
    }
  },

  // ===================================================
  // DELETE TEMPLATE
  // ===================================================
  async deleteTemplate(id) {
    const tpl = this._templates.find(t => t.id === id);
    if (tpl?.active) {
      App.Toast.show('Não é possível excluir o template ativo. Ative outro template primeiro.', 'error');
      return;
    }
    App.confirm('Excluir este template permanentemente?', async () => {
      try {
        await API.deleteTemplate(id);
        App.Toast.show('Template excluído.', 'info');
        await this.loadTemplates();
        CardGenerator._templates = [];
      } catch {
        App.Toast.show('Erro ao excluir template.', 'error');
      }
    });
  },

  // ===================================================
  // SET BACKGROUND FROM PRESET
  // ===================================================
  setBg(url) {
    const el = document.getElementById('tplBgImageUrl');
    if (el) {
      el.value = url;
      this.updatePreview();
    }
  },

  // ===================================================
  // HELPERS
  // ===================================================
  esc(str) {
    if (!str) return '';
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
};
