/* =====================================================
   AgriCard Stine – Settings Module v1.0
   Gerencia configurações do app, perfil do admin,
   troca de senha, gestão de admins e branding.
   ===================================================== */

const Settings = {

  _settings: {},   // cache key→value

  /* ═══════════════════════════════════════════════════
     INIT — carrega seção e popula formulários
  ═══════════════════════════════════════════════════ */
  async init() {
    await this._loadSettings();
    this._renderProfile();
    this._renderBranding();
    this._renderAdmins();
    await this.loadOneDriveConfig();
  },

  /* ─── Carrega todas as configurações do banco ─── */
  async _loadSettings() {
    try {
      const res = await fetch('tables/app_settings?limit=200');
      const data = await res.json();
      this._settings = {};
      (data.data || []).forEach(row => {
        this._settings[row.id] = row;
      });
    } catch (e) {
      console.error('Settings._loadSettings:', e);
    }
  },

  _get(key, fallback = '') {
    return this._settings[key]?.value ?? fallback;
  },

  async _set(key, value, label = '', category = 'system') {
    const existing = this._settings[key];
    const payload  = { id: key, value, label, category, updated_by: Auth.currentUser?.id || '' };
    let row;
    if (existing) {
      const res = await fetch(`tables/app_settings/${existing.id_row || existing.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value, updated_by: payload.updated_by })
      });
      row = await res.json();
    } else {
      const res = await fetch('tables/app_settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      row = await res.json();
    }
    this._settings[key] = { ...payload, ...row };
    return row;
  },

  /* ═══════════════════════════════════════════════════
     PROFILE — editar dados do admin logado
  ═══════════════════════════════════════════════════ */
  _renderProfile() {
    const u = Auth.currentUser;
    if (!u) return;
    this._setVal('settingAdminName',    u.name    || '');
    this._setVal('settingAdminEmail',   u.email   || '');
    this._setVal('settingAdminCompany', u.company || '');
    this._setVal('settingAdminPhone',   u.phone   || '');
    this._setVal('settingAdminRegion',  u.region  || '');

    // Avatar inicial
    const av = document.getElementById('settingProfileAvatar');
    if (av) av.textContent = (u.name || 'A').charAt(0).toUpperCase();
  },

  async saveProfile() {
    const u = Auth.currentUser;
    if (!u) return;

    const name    = this._getVal('settingAdminName').trim();
    const email   = this._getVal('settingAdminEmail').trim().toLowerCase();
    const company = this._getVal('settingAdminCompany').trim();
    const phone   = this._getVal('settingAdminPhone').trim();
    const region  = this._getVal('settingAdminRegion').trim();

    if (!name || !email) {
      App.Toast.show('Nome e e-mail são obrigatórios.', 'error'); return;
    }

    const btn = document.getElementById('btnSaveProfile');
    this._btnLoading(btn, 'Salvando...');

    try {
      const updated = await API.updateUser(u.id, { name, email, company, phone, region });
      // Atualiza sessão local
      const newUser = { ...u, name, email, company, phone, region };
      Auth.currentUser = newUser;
      localStorage.setItem('agricard_user', JSON.stringify(newUser));

      // Atualiza nome no sidebar
      document.getElementById('adminUserName').textContent = name;
      const av = document.getElementById('settingProfileAvatar');
      if (av) av.textContent = name.charAt(0).toUpperCase();

      App.Toast.show('Perfil atualizado com sucesso!', 'success');
    } catch (e) {
      App.Toast.show('Erro ao salvar perfil.', 'error');
    } finally {
      this._btnReset(btn, '<i class="fas fa-save"></i> Salvar Perfil');
    }
  },

  /* ═══════════════════════════════════════════════════
     PASSWORD — trocar senha do admin logado
  ═══════════════════════════════════════════════════ */
  async changePassword() {
    const current  = this._getVal('settingCurrentPass');
    const newPass  = this._getVal('settingNewPass');
    const confirm  = this._getVal('settingConfirmPass');
    const u        = Auth.currentUser;

    if (!current || !newPass || !confirm) {
      App.Toast.show('Preencha todos os campos de senha.', 'error'); return;
    }
    if (u.password !== current) {
      App.Toast.show('Senha atual incorreta.', 'error'); return;
    }
    if (newPass.length < 6) {
      App.Toast.show('Nova senha deve ter pelo menos 6 caracteres.', 'error'); return;
    }
    if (newPass !== confirm) {
      App.Toast.show('As novas senhas não conferem.', 'error'); return;
    }

    const btn = document.getElementById('btnChangePassword');
    this._btnLoading(btn, 'Salvando...');

    try {
      await API.updateUser(u.id, { password: newPass });
      const newUser = { ...u, password: newPass };
      Auth.currentUser = newUser;
      localStorage.setItem('agricard_user', JSON.stringify(newUser));

      ['settingCurrentPass','settingNewPass','settingConfirmPass']
        .forEach(id => this._setVal(id, ''));

      App.Toast.show('Senha alterada com sucesso!', 'success');
    } catch {
      App.Toast.show('Erro ao alterar senha.', 'error');
    } finally {
      this._btnReset(btn, '<i class="fas fa-key"></i> Alterar Senha');
    }
  },

  /* ═══════════════════════════════════════════════════
     BRANDING — logo e identidade visual do app
  ═══════════════════════════════════════════════════ */
  _renderBranding() {
    this._setVal('settingAppName',    this._get('app_name',    'AgriCard'));
    this._setVal('settingAppSubname', this._get('app_subname', 'STINE Sementes'));
    this._setVal('settingAppTagline', this._get('app_tagline', 'Plataforma de Cards de Produtividade Agrícola'));
    this._setVal('settingPrimaryColor', this._get('primary_color', '#2E7D32'));

    // Mostra preview do logo se existir
    const logoUrl = this._get('logo_url', '');
    this._updateLogoPreview(logoUrl);
  },

  _updateLogoPreview(src) {
    const preview = document.getElementById('settingLogoPreview');
    if (!preview) return;
    if (src) {
      preview.innerHTML = `<img src="${this.esc(src)}" alt="Logo" style="max-height:80px;max-width:200px;border-radius:8px;object-fit:contain" />`;
    } else {
      preview.innerHTML = `<div class="setting-logo-placeholder"><i class="fas fa-image"></i><span>Nenhum logo cadastrado</span></div>`;
    }
  },

  async handleLogoUpload(file) {
    if (!file) return;
    const allowedTypes = ['image/png','image/jpeg','image/jpg','image/webp','image/svg+xml'];
    if (!allowedTypes.includes(file.type)) {
      App.Toast.show('Formato inválido. Use PNG, JPG, WEBP ou SVG.', 'error'); return;
    }
    if (file.size > 2 * 1024 * 1024) {
      App.Toast.show('Imagem muito grande. Máximo 2MB.', 'error'); return;
    }

    const reader = new FileReader();
    reader.onload = async e => {
      const dataUrl = e.target.result;
      this._updateLogoPreview(dataUrl);
      this._setVal('settingLogoUrl', dataUrl);
      App.Toast.show('Logo carregado! Clique em "Salvar Identidade Visual" para confirmar.', 'info');
    };
    reader.readAsDataURL(file);
  },

  removeLogo() {
    this._setVal('settingLogoUrl', '');
    this._updateLogoPreview('');
    App.Toast.show('Logo removido. Clique em "Salvar Identidade Visual" para confirmar.', 'info');
  },

  async saveBranding() {
    const appName    = this._getVal('settingAppName').trim();
    const appSubname = this._getVal('settingAppSubname').trim();
    const appTagline = this._getVal('settingAppTagline').trim();
    const logoUrl    = this._getVal('settingLogoUrl').trim();
    const primaryColor = this._getVal('settingPrimaryColor') || '#2E7D32';

    if (!appName) { App.Toast.show('Nome do app é obrigatório.', 'error'); return; }

    const btn = document.getElementById('btnSaveBranding');
    this._btnLoading(btn, 'Salvando...');

    try {
      await Promise.all([
        this._set('app_name',     appName,     'Nome do App',      'branding'),
        this._set('app_subname',  appSubname,  'Subtítulo',        'branding'),
        this._set('app_tagline',  appTagline,  'Tagline',          'branding'),
        this._set('logo_url',     logoUrl,     'Logo URL/Base64',  'branding'),
        this._set('primary_color', primaryColor, 'Cor Principal',  'branding'),
      ]);

      // Aplica no DOM imediatamente
      AppBranding.apply();
      App.Toast.show('Identidade visual atualizada!', 'success');
    } catch {
      App.Toast.show('Erro ao salvar identidade visual.', 'error');
    } finally {
      this._btnReset(btn, '<i class="fas fa-paint-brush"></i> Salvar Identidade Visual');
    }
  },

  /* ═══════════════════════════════════════════════════
     ADMINS — listar, criar, excluir admins
  ═══════════════════════════════════════════════════ */
  async _renderAdmins() {
    const container = document.getElementById('adminsListContainer');
    if (!container) return;

    container.innerHTML = `<div style="padding:16px;text-align:center;color:#888"><span class="loading"></span> Carregando...</div>`;

    try {
      const res    = await API.getUsers();
      const admins = (res.data || []).filter(u => u.role === 'admin');
      const me     = Auth.currentUser;

      if (admins.length === 0) {
        container.innerHTML = `<div class="empty-state"><i class="fas fa-user-shield"></i><p>Nenhum administrador cadastrado</p></div>`;
        return;
      }

      container.innerHTML = `
        <table class="data-table">
          <thead>
            <tr>
              <th>Nome</th>
              <th>E-mail</th>
              <th>Empresa</th>
              <th>Ações</th>
            </tr>
          </thead>
          <tbody>
            ${admins.map(a => `
              <tr class="${a.id === me?.id ? 'settings-row-me' : ''}">
                <td>
                  <div style="display:flex;align-items:center;gap:10px">
                    <div class="settings-admin-avatar">${(a.name||'A').charAt(0).toUpperCase()}</div>
                    <div>
                      <strong>${this.esc(a.name)}</strong>
                      ${a.id === me?.id ? ' <span class="badge badge-approved" style="font-size:10px">Você</span>' : ''}
                    </div>
                  </div>
                </td>
                <td>${this.esc(a.email)}</td>
                <td>${this.esc(a.company || '—')}</td>
                <td>
                  ${a.id !== me?.id ? `
                    <button class="action-btn action-btn-red" onclick="Settings.deleteAdmin('${a.id}', '${this.esc(a.name)}')">
                      <i class="fas fa-user-minus"></i> Remover Admin
                    </button>
                  ` : `<span style="font-size:12px;color:var(--gray-400)">Conta atual</span>`}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>`;
    } catch {
      container.innerHTML = `<div class="empty-state"><i class="fas fa-exclamation-circle"></i><p>Erro ao carregar administradores</p></div>`;
    }
  },

  async addAdmin() {
    const name    = this._getVal('newAdminName').trim();
    const email   = this._getVal('newAdminEmail').trim().toLowerCase();
    const company = this._getVal('newAdminCompany').trim();
    const pass    = this._getVal('newAdminPass');
    const confirm = this._getVal('newAdminPassConfirm');

    if (!name || !email || !pass) {
      App.Toast.show('Nome, e-mail e senha são obrigatórios.', 'error'); return;
    }
    if (pass.length < 6) {
      App.Toast.show('Senha deve ter pelo menos 6 caracteres.', 'error'); return;
    }
    if (pass !== confirm) {
      App.Toast.show('As senhas não conferem.', 'error'); return;
    }

    const btn = document.getElementById('btnAddAdmin');
    this._btnLoading(btn, 'Criando...');

    try {
      // Verifica duplicidade
      const check = await API.getUsers(`search=${encodeURIComponent(email)}`);
      const dup   = (check.data || []).find(u => u.email?.toLowerCase() === email);
      if (dup) {
        App.Toast.show('Este e-mail já está cadastrado.', 'error'); return;
      }

      await API.createUser({
        name, email, company,
        password: pass,
        role: 'admin',
        status: 'approved',
        phone: '', region: ''
      });

      // Limpa form
      ['newAdminName','newAdminEmail','newAdminCompany','newAdminPass','newAdminPassConfirm']
        .forEach(id => this._setVal(id, ''));

      App.Toast.show(`Admin "${name}" criado com sucesso!`, 'success');
      await this._renderAdmins();
      // Fecha accordion se estiver aberto
      const accordion = document.getElementById('addAdminAccordion');
      if (accordion) accordion.classList.remove('open');

    } catch {
      App.Toast.show('Erro ao criar administrador.', 'error');
    } finally {
      this._btnReset(btn, '<i class="fas fa-user-plus"></i> Criar Administrador');
    }
  },

  deleteAdmin(id, name) {
    if (id === Auth.currentUser?.id) {
      App.Toast.show('Não é possível remover a sua própria conta.', 'error'); return;
    }
    App.confirm(`Remover o administrador "${name}"? Esta ação não pode ser desfeita.`, async () => {
      try {
        await API.deleteUser(id);
        App.Toast.show(`Admin "${name}" removido.`, 'info');
        await this._renderAdmins();
      } catch {
        App.Toast.show('Erro ao remover administrador.', 'error');
      }
    });
  },

  toggleAddAdminForm() {
    const accordion = document.getElementById('addAdminAccordion');
    if (accordion) accordion.classList.toggle('open');
  },

  /* ═══════════════════════════════════════════════════
     ONEDRIVE — configuração de integração
  ═══════════════════════════════════════════════════ */
  async loadOneDriveConfig() {
    try {
      const res  = await fetch('tables/onedrive_config?limit=50');
      const data = await res.json();
      const map  = {};
      (data.data || []).forEach(r => { map[r.key] = r; });

      this._odConfig = map;

      const clientId   = map['client_id']?.value   || '';
      const tenantId   = map['tenant_id']?.value    || 'common';
      const pastaTermos = map['pasta_termos']?.value || '/Cards_Produtividade/Termos/';
      const pastaCards  = map['pasta_cards']?.value  || '/Cards_Produtividade/Cards/';
      const enabled     = map['enabled']?.value === 'true';

      this._setVal('odClientId',    clientId);
      this._setVal('odTenantId',    tenantId);
      this._setVal('odPastaTermos', pastaTermos);
      this._setVal('odPastaCards',  pastaCards);

      const odEnabledEl = document.getElementById('odEnabled');
      if (odEnabledEl) odEnabledEl.checked = enabled;

      // Mostra redirect URI
      const redirectEl = document.getElementById('odRedirectUriDisplay');
      if (redirectEl) {
        redirectEl.textContent = window.location.origin + '/auth-callback.html';
      }

      // Atualiza status badge
      this._updateOneDriveStatus(enabled && !!clientId);

    } catch (e) {
      console.error('Settings.loadOneDriveConfig:', e);
    }
  },

  _odConfig: {},

  async saveOneDriveConfig() {
    const clientId    = this._getVal('odClientId').trim();
    const tenantId    = this._getVal('odTenantId').trim() || 'common';
    const pastaTermos = this._getVal('odPastaTermos').trim() || '/Cards_Produtividade/Termos/';
    const pastaCards  = this._getVal('odPastaCards').trim()  || '/Cards_Produtividade/Cards/';
    const enabled     = document.getElementById('odEnabled')?.checked ? 'true' : 'false';

    if (enabled === 'true' && !clientId) {
      App.Toast.show('O Client ID é obrigatório para habilitar o OneDrive.', 'error');
      return;
    }

    const btn = document.getElementById('btnSaveOneDrive');
    this._btnLoading(btn, 'Salvando...');

    try {
      const configs = [
        { key: 'client_id',   value: clientId,    label: 'Client ID (Azure App)' },
        { key: 'tenant_id',   value: tenantId,     label: 'Tenant ID' },
        { key: 'pasta_termos', value: pastaTermos, label: 'Pasta Termos' },
        { key: 'pasta_cards',  value: pastaCards,  label: 'Pasta Cards' },
        { key: 'enabled',     value: enabled,      label: 'Habilitado' },
        { key: 'redirect_uri', value: window.location.origin + '/auth-callback.html', label: 'Redirect URI' }
      ];

      for (const cfg of configs) {
        const existing = this._odConfig[cfg.key];
        if (existing?.id) {
          await fetch(`tables/onedrive_config/${existing.id}`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ value: cfg.value, updated_by: Auth.currentUser?.id || '' })
          });
        } else {
          const newRow = await fetch('tables/onedrive_config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              key: cfg.key,
              value: cfg.value,
              label: cfg.label,
              category: 'onedrive',
              updated_by: Auth.currentUser?.id || ''
            })
          });
          const r = await newRow.json();
          this._odConfig[cfg.key] = r;
        }
      }

      // Atualiza módulo OneDrive em memória
      if (typeof OneDrive !== 'undefined') {
        OneDrive._config = {
          client_id:    clientId,
          tenant_id:    tenantId,
          pasta_termos: pastaTermos,
          pasta_cards:  pastaCards,
          enabled:      enabled,
          redirect_uri: window.location.origin + '/auth-callback.html'
        };
      }

      this._updateOneDriveStatus(enabled === 'true' && !!clientId);
      App.Toast.show('Configuração OneDrive salva com sucesso!', 'success');

    } catch (e) {
      App.Toast.show('Erro ao salvar configuração OneDrive.', 'error');
      console.error(e);
    } finally {
      this._btnReset(btn, '<i class="fas fa-save"></i> Salvar Configuração OneDrive');
    }
  },

  async testOneDriveConnection() {
    if (typeof OneDrive === 'undefined') {
      App.Toast.show('Módulo OneDrive não disponível.', 'error');
      return;
    }

    const clientId = this._getVal('odClientId').trim();
    if (!clientId) {
      App.Toast.show('Configure o Client ID primeiro.', 'error');
      return;
    }

    App.Toast.show('Iniciando autenticação com o OneDrive...', 'info');

    try {
      // Garante que a config está carregada
      if (!OneDrive._config) await OneDrive.loadConfig();

      // Atualiza config com os valores atuais do form
      if (!OneDrive._config) OneDrive._config = {};
      OneDrive._config.client_id    = clientId;
      OneDrive._config.tenant_id    = this._getVal('odTenantId').trim() || 'common';
      OneDrive._config.redirect_uri = window.location.origin + '/auth-callback.html';
      OneDrive._config.enabled      = 'true';

      await OneDrive.authenticate();
      this._updateOneDriveStatus(true, 'Autenticado ✓');
      App.Toast.show('✅ Conexão com OneDrive estabelecida com sucesso!', 'success');

    } catch (err) {
      this._updateOneDriveStatus(false);
      App.Toast.show('Falha na conexão: ' + err.message, 'error');
    }
  },

  _updateOneDriveStatus(connected, label) {
    const el = document.getElementById('onedriveSyncStatus');
    if (!el) return;
    if (connected) {
      el.style.background = '#dcfce7';
      el.style.color      = '#166534';
      el.innerHTML = `<i class="fas fa-circle" style="font-size:8px;color:#22c55e"></i> ${label || 'Configurado'}`;
    } else {
      el.style.background = '#e2e8f0';
      el.style.color      = '#64748b';
      el.innerHTML = `<i class="fas fa-circle" style="font-size:8px"></i> Desconectado`;
    }
  },

  /* ═══════════════════════════════════════════════════
     DANGER ZONE — ações destrutivas
  ═══════════════════════════════════════════════════ */
  clearAllRecords() {
    App.confirm('⚠️ Excluir TODOS os registros de produtividade? Esta ação é irreversível!', async () => {
      try {
        const res     = await API.getRecords();
        const records = res.data || [];
        await Promise.all(records.map(r => API.deleteRecord(r.id)));
        App.Toast.show(`${records.length} registros excluídos.`, 'warning');
      } catch {
        App.Toast.show('Erro ao excluir registros.', 'error');
      }
    });
  },

  clearAllVarieties() {
    App.confirm('⚠️ Excluir TODAS as variedades e seus modelos de card? Esta ação é irreversível!', async () => {
      try {
        const res  = await API.getVarieties();
        const vars = res.data || [];
        await Promise.all(vars.map(v => API.deleteVariety(v.id)));
        CardGenerator._invalidateCache();
        App.Toast.show(`${vars.length} variedades excluídas.`, 'warning');
      } catch {
        App.Toast.show('Erro ao excluir variedades.', 'error');
      }
    });
  },

  /* ═══════════════════════════════════════════════════
     HELPERS
  ═══════════════════════════════════════════════════ */
  _setVal(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value;
  },
  _getVal(id) {
    return document.getElementById(id)?.value || '';
  },
  _btnLoading(btn, text) {
    if (!btn) return;
    btn._orig = btn.innerHTML;
    btn.innerHTML = `<span class="loading"></span> ${text}`;
    btn.disabled  = true;
  },
  _btnReset(btn, html) {
    if (!btn) return;
    btn.innerHTML = html || btn._orig || 'OK';
    btn.disabled  = false;
  },
  esc(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
};

/* ═══════════════════════════════════════════════════
   AppBranding — aplica logo/nome dinamicamente no DOM
═══════════════════════════════════════════════════ */
const AppBranding = {

  async load() {
    try {
      const res  = await fetch('tables/app_settings?limit=200');
      const data = await res.json();
      const rows = data.data || [];
      const map  = {};
      rows.forEach(r => { map[r.id] = r.value; });
      this._cache = map;
    } catch {
      this._cache = {};
    }
    this.apply();
  },

  _cache: {},

  get(key, fallback = '') {
    return this._cache[key] ?? fallback;
  },

  apply() {
    const map = Settings._settings || {};
    const get = (k, fb) => map[k]?.value ?? this._cache[k] ?? fb;

    const appName    = get('app_name',    'AgriCard');
    const appSubname = get('app_subname', 'STINE Sementes');
    const appTagline = get('app_tagline', 'Plataforma de Cards de Produtividade Agrícola');
    const logoUrl    = get('logo_url',    '');
    const primaryColor = get('primary_color', '');

    // ---------- Tela de login ----------
    const authLogoIcon  = document.querySelector('.auth-logo-icon');
    const authLogoBrand = document.querySelector('.auth-logo-brand');
    const authLogoSub   = document.querySelector('.auth-logo-sub');
    const authSubtitle  = document.querySelector('.auth-subtitle');

    if (logoUrl && authLogoIcon) {
      authLogoIcon.innerHTML = `<img src="${logoUrl}" alt="${appName}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit" />`;
    } else if (authLogoIcon) {
      authLogoIcon.innerHTML = '<i class="fas fa-seedling"></i>';
    }
    if (authLogoBrand)  authLogoBrand.textContent  = appName;
    if (authLogoSub)    authLogoSub.textContent    = appSubname;
    if (authSubtitle)   authSubtitle.textContent   = appTagline;

    // ---------- Sidebar Admin ----------
    const sbIconAdmin  = document.querySelector('#adminSidebar .sidebar-logo-icon');
    const sbBrandAdmin = document.querySelector('#adminSidebar .sidebar-brand');
    const sbSubAdmin   = document.querySelector('#adminSidebar .sidebar-sub');

    if (logoUrl && sbIconAdmin) {
      sbIconAdmin.innerHTML = `<img src="${logoUrl}" alt="${appName}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit" />`;
    } else if (sbIconAdmin) {
      sbIconAdmin.innerHTML = '<i class="fas fa-seedling"></i>';
    }
    if (sbBrandAdmin) sbBrandAdmin.textContent = appName;
    if (sbSubAdmin)   sbSubAdmin.textContent   = appSubname;

    // ---------- Mobile header Admin ----------
    const mobLogoAdmin = document.querySelector('#adminScreen .mobile-logo');
    if (mobLogoAdmin) {
      if (logoUrl) {
        mobLogoAdmin.innerHTML = `<img src="${logoUrl}" alt="${appName}" style="height:26px;object-fit:contain" /> ${appName}`;
      } else {
        mobLogoAdmin.innerHTML = `<i class="fas fa-seedling"></i> ${appName}`;
      }
    }

    // ---------- Sidebar Usuário ----------
    const sbIconUser  = document.querySelector('#userSidebar .sidebar-logo-icon');
    const sbBrandUser = document.querySelector('#userSidebar .sidebar-brand');
    const sbSubUser   = document.querySelector('#userSidebar .sidebar-sub');

    if (logoUrl && sbIconUser) {
      sbIconUser.innerHTML = `<img src="${logoUrl}" alt="${appName}" style="width:100%;height:100%;object-fit:contain;border-radius:inherit" />`;
    } else if (sbIconUser) {
      sbIconUser.innerHTML = '<i class="fas fa-seedling"></i>';
    }
    if (sbBrandUser) sbBrandUser.textContent = appName;
    if (sbSubUser)   sbSubUser.textContent   = appSubname;

    // ---------- Mobile header Usuário ----------
    const mobLogoUser = document.querySelector('#userScreen .mobile-logo');
    if (mobLogoUser) {
      if (logoUrl) {
        mobLogoUser.innerHTML = `<img src="${logoUrl}" alt="${appName}" style="height:26px;object-fit:contain" /> ${appName}`;
      } else {
        mobLogoUser.innerHTML = `<i class="fas fa-seedling"></i> ${appName}`;
      }
    }

    // ---------- <title> da página ----------
    document.title = `${appName} – ${appSubname}`;

    // ---------- Cor primária (CSS var) ----------
    if (primaryColor && /^#[0-9a-fA-F]{6}$/.test(primaryColor)) {
      document.documentElement.style.setProperty('--green', primaryColor);
    }
  }
};
