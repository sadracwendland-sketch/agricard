/* =====================================================
   AgriCard Stine – OneDrive Integration Module v1.0
   
   Usa Microsoft Graph API via OAuth2 (PKCE flow)
   para enviar arquivos ao OneDrive pessoal / corporativo.
   
   Fluxo:
   1. Admin configura Client ID, Tenant ID e pasta base
   2. Usuário se autentica via popup OAuth2
   3. Módulo envia termos e cards automaticamente
   ===================================================== */

'use strict';

const OneDrive = {

  /* ── Estado interno ── */
  _accessToken:  null,
  _tokenExpiry:  null,
  _config:       null,
  _authInProgress: false,

  /* ── Escopos necessários ── */
  SCOPES: 'Files.ReadWrite offline_access User.Read',
  GRAPH:  'https://graph.microsoft.com/v1.0',

  /* ═══════════════════════════════════════════════════
     CONFIGURAÇÃO
  ═══════════════════════════════════════════════════ */

  /**
   * Carrega configurações do banco de dados
   */
  async loadConfig() {
    try {
      const res = await fetch('tables/onedrive_config?limit=20');
      const data = await res.json();
      const rows = data.data || [];
      this._config = {};
      rows.forEach(r => { this._config[r.key] = r.value || ''; });
      return this._config;
    } catch (e) {
      console.warn('[OneDrive] Falha ao carregar config:', e);
      return {};
    }
  },

  /**
   * Verifica se OneDrive está habilitado e configurado
   */
  async isEnabled() {
    if (!this._config) await this.loadConfig();
    return this._config.enabled === 'true' && !!this._config.client_id;
  },

  /* ═══════════════════════════════════════════════════
     AUTENTICAÇÃO OAuth2 PKCE
  ═══════════════════════════════════════════════════ */

  /**
   * Verifica se o token ainda é válido
   */
  isTokenValid() {
    return this._accessToken && this._tokenExpiry && Date.now() < this._tokenExpiry;
  },

  /**
   * Obtém o token de acesso (do storage ou autentica)
   */
  async getToken() {
    // Verifica token em memória
    if (this.isTokenValid()) return this._accessToken;

    // Verifica token no sessionStorage
    try {
      const stored = sessionStorage.getItem('od_token');
      if (stored) {
        const t = JSON.parse(stored);
        if (t.expiry > Date.now()) {
          this._accessToken = t.token;
          this._tokenExpiry = t.expiry;
          return this._accessToken;
        }
      }
    } catch {}

    // Precisa autenticar
    return await this.authenticate();
  },

  /**
   * Inicia fluxo OAuth2 PKCE via popup
   */
  async authenticate() {
    if (!this._config) await this.loadConfig();
    
    const clientId  = this._config.client_id;
    const tenantId  = this._config.tenant_id || 'common';
    const redirectUri = this._config.redirect_uri || window.location.origin + '/auth-callback.html';

    if (!clientId) {
      throw new Error('Client ID do OneDrive não configurado. Acesse Admin → Configurações → OneDrive.');
    }

    if (this._authInProgress) {
      throw new Error('Autenticação já em andamento.');
    }

    this._authInProgress = true;

    try {
      // Gera code_verifier e code_challenge (PKCE)
      const verifier   = this._generateCodeVerifier();
      const challenge  = await this._generateCodeChallenge(verifier);
      const state      = this._generateState();

      sessionStorage.setItem('od_pkce_verifier', verifier);
      sessionStorage.setItem('od_pkce_state', state);

      const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
        `client_id=${encodeURIComponent(clientId)}&` +
        `response_type=code&` +
        `redirect_uri=${encodeURIComponent(redirectUri)}&` +
        `scope=${encodeURIComponent(this.SCOPES)}&` +
        `state=${state}&` +
        `code_challenge=${challenge}&` +
        `code_challenge_method=S256&` +
        `prompt=select_account`;

      // Abre popup
      const popup = window.open(authUrl, 'od_auth',
        'width=500,height=600,left=200,top=100,toolbar=no,menubar=no');

      if (!popup) throw new Error('Popup bloqueado. Permita popups para este site.');

      // Aguarda callback
      const code = await this._waitForAuthCode(popup, state);
      const token = await this._exchangeCode(code, verifier, redirectUri, clientId, tenantId);

      this._accessToken = token.access_token;
      this._tokenExpiry = Date.now() + (token.expires_in * 1000) - 60000;

      sessionStorage.setItem('od_token', JSON.stringify({
        token: this._accessToken,
        expiry: this._tokenExpiry
      }));

      return this._accessToken;

    } finally {
      this._authInProgress = false;
    }
  },

  /**
   * Aguarda o código de autorização do popup
   */
  _waitForAuthCode(popup, expectedState) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        popup?.close();
        reject(new Error('Tempo limite de autenticação esgotado.'));
      }, 120000);

      const interval = setInterval(() => {
        try {
          if (popup.closed) {
            clearInterval(interval);
            clearTimeout(timeout);
            reject(new Error('Autenticação cancelada pelo usuário.'));
            return;
          }

          let url;
          try { url = popup.location.href; } catch { return; }

          if (url && (url.includes('code=') || url.includes('error='))) {
            clearInterval(interval);
            clearTimeout(timeout);
            popup.close();

            const params = new URL(url).searchParams;
            if (params.get('error')) {
              reject(new Error('Erro OAuth: ' + params.get('error_description')));
            } else if (params.get('state') !== expectedState) {
              reject(new Error('Estado OAuth inválido (possível ataque CSRF).'));
            } else {
              resolve(params.get('code'));
            }
          }
        } catch {}
      }, 500);
    });
  },

  /**
   * Troca o código de autorização por um access token
   */
  async _exchangeCode(code, verifier, redirectUri, clientId, tenantId) {
    const body = new URLSearchParams({
      client_id:     clientId,
      code,
      redirect_uri:  redirectUri,
      grant_type:    'authorization_code',
      code_verifier: verifier,
      scope:         this.SCOPES,
    });

    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
    );

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error('Falha ao obter token: ' + (err.error_description || res.status));
    }

    return await res.json();
  },

  /**
   * Faz logout (limpa tokens)
   */
  logout() {
    this._accessToken = null;
    this._tokenExpiry = null;
    sessionStorage.removeItem('od_token');
    sessionStorage.removeItem('od_pkce_verifier');
    sessionStorage.removeItem('od_pkce_state');
  },

  /* ═══════════════════════════════════════════════════
     OPERAÇÕES NO ONEDRIVE
  ═══════════════════════════════════════════════════ */

  /**
   * Cria pasta (e subpastas) no OneDrive se não existir
   * @param {string} fullPath - ex: /Cards_Produtividade/Termos/2026/Soja
   */
  async ensureFolderPath(fullPath) {
    const token  = await this.getToken();
    // Divide o caminho em partes
    const parts  = fullPath.replace(/^\//, '').split('/').filter(Boolean);
    
    let parentId = 'root';
    for (const part of parts) {
      parentId = await this._ensureChildFolder(token, parentId, part);
    }
    return parentId;
  },

  /**
   * Garante que a pasta filha existe dentro de um pai
   */
  async _ensureChildFolder(token, parentId, name) {
    const parentRef = parentId === 'root' ? 'root' : `items/${parentId}`;
    
    // Verifica se já existe
    const searchUrl = `${this.GRAPH}/me/drive/${parentRef}/children?$filter=name eq '${encodeURIComponent(name)}' and folder ne null`;
    const res = await fetch(searchUrl, { headers: { 'Authorization': `Bearer ${token}` } });
    
    if (res.ok) {
      const data = await res.json();
      if (data.value && data.value.length > 0) return data.value[0].id;
    }

    // Cria a pasta
    const createUrl = `${this.GRAPH}/me/drive/${parentRef}/children`;
    const createRes = await fetch(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/json'
      },
      body: JSON.stringify({
        name,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'rename'
      })
    });

    if (!createRes.ok) {
      const err = await createRes.json().catch(() => ({}));
      throw new Error(`Falha ao criar pasta '${name}': ${err.error?.message || createRes.status}`);
    }

    const folder = await createRes.json();
    return folder.id;
  },

  /**
   * Faz upload de um arquivo para o OneDrive
   * @param {string|Blob} content - string base64 ou Blob
   * @param {string} fullPath     - caminho completo incluindo filename
   * @param {string} mimeType     - ex: 'image/png', 'application/pdf'
   * @returns {{ id, webUrl, name, path }}
   */
  async uploadFile(content, fullPath, mimeType = 'application/octet-stream') {
    const token = await this.getToken();

    // Converte base64 para Blob se necessário
    let blob;
    if (typeof content === 'string' && content.startsWith('data:')) {
      blob = this._dataUrlToBlob(content);
    } else if (typeof content === 'string') {
      blob = new Blob([content], { type: mimeType });
    } else {
      blob = content;
    }

    // Upload via URL de caminho (simples para arquivos < 4MB)
    const encodedPath = fullPath.replace(/^\//, '').split('/').map(encodeURIComponent).join('/');
    
    let uploadUrl;
    if (blob.size < 4 * 1024 * 1024) {
      // Upload simples
      uploadUrl = `${this.GRAPH}/me/drive/root:/${encodedPath}:/content`;
      const res = await fetch(uploadUrl, {
        method:  'PUT',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type':  mimeType
        },
        body: blob
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(`Upload falhou: ${err.error?.message || res.status}`);
      }

      const item = await res.json();
      return {
        id:     item.id,
        webUrl: item.webUrl,
        name:   item.name,
        path:   fullPath
      };
    } else {
      // Upload em sessão (large files)
      return await this._uploadLargeFile(token, blob, fullPath, mimeType);
    }
  },

  /**
   * Upload em sessão para arquivos > 4MB
   */
  async _uploadLargeFile(token, blob, fullPath, mimeType) {
    const encodedPath = fullPath.replace(/^\//, '').split('/').map(encodeURIComponent).join('/');
    
    // Cria sessão de upload
    const sessionRes = await fetch(
      `${this.GRAPH}/me/drive/root:/${encodedPath}:/createUploadSession`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type':  'application/json'
        },
        body: JSON.stringify({
          item: {
            '@microsoft.graph.conflictBehavior': 'replace',
            name: fullPath.split('/').pop()
          }
        })
      }
    );

    if (!sessionRes.ok) throw new Error('Falha ao criar sessão de upload.');
    const { uploadUrl } = await sessionRes.json();

    // Envia em chunks de 5MB
    const chunkSize = 5 * 1024 * 1024;
    let offset = 0;
    let result = null;

    while (offset < blob.size) {
      const chunk = blob.slice(offset, Math.min(offset + chunkSize, blob.size));
      const end   = Math.min(offset + chunkSize - 1, blob.size - 1);

      const chunkRes = await fetch(uploadUrl, {
        method:  'PUT',
        headers: {
          'Content-Range': `bytes ${offset}-${end}/${blob.size}`,
          'Content-Length': chunk.size.toString()
        },
        body: chunk
      });

      if (chunkRes.status === 201 || chunkRes.status === 200) {
        result = await chunkRes.json();
      } else if (chunkRes.status === 202) {
        // Ainda enviando
      } else {
        throw new Error(`Falha no chunk de upload: ${chunkRes.status}`);
      }

      offset += chunkSize;
    }

    return {
      id:     result?.id,
      webUrl: result?.webUrl,
      name:   result?.name,
      path:   fullPath
    };
  },

  /* ═══════════════════════════════════════════════════
     FLUXO ESPECÍFICO: TERMOS E CARDS
  ═══════════════════════════════════════════════════ */

  /**
   * Envia o termo de autorização para o OneDrive
   * @param {File} file       - arquivo original
   * @param {object} record   - dados do registro
   * @param {string} nomeFile - nome padronizado
   * @returns {{ id, webUrl, path }} ou null se desabilitado
   */
  async uploadTermo(file, record, nomeFile) {
    if (!(await this.isEnabled())) return null;

    const pastaBase = this._config.pasta_termos || '/Cards_Produtividade/Termos/';
    const ano       = this._extractYear(record);
    const cultura   = this._sanitize(record.culture || 'Geral');
    const pasta     = `${pastaBase.replace(/\/$/, '')}/${ano}/${cultura}`;

    await this.ensureFolderPath(pasta);

    const fullPath = `${pasta}/${nomeFile}`;
    const blob     = file instanceof File ? file : this._dataUrlToBlob(file);
    const mime     = file.type || 'application/octet-stream';

    return await this.uploadFile(blob, fullPath, mime);
  },

  /**
   * Envia o card gerado (PNG base64) para o OneDrive
   * @param {string} pngDataUrl - base64 PNG
   * @param {object} record     - dados do registro
   * @param {string} nomeFile   - nome padronizado
   * @returns {{ id, webUrl, path }} ou null se desabilitado
   */
  async uploadCard(pngDataUrl, record, nomeFile) {
    if (!(await this.isEnabled())) return null;

    const pastaBase = this._config.pasta_cards || '/Cards_Produtividade/Cards/';
    const ano       = this._extractYear(record);
    const cultura   = this._sanitize(record.culture || 'Geral');
    const pasta     = `${pastaBase.replace(/\/$/, '')}/${ano}/${cultura}`;

    await this.ensureFolderPath(pasta);

    const fullPath = `${pasta}/${nomeFile}`;
    return await this.uploadFile(pngDataUrl, fullPath, 'image/png');
  },

  /* ═══════════════════════════════════════════════════
     UTILITÁRIOS
  ═══════════════════════════════════════════════════ */

  /**
   * Gera o nome padronizado para o arquivo
   * Formato: nome_produtor_variedade_cidade_YYYY-MM-DD.ext
   */
  buildFilename(record, ext) {
    const sanitize = s => (s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // remove acentos
      .replace(/[^a-zA-Z0-9\s_-]/g, '')
      .trim()
      .replace(/\s+/g, '_');

    const date = this._isoDate(record);
    const name = [
      sanitize(record.producer_name),
      sanitize(record.variety_name),
      sanitize(record.city),
      date
    ].filter(Boolean).join('_');

    return `${name}.${ext}`;
  },

  /**
   * Extrai o ano do registro (da safra ou da colheita)
   */
  _extractYear(record) {
    // Tenta extrair do campo season (ex: "2025/26" → "2025")
    if (record.season) {
      const m = record.season.match(/(\d{4})/);
      if (m) return m[1];
    }
    // Tenta data de colheita
    if (record.harvest_date) {
      const m = record.harvest_date.match(/(\d{4})/);
      if (m) return m[1];
    }
    return new Date().getFullYear().toString();
  },

  _isoDate(record) {
    // Prefere harvest_date; fallback para data atual
    const raw = record.harvest_date || '';
    // Tenta formato DD/MM/YYYY
    const m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    // Tenta YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    return new Date().toISOString().slice(0, 10);
  },

  _sanitize(str) {
    return (str || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9\s]/g, '')
      .trim()
      .replace(/\s+/g, '_') || 'Geral';
  },

  _dataUrlToBlob(dataUrl) {
    const [header, data] = dataUrl.split(',');
    const mime = header.match(/:(.*?);/)[1];
    const binary = atob(data);
    const array  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) array[i] = binary.charCodeAt(i);
    return new Blob([array], { type: mime });
  },

  /* ── PKCE helpers ── */
  _generateCodeVerifier() {
    const array = new Uint8Array(32);
    crypto.getRandomValues(array);
    return btoa(String.fromCharCode(...array))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  },

  async _generateCodeChallenge(verifier) {
    const data    = new TextEncoder().encode(verifier);
    const digest  = await crypto.subtle.digest('SHA-256', data);
    return btoa(String.fromCharCode(...new Uint8Array(digest)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  },

  _generateState() {
    const array = new Uint8Array(16);
    crypto.getRandomValues(array);
    return Array.from(array).map(b => b.toString(16).padStart(2, '0')).join('');
  },

  /* ═══════════════════════════════════════════════════
     API: AUDIT LOGS
  ═══════════════════════════════════════════════════ */

  /**
   * Registra uma ação no log de auditoria
   */
  async log(action, record, extras = {}) {
    try {
      const user = typeof Auth !== 'undefined' ? Auth.currentUser : null;
      await fetch('tables/audit_logs', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          user_id:            user?.id    || '',
          user_name:          user?.name  || '',
          action,
          record_id:          record?.id          || '',
          card_filename:      extras.card_filename || '',
          card_onedrive_path: extras.card_path     || '',
          card_onedrive_id:   extras.card_id       || '',
          termo_filename:     extras.termo_filename || '',
          termo_onedrive_path:extras.termo_path     || '',
          termo_onedrive_id:  extras.termo_id       || '',
          producer_name:      record?.producer_name || '',
          variety_name:       record?.variety_name  || '',
          culture:            record?.culture        || '',
          city:               record?.city           || '',
          state:              record?.state          || '',
          productivity:       record?.productivity   || '',
          unit:               record?.unit           || '',
          season:             record?.season         || '',
          details:            JSON.stringify({ ...extras, timestamp: new Date().toISOString() })
        })
      });
    } catch (e) {
      console.warn('[OneDrive.log] Falha no log de auditoria:', e);
    }
  }
};
