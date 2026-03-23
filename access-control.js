/* =====================================================
   AgriCard Stine – Access Control Module v1.0
   
   Define permissões por perfil e aplica restrições
   de UI e ações conforme o papel do usuário logado.
   
   Perfis:
   - admin : acesso total
   - user  : somente criação e visualização de seus cards
   ===================================================== */

'use strict';

const AccessControl = {

  /* ─────────────────────────────────────────────────
     DEFINIÇÃO DE PERMISSÕES
  ───────────────────────────────────────────────── */
  PERMISSIONS: {
    admin: {
      can_create_card:        true,
      can_view_own_cards:     true,
      can_view_all_cards:     true,
      can_edit_card:          true,
      can_delete_card:        true,
      can_download_card:      true,
      can_export_data:        true,
      can_access_onedrive:    true,
      can_view_termos:        true,
      can_manage_users:       true,
      can_manage_varieties:   true,
      can_manage_settings:    true,
      can_view_audit_logs:    true,
    },
    user: {
      can_create_card:        true,
      can_view_own_cards:     true,
      can_view_all_cards:     false,
      can_edit_card:          false,    // usuário NÃO pode editar
      can_delete_card:        false,    // usuário NÃO pode excluir
      can_download_card:      true,
      can_export_data:        false,
      can_access_onedrive:    false,
      can_view_termos:        false,
      can_manage_users:       false,
      can_manage_varieties:   false,
      can_manage_settings:    false,
      can_view_audit_logs:    false,
    }
  },

  /* ─────────────────────────────────────────────────
     VERIFICAÇÃO DE PERMISSÃO
  ───────────────────────────────────────────────── */

  /**
   * Retorna o perfil do usuário atual
   */
  getRole() {
    return Auth?.currentUser?.role || 'user';
  },

  /**
   * Verifica se o usuário tem a permissão solicitada
   * @param {string} permission - chave da permissão
   * @returns {boolean}
   */
  can(permission) {
    const role  = this.getRole();
    const perms = this.PERMISSIONS[role] || this.PERMISSIONS.user;
    return perms[permission] === true;
  },

  /**
   * Verifica se o usuário é admin
   */
  isAdmin() {
    return this.getRole() === 'admin';
  },

  /**
   * Garante que o usuário tem permissão; lança erro se não tiver
   * @param {string} permission
   * @param {string} [errorMsg] - mensagem de erro personalizada
   */
  require(permission, errorMsg) {
    if (!this.can(permission)) {
      const msg = errorMsg || `Você não tem permissão para esta ação.`;
      App.Toast.show(msg, 'error');
      throw new Error(msg);
    }
  },

  /* ─────────────────────────────────────────────────
     APLICAÇÃO DE RESTRIÇÕES NA UI
  ───────────────────────────────────────────────── */

  /**
   * Aplica restrições visuais com base no perfil.
   * Deve ser chamado após o login.
   */
  applyUI() {
    const role = this.getRole();

    if (role === 'user') {
      this._applyUserRestrictions();
    }
    // Admin: nenhuma restrição necessária (acesso total)
  },

  /**
   * Restrições aplicadas ao perfil "usuário comum"
   */
  _applyUserRestrictions() {
    // Oculta botões de editar e excluir em todos os contextos
    this._hideElements([
      '.btn-edit-record',
      '.btn-delete-record',
      '[data-admin-only]',
    ]);

    // Desabilita campos de edição na galeria do usuário
    this._disableElements([
      '#btnSaveDraft',   // Rascunho (somente após 1ª criação)
    ]);
  },

  /**
   * Verifica permissão e retorna HTML dos botões de ação
   * para a tabela de registros do USUÁRIO
   */
  renderUserRecordActions(recordId) {
    // Usuário comum: só pode ver o card (download)
    // Admin: pode editar e excluir
    if (this.isAdmin()) {
      return `
        <div class="table-actions">
          <button class="action-btn action-btn-blue" onclick="User.previewRecord('${recordId}')">
            <i class="fas fa-image"></i> Card
          </button>
          <button class="action-btn action-btn-orange btn-edit-record" onclick="User.editRecord('${recordId}')">
            <i class="fas fa-edit"></i>
          </button>
          <button class="action-btn action-btn-red btn-delete-record" onclick="User.deleteRecord('${recordId}')">
            <i class="fas fa-trash"></i>
          </button>
        </div>`;
    } else {
      // Usuário comum: apenas visualização e download
      return `
        <div class="table-actions">
          <button class="action-btn action-btn-blue" onclick="User.previewRecord('${recordId}')">
            <i class="fas fa-image"></i> Ver Card
          </button>
        </div>`;
    }
  },

  /**
   * Verifica permissão e retorna HTML dos botões de ação
   * para a tabela de registros do ADMIN
   */
  renderAdminRecordActions(recordId) {
    // Admin sempre tem acesso total
    return `
      <div class="table-actions">
        <button class="action-btn action-btn-blue" onclick="CardGenerator.openPreviewById('${recordId}')">
          <i class="fas fa-image"></i> Card
        </button>
        <button class="action-btn action-btn-orange" onclick="Admin.editRecord('${recordId}')">
          <i class="fas fa-edit"></i>
        </button>
        <button class="action-btn action-btn-red" onclick="Admin.deleteRecord('${recordId}')">
          <i class="fas fa-trash"></i>
        </button>
        <button class="action-btn" title="Ver Termo" onclick="Admin.viewTermo('${recordId}')">
          <i class="fas fa-file-alt"></i>
        </button>
      </div>`;
  },

  /* ─────────────────────────────────────────────────
     VALIDAÇÃO DO TERMO DE AUTORIZAÇÃO
  ───────────────────────────────────────────────── */

  /**
   * Valida um arquivo de termo de autorização
   * @param {File} file
   * @returns {{ valid: boolean, error?: string }}
   */
  validateTermo(file) {
    if (!file) {
      return { valid: false, error: 'O termo de autorização é obrigatório.' };
    }

    const MAX_SIZE = 10 * 1024 * 1024; // 10MB
    if (file.size > MAX_SIZE) {
      return { valid: false, error: `Arquivo muito grande (máx. 10MB). Tamanho atual: ${(file.size / 1024 / 1024).toFixed(1)}MB.` };
    }

    const ALLOWED = ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png'];
    const ALLOWED_EXT = ['.pdf', '.jpg', '.jpeg', '.png'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();

    if (!ALLOWED.includes(file.type) && !ALLOWED_EXT.includes(ext)) {
      return { valid: false, error: 'Formato inválido. Use PDF, JPG ou PNG.' };
    }

    return { valid: true };
  },

  /**
   * Lê o arquivo como base64
   */
  readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result);
      reader.onerror = () => reject(new Error('Falha ao ler arquivo'));
      reader.readAsDataURL(file);
    });
  },

  /* ─────────────────────────────────────────────────
     HELPERS DOM
  ───────────────────────────────────────────────── */

  _hideElements(selectors) {
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        el.style.display = 'none';
      });
    });
  },

  _disableElements(selectors) {
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        el.disabled = true;
        el.title    = 'Não disponível para seu perfil.';
      });
    });
  },

  /**
   * Retorna badge HTML indicando o perfil do usuário
   */
  roleBadge(role) {
    if (role === 'admin') {
      return '<span class="badge" style="background:#7c3aed;color:#fff"><i class="fas fa-shield-alt"></i> Admin</span>';
    }
    return '<span class="badge badge-draft"><i class="fas fa-user"></i> Usuário</span>';
  }
};
