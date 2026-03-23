/* =====================================================
   AgriCard Stine - Main App Controller
   ===================================================== */

const App = {

  // ===================================================
  // TOAST
  // ===================================================
  Toast: {
    _timer: null,
    show(msg, type = 'info') {
      const el = document.getElementById('toast');
      const icons = {
        success: 'fa-check-circle',
        error:   'fa-exclamation-circle',
        info:    'fa-info-circle',
        warning: 'fa-exclamation-triangle'
      };
      el.innerHTML = `<i class="fas ${icons[type] || 'fa-info-circle'}"></i> ${msg}`;
      el.className = `toast ${type}`;
      el.classList.remove('hidden');
      clearTimeout(this._timer);
      this._timer = setTimeout(() => el.classList.add('hidden'), 4000);
    }
  },

  // ===================================================
  // CONFIRM DIALOG
  // ===================================================
  confirm(message, onConfirm) {
    const modal   = document.getElementById('confirmModal');
    const msgEl   = document.getElementById('confirmMessage');
    const okBtn   = document.getElementById('btnConfirmOk');
    const cancelBtn = document.getElementById('btnConfirmCancel');

    msgEl.textContent = message;
    modal.classList.remove('hidden');

    const cleanup = () => modal.classList.add('hidden');

    okBtn.onclick = () => {
      cleanup();
      onConfirm();
    };
    cancelBtn.onclick = cleanup;
  },

  // ===================================================
  // SCREENS
  // ===================================================
  showScreen(id) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    const el = document.getElementById(id);
    if (el) el.classList.add('active');
    window.scrollTo(0, 0);
  },

  // ===================================================
  // NAVIGATE (lazy-load sections)
  // ===================================================
  navigate(sectionId) {
    const section = document.getElementById(sectionId);
    if (!section) return;

    if (sectionId.startsWith('admin-')) {
      document.querySelectorAll('#adminScreen .content-section').forEach(s => s.classList.remove('active'));
      document.querySelectorAll('#adminScreen .nav-item').forEach(n => n.classList.remove('active'));
      section.classList.add('active');
      const nav = document.querySelector(`#adminScreen [data-section="${sectionId}"]`);
      if (nav) nav.classList.add('active');
      this._loadAdminSection(sectionId);
    } else if (sectionId.startsWith('user-')) {
      document.querySelectorAll('#userScreen .content-section').forEach(s => s.classList.remove('active'));
      document.querySelectorAll('#userScreen .nav-item').forEach(n => n.classList.remove('active'));
      section.classList.add('active');
      const nav = document.querySelector(`#userScreen [data-section="${sectionId}"]`);
      if (nav) nav.classList.add('active');
      this._loadUserSection(sectionId);
    }

    // Close mobile sidebar
    this._closeMobileSidebar();
  },

  _loadAdminSection(id) {
    switch (id) {
      case 'admin-dashboard':     Admin.loadDashboard();    break;
      case 'admin-users':         Admin.loadUsers();         break;
      case 'admin-varieties':     Admin.loadVarieties();     break;
      case 'admin-records':       Admin.loadAllRecords();    break;
      case 'admin-cards-gallery': Admin.loadCardsGallery();  break;
      case 'admin-termos':        Admin.loadTermos();        break;
      case 'admin-audit':         Admin.loadAuditLogs();     break;
      case 'admin-pptx-studio':   /* static page, no init needed */ break;
      case 'admin-settings':      Settings.init();           break;
      case 'admin-templates':     TemplatesManager.init().then(() => {
        const countEl  = document.getElementById('tplCountText');
        const activeEl = document.getElementById('tplActiveText');
        if (countEl)  countEl.textContent  = `${TemplatesManager._templates.length} templates cadastrados`;
        if (activeEl) {
          const active = TemplatesManager._templates.find(t => t.active);
          activeEl.textContent = active ? `Ativo: ${active.name}` : 'Nenhum ativo';
        }
      }); break;
    }
  },

  _loadUserSection(id) {
    switch (id) {
      case 'user-dashboard':  User.loadDashboard();        break;
      case 'user-records':    User.loadMyRecords();         break;
      case 'user-new-record': User.loadVarietiesSelect();   break;
      case 'user-cards':      User.loadMyCards();           break;
    }
  },

  // ===================================================
  // AFTER LOGIN
  // ===================================================
  afterLogin(user) {
    if (user.role === 'admin') {
      document.getElementById('adminUserName').textContent = user.name;
      const av = document.getElementById('settingProfileAvatar');
      if (av) av.textContent = (user.name||'A').charAt(0).toUpperCase();
      this.showScreen('adminScreen');
      Admin.init();
      this._setupAdminNav();
    } else {
      document.getElementById('userUserName').textContent  = user.name;
      document.getElementById('userCompanyName').textContent = user.company || 'Produtor';
      this.showScreen('userScreen');
      User.init();
      this._setupUserNav();
    }

    // Show date on dashboard
    const dateEl = document.getElementById('dashboardDate');
    if (dateEl) {
      dateEl.textContent = new Date().toLocaleDateString('pt-BR', {
        weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
      });
    }

    // Aplica branding após login
    AppBranding.apply();
  },

  // ===================================================
  // NAV SETUP
  // ===================================================
  _setupAdminNav() {
    document.querySelectorAll('#adminScreen .nav-item').forEach(item => {
      item.addEventListener('click', e => {
        e.preventDefault();
        this.navigate(item.dataset.section);
      });
    });
    document.getElementById('btnAddVariety')?.addEventListener('click', () => {
      Admin.openVarietyModal(null);
    });

    // Cor da variedade live hex
    document.getElementById('varColor')?.addEventListener('input', (e) => {
      const hex = document.getElementById('varColorHex');
      if (hex) hex.textContent = e.target.value;
    });

    // Cor principal settings live hex
    document.getElementById('settingPrimaryColor')?.addEventListener('input', e => {
      const hex = document.getElementById('settingPrimaryColorHex');
      if (hex) hex.textContent = e.target.value;
    });

    // Logout buttons
    document.getElementById('btnAdminLogout')?.addEventListener('click', () => Auth.logout());
    document.getElementById('btnAdminLogoutMobile')?.addEventListener('click', () => Auth.logout());
  },

  _setupUserNav() {
    document.querySelectorAll('#userScreen .nav-item').forEach(item => {
      item.addEventListener('click', e => {
        e.preventDefault();
        this.navigate(item.dataset.section);
      });
    });
  },

  // ===================================================
  // MODALS
  // ===================================================
  closeVarietyModal() {
    Admin.closeVarietyModal();
  },

  // ===================================================
  // MOBILE SIDEBAR
  // ===================================================
  toggleMobileSidebar(panel) {
    const sidebar  = document.getElementById(panel === 'admin' ? 'adminSidebar'  : 'userSidebar');
    const overlay  = document.getElementById(panel === 'admin' ? 'adminSidebarOverlay' : 'userSidebarOverlay');
    const hamburger = document.getElementById(panel === 'admin' ? 'adminHamburger' : 'userHamburger');

    if (!sidebar) return;
    const isOpen = sidebar.classList.contains('open');

    if (isOpen) {
      sidebar.classList.remove('open');
      overlay?.classList.remove('open');
      hamburger?.classList.remove('open');
    } else {
      sidebar.classList.add('open');
      overlay?.classList.add('open');
      hamburger?.classList.add('open');
    }
  },

  _closeMobileSidebar() {
    ['adminSidebar','userSidebar'].forEach(id => {
      document.getElementById(id)?.classList.remove('open');
    });
    ['adminSidebarOverlay','userSidebarOverlay'].forEach(id => {
      document.getElementById(id)?.classList.remove('open');
    });
    ['adminHamburger','userHamburger'].forEach(id => {
      document.getElementById(id)?.classList.remove('open');
    });
  },

  // ===================================================
  // INIT
  // ===================================================
  init() {
    Auth.init();

    // Carrega branding antes do login
    AppBranding.load();

    // Auto-login if session exists
    if (Auth.currentUser) {
      this.afterLogin(Auth.currentUser);
    }

    // Setup file importer drag & drop
    if (typeof FileImporter !== 'undefined') {
      FileImporter.setupDrop();
    }

    // Close sidebar on resize to desktop
    window.addEventListener('resize', () => {
      if (window.innerWidth > 900) {
        this._closeMobileSidebar();
      }
    });

    // Keyboard ESC → close modals
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        document.querySelectorAll('.modal:not(.hidden)').forEach(m => {
          m.classList.add('hidden');
          document.body.style.overflow = '';
        });
      }
    });
  }
};

// ===== BOOT =====
document.addEventListener('DOMContentLoaded', () => App.init());
