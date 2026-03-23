/* =====================================================
   AgriCard Stine - Auth Module
   ===================================================== */

const Auth = {
  currentUser: null,

  init() {
    // Restore session
    try {
      const saved = localStorage.getItem('agricard_user');
      if (saved) this.currentUser = JSON.parse(saved);
    } catch {
      localStorage.removeItem('agricard_user');
    }

    // Tab switching
    document.querySelectorAll('.auth-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.auth-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.auth-form').forEach(f => f.classList.remove('active'));
        tab.classList.add('active');
        const target = tab.dataset.tab === 'login' ? 'loginForm' : 'registerForm';
        document.getElementById(target).classList.add('active');
        document.getElementById('authMessage').classList.add('hidden');
      });
    });

    // Logout buttons (desktop + mobile)
    document.getElementById('btnAdminLogout')?.addEventListener('click', () => this.logout());
    document.getElementById('btnUserLogout')?.addEventListener('click',  () => this.logout());
    document.getElementById('btnAdminLogoutMobile')?.addEventListener('click', () => this.logout());
    document.getElementById('btnUserLogoutMobile')?.addEventListener('click',  () => this.logout());

    // Color picker preview in variety modal
    const varColor = document.getElementById('varColor');
    if (varColor) {
      varColor.addEventListener('input', e => {
        const hex = document.getElementById('varColorHex');
        if (hex) hex.textContent = e.target.value;
      });
    }
  },

  async login() {
    const email    = document.getElementById('loginEmail').value.trim().toLowerCase();
    const password = document.getElementById('loginPassword').value;

    if (!email || !password) {
      return this.showMessage('Preencha e-mail e senha.', 'error');
    }

    const btn = document.getElementById('btnLogin');
    btn.innerHTML = '<span class="loading"></span> Entrando...';
    btn.disabled = true;

    try {
      const res   = await API.getUsers(`search=${encodeURIComponent(email)}`);
      const users = res.data || [];
      const user  = users.find(u =>
        u.email && u.email.toLowerCase() === email && u.password === password
      );

      if (!user) {
        return this.showMessage('E-mail ou senha incorretos.', 'error');
      }
      if (user.status === 'pending') {
        return this.showMessage('Sua conta ainda aguarda aprovação do administrador.', 'error');
      }
      if (user.status === 'rejected') {
        return this.showMessage('Solicitação rejeitada. Entre em contato com o administrador.', 'error');
      }

      this.currentUser = user;
      localStorage.setItem('agricard_user', JSON.stringify(user));
      App.afterLogin(user);
      // Aplica controle de acesso conforme perfil
      if (typeof AccessControl !== 'undefined') {
        setTimeout(() => AccessControl.applyUI(), 100);
      }

    } catch (err) {
      this.showMessage('Erro ao conectar. Tente novamente.', 'error');
      console.error(err);
    } finally {
      btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Entrar';
      btn.disabled = false;
    }
  },

  async register() {
    const name    = document.getElementById('regName').value.trim();
    const email   = document.getElementById('regEmail').value.trim().toLowerCase();
    const company = document.getElementById('regCompany').value.trim();
    const phone   = document.getElementById('regPhone').value.trim();
    const region  = document.getElementById('regRegion').value.trim();
    const pass    = document.getElementById('regPassword').value;
    const confirm = document.getElementById('regPasswordConfirm').value;
    const lgpd    = document.getElementById('regLgpd')?.checked;

    if (!name || !email || !pass) {
      return this.showMessage('Preencha todos os campos obrigatórios.', 'error');
    }
    if (pass.length < 6) {
      return this.showMessage('A senha deve ter pelo menos 6 caracteres.', 'error');
    }
    if (pass !== confirm) {
      return this.showMessage('As senhas não conferem.', 'error');
    }
    if (!lgpd) {
      return this.showMessage('Você precisa aceitar os termos LGPD para continuar.', 'error');
    }

    const btn = document.getElementById('btnRegister');
    btn.innerHTML = '<span class="loading"></span> Enviando...';
    btn.disabled = true;

    try {
      // Check duplicate
      const res = await API.getUsers(`search=${encodeURIComponent(email)}`);
      const existing = (res.data || []).find(u =>
        u.email && u.email.toLowerCase() === email
      );
      if (existing) {
        return this.showMessage('Este e-mail já está cadastrado.', 'error');
      }

      await API.createUser({
        name, email, password: pass,
        company: company || '',
        phone:   phone   || '',
        region:  region  || '',
        role:   'user',
        status: 'pending',
        lgpd_accepted_at: new Date().toISOString()
      });

      this.showMessage('Solicitação enviada! Aguarde a aprovação do administrador.', 'success');
      ['regName','regEmail','regCompany','regPhone','regRegion','regPassword','regPasswordConfirm']
        .forEach(id => { document.getElementById(id).value = ''; });

    } catch (err) {
      this.showMessage('Erro ao cadastrar. Tente novamente.', 'error');
      console.error(err);
    } finally {
      btn.innerHTML = '<i class="fas fa-user-plus"></i> Solicitar Acesso';
      btn.disabled = false;
    }
  },

  logout() {
    this.currentUser = null;
    localStorage.removeItem('agricard_user');
    App.showScreen('authScreen');
    App.Toast.show('Você saiu do sistema.', 'info');
  },

  togglePassword(inputId, btn) {
    const input = document.getElementById(inputId);
    const icon  = btn.querySelector('i');
    if (input.type === 'password') {
      input.type = 'text';
      icon.className = 'fas fa-eye-slash';
    } else {
      input.type = 'password';
      icon.className = 'fas fa-eye';
    }
  },

  showMessage(msg, type) {
    const el = document.getElementById('authMessage');
    el.textContent = msg;
    el.className = `auth-message ${type}`;
    el.classList.remove('hidden');
    clearTimeout(this._msgTimer);
    this._msgTimer = setTimeout(() => el.classList.add('hidden'), 6000);
  }
};
