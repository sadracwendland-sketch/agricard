/* =====================================================
   AgriCard Stine - API Helper (RESTful Table API)
   ===================================================== */

const API = {
  baseUrl: 'tables',

  async request(method, endpoint, body = null) {
    const opts = {
      method,
      headers: { 'Content-Type': 'application/json' }
    };
    if (body) opts.body = JSON.stringify(body);
    const res = await fetch(`${this.baseUrl}/${endpoint}`, opts);
    if (res.status === 204) return null;
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.message || `HTTP ${res.status}`);
    }
    return await res.json();
  },

  // ====== USERS ======
  async getUsers(params = '') {
    return this.request('GET', `users?limit=500${params ? '&' + params : ''}`);
  },
  async getUser(id) {
    return this.request('GET', `users/${id}`);
  },
  async createUser(data) {
    return this.request('POST', 'users', data);
  },
  async updateUser(id, data) {
    return this.request('PATCH', `users/${id}`, data);
  },
  async deleteUser(id) {
    return this.request('DELETE', `users/${id}`);
  },

  // ====== VARIETIES ======
  async getVarieties(params = '') {
    return this.request('GET', `varieties?limit=500${params ? '&' + params : ''}`);
  },
  async getVariety(id) {
    return this.request('GET', `varieties/${id}`);
  },
  async createVariety(data) {
    return this.request('POST', 'varieties', data);
  },
  async updateVariety(id, data) {
    return this.request('PATCH', `varieties/${id}`, data);
  },
  async deleteVariety(id) {
    return this.request('DELETE', `varieties/${id}`);
  },

  // ====== CARD TEMPLATES ======
  async getTemplates(params = '') {
    return this.request('GET', `card_templates?limit=100${params ? '&' + params : ''}`);
  },
  async getTemplate(id) {
    return this.request('GET', `card_templates/${id}`);
  },
  async createTemplate(data) {
    return this.request('POST', 'card_templates', data);
  },
  async updateTemplate(id, data) {
    return this.request('PATCH', `card_templates/${id}`, data);
  },
  async deleteTemplate(id) {
    return this.request('DELETE', `card_templates/${id}`);
  },

  // ====== PRODUCTIVITY RECORDS ======
  async getRecords(params = '') {
    return this.request('GET', `productivity_records?limit=500${params ? '&' + params : ''}`);
  },
  async getRecord(id) {
    return this.request('GET', `productivity_records/${id}`);
  },
  async createRecord(data) {
    return this.request('POST', 'productivity_records', data);
  },
  async updateRecord(id, data) {
    return this.request('PATCH', `productivity_records/${id}`, data);
  },
  async deleteRecord(id) {
    return this.request('DELETE', `productivity_records/${id}`);
  }
};
