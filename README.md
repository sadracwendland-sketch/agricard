# AgriCard Stine — Plataforma de Cards de Produtividade Agrícola

**Versão:** v9.0  
**Data:** 2026-03-22

---

## 🌱 Sobre o Projeto

Plataforma web para geração de **cards profissionais de produtividade agrícola** a partir de templates PPTX personalizados. Utilizada por representantes e consultores da STINE Sementes para registrar e divulgar resultados de campo.

---

## ✅ Funcionalidades Implementadas

### 🔐 Autenticação e Controle de Acesso
- Login/logout com sessão persistente (localStorage)
- Cadastro de novos usuários com status "pendente" (aguarda aprovação do admin)
- **Dois perfis de acesso:**
  - **Administrador:** acesso total — editar/excluir qualquer card, exportar dados, acessar OneDrive, ver termos, gerenciar usuários/variedades/configurações, visualizar auditoria
  - **Usuário comum:** criar registros, visualizar e baixar apenas seus próprios cards; **sem** edição, exclusão, OneDrive ou acesso a termos de terceiros
- Checkbox LGPD obrigatório no cadastro (registra timestamp de aceite)
- Checkbox de consentimento LGPD obrigatório em cada novo registro (confirmação do produtor)

### 📋 Registros de Produtividade
- Formulário completo: variedade, tecnologia, cultura, safra, datas, produtividade, área, localização, produtor, fazenda
- Validação de campos obrigatórios
- **Upload obrigatório do Termo de Autorização** (PDF/JPG/PNG, máx. 10 MB) com drag & drop
  - Nome padronizado automático: `nome_produtor_variedade_cidade_YYYY-MM-DD.ext`
  - Aceita apenas novos registros (edição não exige novo upload)
- Separação automática da produtividade em `int` e `dec` (aceita vírgula e ponto)
- Salvar como rascunho ou gerar card imediatamente

### 🃏 Geração de Cards (PPTX Pixel-Perfect)
- Parser PPTX detecta automaticamente o `slideLayout` correto (ex: `slideLayout5`)
- Extrai: imagem de fundo (slideMaster), logo da variedade (`{{logo_variedade}}`), posições/fontes/cores dos placeholders
- Renderização fiel com `CardRenderer.render()` (EMU → px, escala do canvas)
- **Geração em alta resolução:** PNG 3× (até 2160px)
- Modo legado com coordenadas calibradas manualmente (para variedades sem PPTX)
- **Fluxo pós-geração (usuário comum):**
  1. Download automático do PNG
  2. Upload para OneDrive (se configurado)
  3. Tela de sucesso com preview e botão de download
  4. Redirecionamento para "Meus Cards"
  5. Botão "Editar" oculto para usuários comuns

### ☁️ Integração OneDrive (Microsoft Graph API)
- Autenticação OAuth2 PKCE (sem senha armazenada)
- Criação automática de pastas hierárquicas:
  - `/Termos/{ano}/{cultura}/`
  - `/Cards/{ano}/{cultura}/`
- Upload automático de termos após registro
- Upload automático de cards PNG após geração
- Configuração via Admin → Configurações → OneDrive (Client ID, Tenant ID, pastas base)
- Testar conexão diretamente no painel
- Armazena `card_onedrive_path`, `card_onedrive_id`, `termo_onedrive_path` no registro

### 👨‍💼 Painel Administrador
- **Dashboard:** métricas (usuários, pendentes, registros, variedades), alertas de aprovação
- **Usuários:** aprovar/rejeitar/bloquear/excluir; filtros por status
- **Variedades:** CRUD + upload PPTX com extração automática de placeholders, compressão de imagens
- **Registros:** listagem completa com colunas de usuário e termo; busca; exportação CSV
- **Termos Enviados:** visualizar arquivo (PDF abre em nova aba, imagem em janela) + exportar CSV
- **Auditoria:** log de ações (criação, geração de card, download, upload de termo) + exportar CSV
- **Galeria de Cards:** todos os cards publicados
- **Configurações:** perfil, senha, branding, **OneDrive** (Client ID, Tenant ID, pastas), admins, zona de perigo

### 📊 Exportação de Dados
- **CSV de Registros:** usuário, produtor, variedade, localização, produtividade, status, termo, OneDrive
- **CSV de Termos:** produtor, variedade, arquivo, caminho OneDrive
- **CSV de Auditoria:** data/hora, usuário, ação, produtor, variedade, paths de card e termo
- BOM UTF-8 para compatibilidade com Excel

### 🔍 Auditoria / Logging
- Registra automaticamente: criação de registro, geração de card, download, upload de termo
- Campos: user_id, user_name, action, record_id, card_filename, card_onedrive_path, termo_filename, termo_onedrive_path, producer_name, variety_name, culture, city, state, productivity, unit, season, details (JSON), timestamps

### 📱 UX / Interface
- Responsivo para desktop e mobile
- Sidebar colapsável no mobile
- Toast notifications para feedback
- Modais com scroll e foco
- PPTX Card Studio para criação manual de cards

---

## 🗂️ Estrutura de Arquivos

```
index.html              — Aplicação principal (SPA)
auth-callback.html      — Página de retorno OAuth2 OneDrive
css/
  style.css             — Estilos completos
js/
  api.js                — Helper RESTful API
  auth.js               — Autenticação, login, cadastro, LGPD
  access-control.js     — Controle de acesso por perfil (PERMISSIONS)
  onedrive.js           — Integração OneDrive OAuth2 PKCE
  card.js               — CardRenderer + CardGenerator (geração de cards)
  card-renderer.js      — SlideRenderer para PPTX Studio
  admin.js              — Painel admin (usuários, variedades, registros, termos, auditoria, CSV)
  user.js               — Painel usuário (registros, formulário, galeria, termo upload)
  settings.js           — Configurações admin + OneDrive + AppBranding
  templates.js          — Gerenciador de templates de card (legado)
  fileimport.js         — Importação de PPTX/PDF como template
  pptx-engine.js        — Parser PPTX (PptxParser v8)
  pptx-studio.js        — PPTX Card Studio interativo
  app.js                — Controlador principal (App, Toast, navegação, roteamento)
uploads/
  card_template_v2.pptx — Template PPTX da STINE (exemplo)
```

---

## 🗄️ Modelo de Dados (Tabelas)

| Tabela | Campos principais |
|--------|-------------------|
| `users` | id, name, email, password, company, role (admin/user), status (pending/approved/rejected), phone, region, lgpd_accepted_at |
| `productivity_records` | id, user_id, user_name, variety_id, variety_name, brand, technology, culture, season, planting_date, harvest_date, productivity, unit, city, state, producer_name, farm_name, area, status, notes, plant_population, **termo_file** (base64), **termo_filename**, **termo_nome_padronizado**, **termo_onedrive_path**, **termo_onedrive_id**, **card_filename**, **card_onedrive_path**, **card_onedrive_id**, lgpd_accepted |
| `varieties` | id, name, brand, technology, culture, maturity_group, primary_color, logo_url, description, template_image (JPEG base64), field_coords (JSON), pptx_elements (JSON), pptx_slide_w, pptx_slide_h, pptx_logo (base64) |
| `audit_logs` | id, user_id, user_name, action, record_id, card_filename, card_onedrive_path, card_onedrive_id, termo_filename, termo_onedrive_path, termo_onedrive_id, producer_name, variety_name, culture, city, state, productivity, unit, season, details |
| `onedrive_config` | id, key, value, label, category, updated_by |
| `app_settings` | id (chave), value, label, category, updated_by |
| `card_templates` | id, name, description, layout_type, bg_image_url, header_color, header_text, show_ranking_badge, slogan, footer_logo_text, badge_label, active, sort_order, created_by |

---

## 🔑 Permissões por Perfil

| Permissão | Admin | Usuário |
|-----------|-------|---------|
| Criar card | ✅ | ✅ |
| Ver próprios cards | ✅ | ✅ |
| Ver todos os cards | ✅ | ❌ |
| Editar card | ✅ | ❌ |
| Excluir card | ✅ | ❌ |
| Baixar card | ✅ | ✅ |
| Exportar dados CSV | ✅ | ❌ |
| Acessar OneDrive | ✅ | ❌ |
| Ver termos | ✅ | ❌ |
| Gerenciar usuários | ✅ | ❌ |
| Gerenciar variedades | ✅ | ❌ |
| Gerenciar configurações | ✅ | ❌ |
| Ver logs de auditoria | ✅ | ❌ |

---

## 🚀 Configuração OneDrive

1. Acesse [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App Registrations → Novo registro
2. Nome: `AgriCard STINE`, Conta: "Multitenant" ou conta pessoal
3. URI de Redirecionamento: `https://seu-dominio/auth-callback.html` (tipo: SPA)
4. Copie o **Application (Client) ID** e **Directory (Tenant) ID**
5. No AgriCard: Admin → Configurações → OneDrive → cole os IDs e configure as pastas base
6. Clique em **Testar Conexão** para autenticar e verificar

---

## 🛤️ Rotas / Seções

| URL/Seção | Descrição |
|-----------|-----------|
| `/` (authScreen) | Tela de login/cadastro |
| `adminScreen > admin-dashboard` | Dashboard admin |
| `adminScreen > admin-users` | Gerenciar usuários |
| `adminScreen > admin-varieties` | Gerenciar variedades + PPTX |
| `adminScreen > admin-records` | Todos os registros |
| `adminScreen > admin-termos` | Termos de autorização |
| `adminScreen > admin-audit` | Log de auditoria |
| `adminScreen > admin-cards-gallery` | Galeria de cards publicados |
| `adminScreen > admin-settings` | Configurações (perfil, branding, OneDrive, admins) |
| `userScreen > user-dashboard` | Painel do usuário |
| `userScreen > user-new-record` | Novo registro + upload termo |
| `userScreen > user-records` | Meus registros |
| `userScreen > user-cards` | Meus cards |

---

## 📋 Pendências / Próximos Passos

- [ ] Adicionar paginação na lista de registros e auditoria (>500 registros)
- [ ] Notificação por e-mail ao admin quando novo usuário se cadastra
- [ ] Preview do termo (PDF inline) dentro do modal
- [ ] Filtros avançados por data/cultura/safra nos registros
- [ ] Relatório comparativo de produtividade por variedade/região
- [ ] Impressão/PDF do card direto pelo navegador
- [ ] Modo offline (service worker) para geração sem internet
