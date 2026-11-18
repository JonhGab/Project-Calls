const STORAGE_KEY = 'chamados_express_spa_v1';
const SESSION_KEY = 'chamados_express_spa_session';
const SHEETS_URL_KEY = 'chamados_express_sheets_url';

// Se quiser deixar fixo para toda a equipe sem precisar colar, coloque sua URL /exec aqui:
const DEFAULT_SHEETS_URL = 'https://script.google.com/macros/s/AKfycbzOi2e7I6ka6twkLguSAmj-PAqoEoUkDrdmQfw-9wO-m7RJ-D21kZDp9GtjTJduhp15AA/exec';

const DEFAULT_SETTINGS = {
  portalUrl: 'https://suporte.empresa.com.br/chamados',
  senderEmail: '',
  emailProvider: 'outlook_office',
  technicians: ['João Silva', 'Ana Paula', 'Carlos Eduardo'],
  reminderTemplate:
    'Olá {nome}! 👋\n\nRealizamos um atendimento de TI para você em {data} (Técnico responsável: {tecnico}).\n\n📌 Problema resolvido:\n{problema}\n\nPor favor, lembre-se de registrar o chamado formal no portal de TI para controle da equipe:\n🔗 {link_portal}\n\nMuito obrigado!',
};

function createSeedStore() {
  const now = Date.now();
  const twoHoursAgo = new Date(now - 2 * 60 * 60 * 1000).toISOString();
  const fourDaysAgo = new Date(now - 4 * 24 * 60 * 60 * 1000).toISOString();

  return {
    settings: { ...DEFAULT_SETTINGS },
    users: [
      {
        id: 'tech-joao',
        name: 'João Silva',
        email: 'joao.silva@empresa.com.br',
        password: '123456',
        createdAt: fourDaysAgo,
      },
      {
        id: 'tech-ana',
        name: 'Ana Paula',
        email: 'ana.paula@empresa.com.br',
        password: '123456',
        createdAt: fourDaysAgo,
      },
    ],
    records: [
      {
        id: crypto.randomUUID(),
        userName: 'Mariana Costa',
        userEmail: 'mariana.costa@empresa.com.br',
        department: 'Financeiro',
        technician: 'João Silva',
        createdById: 'tech-joao',
        createdByName: 'João Silva',
        category: 'Rede / VPN',
        problemDescription:
          'Usuário não conseguia conectar na VPN corporativa após troca de senha do Windows. Reconfigurado perfil FortiClient e validado acesso ao ERP.',
        status: 'pending',
        officialTicketId: null,
        attendedAt: twoHoursAgo,
        createdAt: twoHoursAgo,
        updatedAt: twoHoursAgo,
        reminders: [],
      },
      {
        id: crypto.randomUUID(),
        userName: 'Rafael Mendes',
        userEmail: 'rafael.mendes@empresa.com.br',
        department: 'Comercial',
        technician: 'Ana Paula',
        createdById: 'tech-ana',
        createdByName: 'Ana Paula',
        category: 'Impressora / Periféricos',
        problemDescription:
          'Impressora do 3º andar estava offline na fila de impressão. Reiniciado spooler e reinstalado driver padrão.',
        status: 'reminded',
        officialTicketId: null,
        attendedAt: fourDaysAgo,
        createdAt: fourDaysAgo,
        updatedAt: fourDaysAgo,
        reminders: [
          {
            id: crypto.randomUUID(),
            channel: 'teams_deeplink',
            sentAt: fourDaysAgo,
            sentBy: 'Ana Paula',
            status: 'opened',
            details: 'Chat direto do Teams aberto para rafael.mendes@empresa.com.br',
          },
        ],
      },
    ],
    auditLog: [
      {
        id: crypto.randomUUID(),
        timestamp: fourDaysAgo,
        actorId: 'tech-ana',
        actorName: 'Ana Paula',
        action: 'create',
        summary: 'Registrou atendimento para Rafael Mendes (Impressora / Periféricos)',
      },
      {
        id: crypto.randomUUID(),
        timestamp: twoHoursAgo,
        actorId: 'tech-joao',
        actorName: 'João Silva',
        action: 'create',
        summary: 'Registrou atendimento para Mariana Costa (Rede / VPN)',
      },
    ],
  };
}

function loadStore() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    const seed = createSeedStore();
    saveStore(seed);
    return seed;
  }
  try {
    const parsed = JSON.parse(raw);
    return {
      settings: { ...DEFAULT_SETTINGS, ...(parsed.settings || {}) },
      users: Array.isArray(parsed.users) ? parsed.users : [],
      records: Array.isArray(parsed.records) ? parsed.records : [],
      auditLog: Array.isArray(parsed.auditLog) ? parsed.auditLog : [],
    };
  } catch {
    const seed = createSeedStore();
    saveStore(seed);
    return seed;
  }
}

function saveStore(store) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
}

const state = {
  sheetsUrl: localStorage.getItem(SHEETS_URL_KEY) || DEFAULT_SHEETS_URL,
  store: loadStore(),
  user: null,
  statusFilter: 'all',
  techFilter: '',
  search: '',
  overdueAlertShown: false,
};

const el = (id) => document.getElementById(id);

function showToast(message, type = 'info') {
  const container = el('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4500);
}

// Google Sheets Apps Script API helpers
async function sheetsGet() {
  if (!state.sheetsUrl) return null;
  const res = await fetch(state.sheetsUrl, { method: 'GET' });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || 'Erro ao ler Google Sheets.');
  return data;
}

async function sheetsPost(payload) {
  if (!state.sheetsUrl) return null;
  const res = await fetch(state.sheetsUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || 'Erro ao gravar no Google Sheets.');
  return data;
}

async function syncFromSheets(quiet = false) {
  if (!state.sheetsUrl) return;
  try {
    const remote = await sheetsGet();
    if (!remote) return;
    state.store = {
      settings: { ...DEFAULT_SETTINGS, ...(remote.settings || {}) },
      users: Array.isArray(remote.users) ? remote.users : [],
      records: Array.isArray(remote.records) ? remote.records : [],
      auditLog: Array.isArray(remote.auditLog) ? remote.auditLog : [],
    };
    saveStore(state.store);
    if (state.user) renderAll();
    if (!quiet) showToast('Sincronizado com o Google Sheets!', 'success');
  } catch (err) {
    showToast(`Falha na sincronização com Google Sheets: ${err.message}`, 'error');
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate(iso) {
  if (!iso) return '—';
  return new Intl.DateTimeFormat('pt-BR', {
    dateStyle: 'short',
    timeStyle: 'short',
  }).format(new Date(iso));
}

function getDaysOpen(record) {
  const ref = new Date(record.attendedAt || record.createdAt).getTime();
  return Math.floor((Date.now() - ref) / (24 * 60 * 60 * 1000));
}

function isOverdue(record) {
  return record.status !== 'resolved' && getDaysOpen(record) >= 3;
}

function setDefaultAttendedAt() {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const local = new Date(now.getTime() - offset * 60000);
  el('attendedAt').value = local.toISOString().slice(0, 16);
}

function addAudit(action, summary) {
  state.store.auditLog.unshift({
    id: crypto.randomUUID(),
    timestamp: new Date().toISOString(),
    actorId: state.user?.id || null,
    actorName: state.user?.name || 'Sistema',
    action,
    summary,
  });
  if (state.store.auditLog.length > 500) {
    state.store.auditLog = state.store.auditLog.slice(0, 500);
  }
}

function buildFormattedMessage(record) {
  const template = state.store.settings?.reminderTemplate || '';
  const portalUrl = state.store.settings?.portalUrl || '';
  return template
    .replace(/\{nome\}/g, record.userName || '')
    .replace(/\{email\}/g, record.userEmail || '')
    .replace(/\{tecnico\}/g, record.technician || '')
    .replace(/\{problema\}/g, record.problemDescription || '')
    .replace(/\{categoria\}/g, record.category || '')
    .replace(/\{data\}/g, formatDate(record.attendedAt || record.createdAt))
    .replace(/\{link_portal\}/g, portalUrl);
}

function buildTeamsDeepLink(record, useDesktopProtocol = false) {
  const message = buildFormattedMessage(record);
  const query = `users=${encodeURIComponent(record.userEmail)}&message=${encodeURIComponent(
    message
  )}`;
  return useDesktopProtocol
    ? `msteams:/l/chat/0/0?${query}`
    : `https://teams.microsoft.com/l/chat/0/0?${query}`;
}

function buildWebmailLink(record, customSubject, customBody) {
  const subject = customSubject || `Lembrete: Abrir chamado de TI (${record.category || 'Atendimento'})`;
  const body = customBody || buildFormattedMessage(record);
  const provider = state.store.settings?.emailProvider || 'outlook_office';

  if (provider === 'gmail') {
    const params = new URLSearchParams({
      view: 'cm',
      fs: '1',
      to: record.userEmail,
      su: subject,
      body,
    });
    return `https://mail.google.com/mail/?${params.toString()}`;
  }

  const base =
    provider === 'outlook_live'
      ? 'https://outlook.live.com/mail/0/deeplink/compose'
      : 'https://outlook.office.com/mail/deeplink/compose';

  const params = new URLSearchParams({
    to: record.userEmail,
    subject,
    body,
  });
  return `${base}?${params.toString()}`;
}

function triggerProtocol(url) {
  const a = document.createElement('a');
  a.href = url;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function showAuthView() {
  el('authView').classList.remove('hidden');
  el('appView').classList.add('hidden');
  el('authSheetsUrl').value = state.sheetsUrl || '';
}

function showAppView() {
  el('authView').classList.add('hidden');
  el('appView').classList.remove('hidden');
  el('currentUserChip').textContent = `👤 ${state.user.name}`;
  if (!el('technician').value) {
    el('technician').value = state.user.name;
  }
  renderAll();
  checkOverdueOnStartup();
}

function logout(notify = true) {
  state.user = null;
  state.overdueAlertShown = false;
  localStorage.removeItem(SESSION_KEY);
  showAuthView();
  if (notify) showToast('Você saiu do sistema.', 'info');
}

function checkOverdueOnStartup() {
  const overdue = state.store.records.filter(isOverdue);
  if (overdue.length > 0 && !state.overdueAlertShown) {
    state.overdueAlertShown = true;
    openOverdueModal();
  }
}

function renderDbModePill() {
  const pill = el('dbModePill');
  const syncBtn = el('syncSheetsBtn');
  if (state.sheetsUrl) {
    pill.textContent = 'Google Sheets: Conectado';
    pill.classList.add('active');
    syncBtn.classList.remove('hidden');
  } else {
    pill.textContent = 'Modo Local';
    pill.classList.remove('active');
    syncBtn.classList.add('hidden');
  }
}

function renderOverdueBanner() {
  const overdue = state.store.records.filter(isOverdue);
  const banner = el('overdueBanner');
  if (overdue.length === 0) {
    banner.classList.add('hidden');
    return;
  }
  banner.classList.remove('hidden');
  el('overdueBannerTitle').textContent = `Atenção: ${overdue.length} atendimento(s) com 3 dias ou mais sem chamado aberto!`;
  const names = overdue
    .slice(0, 3)
    .map((r) => `${r.userName} (${getDaysOpen(r)}d)`)
    .join(', ');
  el('overdueBannerText').textContent = `Solicitantes pendentes: ${names}${
    overdue.length > 3 ? ` e mais ${overdue.length - 3}...` : '.'
  }`;
}

function openOverdueModal() {
  const overdue = state.store.records.filter(isOverdue);
  const container = el('overdueModalList');
  if (overdue.length === 0) {
    container.innerHTML = '<p class="helper-text">Nenhuma pendência com 3+ dias no momento.</p>';
  } else {
    container.innerHTML = overdue
      .map(
        (r) => `
      <div class="overdue-modal-item" data-id="${escapeHtml(r.id)}">
        <div class="overdue-modal-item-top">
          <strong>${escapeHtml(r.userName)} (${escapeHtml(r.userEmail)})</strong>
          <span class="badge badge-overdue">⚠️ ${getDaysOpen(r)} dias sem chamado</span>
        </div>
        <div class="record-meta">
          <span>🛠️ Técnico: <strong>${escapeHtml(r.technician)}</strong></span>
          <span>•</span>
          <span>🕒 ${escapeHtml(formatDate(r.attendedAt || r.createdAt))}</span>
        </div>
        <div style="font-size: 13px;">${escapeHtml(r.problemDescription)}</div>
        <div class="record-actions">
          <button type="button" class="btn btn-teams btn-sm" data-overdue-action="teams-chat">
            💬 Cobrar no Teams
          </button>
          <button type="button" class="btn btn-primary btn-sm" data-overdue-action="email-preview">
            ✉️ Gerar E-mail
          </button>
          <button type="button" class="btn btn-success btn-sm" data-overdue-action="resolve">
            ✓ Regularizar
          </button>
        </div>
      </div>`
      )
      .join('');
  }
  el('overdueModal').classList.remove('hidden');
}

function closeOverdueModal() {
  el('overdueModal').classList.add('hidden');
}

function renderTechniciansOptions() {
  const techs = new Set(state.store.settings?.technicians || []);
  state.store.records.forEach((r) => {
    if (r.technician) techs.add(r.technician);
  });
  state.store.users.forEach((u) => {
    if (u.name) techs.add(u.name);
  });
  const sorted = Array.from(techs).sort((a, b) => a.localeCompare(b, 'pt-BR'));

  el('techniciansList').innerHTML = sorted
    .map((t) => `<option value="${escapeHtml(t)}"></option>`)
    .join('');

  const filter = el('techFilter');
  const current = filter.value;
  filter.innerHTML =
    '<option value="">Todos os técnicos</option>' +
    sorted.map((t) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('');
  filter.value = current;
}

function renderKPIs() {
  const records = state.store.records;
  el('kpiTotal').textContent = records.length;
  el('kpiPending').textContent = records.filter((r) => r.status === 'pending').length;
  el('kpiReminded').textContent = records.filter((r) => r.status === 'reminded').length;
  el('kpiResolved').textContent = records.filter((r) => r.status === 'resolved').length;
}

function getStatusBadge(record) {
  if (record.status === 'resolved') {
    const ticket = record.officialTicketId ? ` (${escapeHtml(record.officialTicketId)})` : '';
    return `<span class="badge badge-resolved">✓ Chamado Regularizado${ticket}</span>`;
  }
  const overdueBadge = isOverdue(record)
    ? `<span class="badge badge-overdue">⚠️ ${getDaysOpen(record)}d sem chamado</span> `
    : '';
  if (record.status === 'reminded') {
    return `${overdueBadge}<span class="badge badge-reminded">🔔 Lembrete Enviado</span>`;
  }
  return `${overdueBadge}<span class="badge badge-pending">⏳ Aguardando Chamado</span>`;
}

function getFilteredRecords() {
  const q = state.search.trim().toLowerCase();
  return [...state.store.records]
    .sort((a, b) => new Date(b.attendedAt || b.createdAt) - new Date(a.attendedAt || a.createdAt))
    .filter((r) => {
      if (state.statusFilter === 'overdue') {
        if (!isOverdue(r)) return false;
      } else if (state.statusFilter !== 'all' && r.status !== state.statusFilter) {
        return false;
      }

      if (state.techFilter && r.technician !== state.techFilter) return false;
      if (q) {
        const hay = [
          r.userName,
          r.userEmail,
          r.department,
          r.technician,
          r.category,
          r.problemDescription,
          r.officialTicketId,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
}

function renderRecords() {
  const container = el('recordsList');
  const filtered = getFilteredRecords();

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <p>Nenhum atendimento encontrado com os filtros atuais.</p>
      </div>
    `;
    return;
  }

  container.innerHTML = filtered
    .map((r) => {
      const isOwner = state.user && r.createdById === state.user.id;
      const overdueClass = isOverdue(r) ? 'is-overdue' : '';

      const historyHtml =
        r.reminders && r.reminders.length > 0
          ? `<div class="record-history">
              <strong>Histórico de lembretes (${r.reminders.length}):</strong>
              ${r.reminders
                .slice(-3)
                .reverse()
                .map(
                  (rem) =>
                    `<span>• ${escapeHtml(formatDate(rem.sentAt))} — [${escapeHtml(
                      rem.channel
                    )}] ${escapeHtml(rem.details)}${
                      rem.sentBy ? ` (por ${escapeHtml(rem.sentBy)})` : ''
                    }</span>`
                )
                .join('')}
            </div>`
          : '';

      return `
        <article class="record-item ${overdueClass}" data-id="${escapeHtml(r.id)}">
          <div class="record-top">
            <div class="record-user">
              <h3>${escapeHtml(r.userName)} ${
        r.department ? `<span class="badge badge-neutral">${escapeHtml(r.department)}</span>` : ''
      }</h3>
              <div class="record-meta">
                <span>✉️ ${escapeHtml(r.userEmail)}</span>
                <span>•</span>
                <span>🛠️ Técnico: <strong>${escapeHtml(r.technician)}</strong></span>
                <span>•</span>
                <span>🏷️ ${escapeHtml(r.category || 'Geral')}</span>
                <span>•</span>
                <span>🕒 ${escapeHtml(formatDate(r.attendedAt || r.createdAt))}</span>
                ${
                  r.createdByName
                    ? `<span>•</span><span>📝 Por: ${escapeHtml(r.createdByName)}</span>`
                    : ''
                }
              </div>
            </div>
            <div>${getStatusBadge(r)}</div>
          </div>

          <div class="record-problem">${escapeHtml(r.problemDescription)}</div>

          <div class="record-actions">
            <button type="button" class="btn btn-teams btn-sm" data-action="teams-desktop">
              🖥️ Teams Desktop
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-action="teams-chat">
              💬 Teams Web
            </button>
            <button type="button" class="btn btn-primary btn-sm" data-action="email-preview">
              ✉️ Gerar E-mail
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-action="copy">
              📋 Copiar Texto
            </button>
            ${
              r.status !== 'resolved'
                ? `<button type="button" class="btn btn-success btn-sm" data-action="resolve">
                    ✓ Regularizar
                  </button>`
                : ''
            }
            ${
              isOwner
                ? `<button type="button" class="btn btn-danger-ghost btn-sm" data-action="delete">
                    🗑️ Excluir
                  </button>`
                : ''
            }
          </div>
          ${historyHtml}
        </article>
      `;
    })
    .join('');
}

function renderAll() {
  renderDbModePill();
  renderOverdueBanner();
  renderTechniciansOptions();
  renderKPIs();
  renderRecords();
}

// Auth UI Events
document.querySelectorAll('.auth-tab').forEach((tab) => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.auth-tab').forEach((t) => t.classList.remove('active'));
    tab.classList.add('active');
    const mode = tab.dataset.mode;
    el('loginForm').classList.toggle('hidden', mode !== 'login');
    el('registerForm').classList.toggle('hidden', mode !== 'register');
  });
});

el('loginForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const sheetsInput = el('authSheetsUrl').value.trim();
  if (sheetsInput !== state.sheetsUrl) {
    state.sheetsUrl = sheetsInput;
    localStorage.setItem(SHEETS_URL_KEY, sheetsInput);
  }
  if (state.sheetsUrl) {
    await syncFromSheets(true);
  }

  const email = el('loginEmail').value.trim().toLowerCase();
  const password = el('loginPassword').value;
  const user = state.store.users.find(
    (u) => u.email.toLowerCase() === email && u.password === password
  );
  if (!user) {
    showToast('E-mail ou senha inválidos.', 'error');
    return;
  }
  state.user = user;
  localStorage.setItem(SESSION_KEY, user.id);
  showAppView();
  showToast(`Bem-vindo, ${user.name}!`, 'success');
});

el('registerForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const name = el('regName').value.trim();
  const email = el('regEmail').value.trim().toLowerCase();
  const password = el('regPassword').value;

  if (state.sheetsUrl) {
    try {
      const res = await sheetsPost({ action: 'register', name, email, password });
      state.user = res.user;
      await syncFromSheets(true);
      localStorage.setItem(SESSION_KEY, res.user.id);
      showAppView();
      showToast(`Técnico ${name} cadastrado no Google Sheets!`, 'success');
      return;
    } catch (err) {
      showToast(err.message, 'error');
      return;
    }
  }

  if (state.store.users.some((u) => u.email.toLowerCase() === email)) {
    showToast('Já existe um técnico cadastrado com este e-mail.', 'error');
    return;
  }

  const user = {
    id: crypto.randomUUID(),
    name,
    email,
    password,
    createdAt: new Date().toISOString(),
  };
  state.store.users.push(user);
  if (!state.store.settings.technicians.includes(name)) {
    state.store.settings.technicians.push(name);
  }
  state.user = user;
  addAudit('register', `Técnico ${name} (${email}) cadastrado no sistema.`);
  saveStore(state.store);
  localStorage.setItem(SESSION_KEY, user.id);
  showAppView();
  showToast(`Técnico ${name} cadastrado com sucesso!`, 'success');
});

el('logoutBtn').addEventListener('click', () => logout(true));
el('syncSheetsBtn').addEventListener('click', () => syncFromSheets(false));

// Export / Import JSON
el('exportJsonBtn').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(state.store, null, 2)], {
    type: 'application/json',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const dateTag = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `chamados-backup-${dateTag}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showToast('Backup JSON baixado!', 'success');
});

el('importJsonBtn').addEventListener('click', () => el('importFileInput').click());

el('importFileInput').addEventListener('change', async (e) => {
  const file = e.target.files?.[0];
  e.target.value = '';
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed.records) || !Array.isArray(parsed.users)) {
      throw new Error('Arquivo JSON inválido.');
    }
    state.store = {
      settings: { ...DEFAULT_SETTINGS, ...(parsed.settings || {}) },
      users: parsed.users,
      records: parsed.records,
      auditLog: Array.isArray(parsed.auditLog) ? parsed.auditLog : [],
    };
    addAudit('import', `Importou backup JSON (${file.name}) com ${parsed.records.length} registros.`);
    saveStore(state.store);
    renderAll();
    showToast(`Backup importado (${parsed.records.length} atendimentos)!`, 'success');
  } catch (err) {
    showToast(`Erro ao importar JSON: ${err.message}`, 'error');
  }
});

// Form submission
el('recordForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const btn = el('submitBtn');
  btn.disabled = true;

  const now = new Date().toISOString();
  const tech = el('technician').value.trim() || state.user.name;

  const record = {
    id: crypto.randomUUID(),
    userName: el('userName').value.trim(),
    userEmail: el('userEmail').value.trim(),
    department: el('department').value.trim(),
    category: el('category').value,
    technician: tech,
    createdById: state.user.id,
    createdByName: state.user.name,
    attendedAt: el('attendedAt').value ? new Date(el('attendedAt').value).toISOString() : now,
    createdAt: now,
    updatedAt: now,
    problemDescription: el('problemDescription').value.trim(),
    status: 'pending',
    officialTicketId: null,
    reminders: [],
  };

  try {
    if (state.sheetsUrl) {
      await sheetsPost({ action: 'create_record', record });
    }
    state.store.records.unshift(record);
    if (!state.store.settings.technicians.includes(tech)) {
      state.store.settings.technicians.push(tech);
    }
    addAudit('create', `Registrou atendimento para ${record.userName} (${record.category}) — Técnico: ${tech}`);
    saveStore(state.store);

    el('recordForm').reset();
    el('technician').value = state.user.name;
    setDefaultAttendedAt();
    renderAll();
    showToast('Atendimento registrado com sucesso!', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});

async function logReminder(record, channel, details) {
  record.reminders.push({
    id: crypto.randomUUID(),
    channel,
    sentAt: new Date().toISOString(),
    sentBy: state.user.name,
    status: 'opened',
    details,
  });
  if (record.status === 'pending') {
    record.status = 'reminded';
  }
  record.updatedAt = new Date().toISOString();
  const summary = `Enviou lembrete (${channel}) para ${record.userName} (${record.userEmail})`;
  addAudit('remind', summary);
  saveStore(state.store);
  renderAll();

  if (state.sheetsUrl) {
    try {
      await sheetsPost({
        action: 'update_record',
        id: record.id,
        record,
        actorId: state.user.id,
        actorName: state.user.name,
        auditAction: 'remind',
        auditSummary: summary,
      });
    } catch (err) {
      showToast(`Aviso: salvo localmente, mas falhou ao gravar no Sheets: ${err.message}`, 'error');
    }
  }
}

async function handleRecordAction(id, action) {
  const record = state.store.records.find((r) => r.id === id);
  if (!record) return;

  if (action === 'teams-desktop') {
    triggerProtocol(buildTeamsDeepLink(record, true));
    logReminder(record, 'teams_deeplink', `Teams Desktop acionado para ${record.userEmail}`);
    showToast(`Microsoft Teams Desktop acionado para ${record.userEmail}`, 'info');
    return;
  }

  if (action === 'teams-chat') {
    window.open(buildTeamsDeepLink(record, false), '_blank', 'noopener,noreferrer');
    logReminder(record, 'teams_deeplink', `Teams Web aberto para ${record.userEmail}`);
    showToast(`Teams Web aberto para ${record.userEmail}`, 'info');
    return;
  }

  if (action === 'email-preview') {
    openEmailPreviewModal(record);
    return;
  }

  if (action === 'copy') {
    navigator.clipboard.writeText(buildFormattedMessage(record));
    showToast('Mensagem de lembrete copiada!', 'success');
    return;
  }

  if (action === 'resolve') {
    el('resolveRecordId').value = id;
    el('officialTicketId').value = '';
    el('resolveModal').classList.remove('hidden');
    el('officialTicketId').focus();
    return;
  }

  if (action === 'delete') {
    if (record.createdById !== state.user.id) {
      showToast('Você só pode excluir atendimentos registrados por você.', 'error');
      return;
    }
    if (!confirm(`Excluir o registro de atendimento de ${record.userName}?`)) return;
    try {
      if (state.sheetsUrl) {
        await sheetsPost({
          action: 'delete_record',
          id,
          actorId: state.user.id,
          actorName: state.user.name,
        });
      }
      state.store.records = state.store.records.filter((r) => r.id !== id);
      addAudit('delete', `Excluiu o atendimento de ${record.userName} (${record.category})`);
      saveStore(state.store);
      renderAll();
      showToast('Atendimento excluído.', 'info');
    } catch (err) {
      showToast(err.message, 'error');
    }
  }
}

el('recordsList').addEventListener('click', (e) => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const article = btn.closest('.record-item');
  if (article?.dataset.id) handleRecordAction(article.dataset.id, btn.dataset.action);
});

el('overdueModalList').addEventListener('click', (e) => {
  const btn = e.target.closest('button[data-overdue-action]');
  if (!btn) return;
  const item = btn.closest('.overdue-modal-item');
  if (!item?.dataset.id) return;
  handleRecordAction(item.dataset.id, btn.dataset.overdueAction);
  if (btn.dataset.overdueAction === 'resolve') closeOverdueModal();
});

// Filters
el('searchInput').addEventListener('input', (e) => {
  state.search = e.target.value;
  renderRecords();
});

el('techFilter').addEventListener('change', (e) => {
  state.techFilter = e.target.value;
  renderRecords();
});

el('statusTabs').addEventListener('click', (e) => {
  const tab = e.target.closest('.tab[data-status]');
  if (!tab) return;
  document.querySelectorAll('#statusTabs .tab').forEach((t) => t.classList.remove('active'));
  tab.classList.add('active');
  state.statusFilter = tab.dataset.status;
  renderRecords();
});

// Overdue modal triggers
el('viewOverdueBtn').addEventListener('click', openOverdueModal);
el('closeOverdueBtn').addEventListener('click', closeOverdueModal);
el('ackOverdueBtn').addEventListener('click', closeOverdueModal);

// Audit modal
el('openAuditBtn').addEventListener('click', () => {
  const list = el('auditLogList');
  if (!state.store.auditLog.length) {
    list.innerHTML = '<p class="helper-text">Nenhuma atividade registrada ainda.</p>';
  } else {
    list.innerHTML = state.store.auditLog
      .slice(0, 200)
      .map(
        (entry) => `
        <div class="audit-item">
          <div>
            <strong>${escapeHtml(entry.actorName)}</strong> — ${escapeHtml(entry.summary)}
          </div>
          <span class="audit-time">${escapeHtml(formatDate(entry.timestamp))}</span>
        </div>`
      )
      .join('');
  }
  el('auditModal').classList.remove('hidden');
});

el('closeAuditBtn').addEventListener('click', () => el('auditModal').classList.add('hidden'));

// Resolve Modal
function closeResolveModal() {
  el('resolveModal').classList.add('hidden');
}

el('closeResolveBtn').addEventListener('click', closeResolveModal);
el('cancelResolveBtn').addEventListener('click', closeResolveModal);

el('resolveForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const id = el('resolveRecordId').value;
  const officialTicketId = el('officialTicketId').value.trim() || null;
  const record = state.store.records.find((r) => r.id === id);
  if (!record) return;

  record.status = 'resolved';
  record.officialTicketId = officialTicketId;
  record.updatedAt = new Date().toISOString();
  const ticketLabel = officialTicketId ? ` (${officialTicketId})` : '';
  const summary = `Marcou atendimento de ${record.userName} como regularizado${ticketLabel}`;

  try {
    if (state.sheetsUrl) {
      await sheetsPost({
        action: 'update_record',
        id: record.id,
        record,
        actorId: state.user.id,
        actorName: state.user.name,
        auditAction: 'resolve',
        auditSummary: summary,
      });
    }
    addAudit('resolve', summary);
    saveStore(state.store);
    renderAll();
    closeResolveModal();
    showToast('Chamado marcado como regularizado!', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  }
});

// Email Preview Modal
function openEmailPreviewModal(record) {
  el('previewRecordId').value = record.id;
  el('previewFrom').value = state.store.settings?.senderEmail || state.user?.email || '';
  el('previewTo').value = record.userEmail;
  el('previewSubject').value = `Lembrete: Abrir chamado de TI (${record.category || 'Atendimento'})`;
  el('previewBody').value = buildFormattedMessage(record);
  el('emailPreviewModal').classList.remove('hidden');
}

function closeEmailPreviewModal() {
  el('emailPreviewModal').classList.add('hidden');
}

el('closeEmailPreviewBtn').addEventListener('click', closeEmailPreviewModal);

el('previewCopyBtn').addEventListener('click', async () => {
  const full = `Para: ${el('previewTo').value}\nAssunto: ${el('previewSubject').value}\n\n${el('previewBody').value}`;
  await navigator.clipboard.writeText(full);
  showToast('E-mail completo copiado!', 'success');
});

el('previewOpenDesktopBtn').addEventListener('click', () => {
  const id = el('previewRecordId').value;
  const record = state.store.records.find((r) => r.id === id);
  if (!record) return;
  const subject = el('previewSubject').value;
  const body = el('previewBody').value.replace(/\r?\n/g, '\r\n');
  const mailtoUrl = `mailto:${encodeURIComponent(record.userEmail)}?subject=${encodeURIComponent(
    subject
  )}&body=${encodeURIComponent(body)}`;
  triggerProtocol(mailtoUrl);
  logReminder(record, 'outlook_desktop', `New Outlook Desktop acionado para ${record.userEmail}`);
  closeEmailPreviewModal();
  showToast(`New Outlook Desktop acionado para ${record.userEmail}!`, 'success');
});

el('previewOpenWebmailBtn').addEventListener('click', () => {
  const id = el('previewRecordId').value;
  const record = state.store.records.find((r) => r.id === id);
  if (!record) return;
  const url = buildWebmailLink(record, el('previewSubject').value, el('previewBody').value);
  window.open(url, '_blank', 'noopener,noreferrer');
  logReminder(record, 'outlook_web', `Webmail aberto para ${record.userEmail}`);
  closeEmailPreviewModal();
  showToast('Rascunho aberto no seu webmail! Agora é só clicar em Enviar.', 'success');
});

// Settings Modal
function openSettingsModal() {
  const s = state.store.settings;
  el('cfgSenderEmail').value = s.senderEmail || state.user?.email || '';
  el('cfgEmailProvider').value = s.emailProvider || 'outlook_office';
  el('cfgPortalUrl').value = s.portalUrl || '';
  el('cfgTemplate').value = s.reminderTemplate || '';
  el('cfgSheetsUrl').value = state.sheetsUrl || '';
  el('settingsModal').classList.remove('hidden');
}

function closeSettingsModal() {
  el('settingsModal').classList.add('hidden');
}

el('openSettingsBtn').addEventListener('click', openSettingsModal);
el('closeSettingsBtn').addEventListener('click', closeSettingsModal);
el('cancelSettingsBtn').addEventListener('click', closeSettingsModal);

el('testSheetsBtn').addEventListener('click', async () => {
  const url = el('cfgSheetsUrl').value.trim();
  if (!url) {
    showToast('Cole a URL do Apps Script (/exec) antes de testar.', 'error');
    return;
  }
  state.sheetsUrl = url;
  localStorage.setItem(SHEETS_URL_KEY, url);
  await syncFromSheets(false);
});

el('settingsForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const sheetsUrl = el('cfgSheetsUrl').value.trim();
  state.sheetsUrl = sheetsUrl;
  localStorage.setItem(SHEETS_URL_KEY, sheetsUrl);

  state.store.settings.senderEmail = el('cfgSenderEmail').value.trim();
  state.store.settings.emailProvider = el('cfgEmailProvider').value;
  state.store.settings.portalUrl = el('cfgPortalUrl').value.trim();
  state.store.settings.reminderTemplate = el('cfgTemplate').value;

  try {
    if (state.sheetsUrl) {
      await sheetsPost({
        action: 'save_settings',
        settings: state.store.settings,
        actorId: state.user.id,
        actorName: state.user.name,
      });
    }
    addAudit('settings', 'Atualizou configurações de e-mail e lembretes.');
    saveStore(state.store);
    renderDbModePill();
    closeSettingsModal();
    showToast('Configurações salvas com sucesso!', 'success');
  } catch (err) {
    showToast(err.message, 'error');
  }
});

// Init
async function bootstrap() {
  setDefaultAttendedAt();
  if (state.sheetsUrl) {
    await syncFromSheets(true);
  }
  const savedSessionId = localStorage.getItem(SESSION_KEY);
  if (savedSessionId) {
    const found = state.store.users.find((u) => u.id === savedSessionId);
    if (found) {
      state.user = found;
      showAppView();
      return;
    }
  }
  showAuthView();
}

bootstrap();
