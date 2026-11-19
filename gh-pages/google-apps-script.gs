/**
 * Chamados Express — Backend Serverless no Google Sheets (Google Apps Script)
 * Com suporte a Administrador, Lixeira/Recuperação, Gestão de Usuários e Segurança Server-Side.
 *
 * IMPORTANTE: Após atualizar este código no Apps Script, clique em:
 * Implantar -> Gerenciar implantações -> Editar (ícone de lápis) -> Versão: Nova versão -> Implantar
 */

const SHEETS = {
  RECORDS: 'Atendimentos',
  USERS: 'Tecnicos',
  AUDIT: 'Logs',
  CONFIG: 'Config',
};

const RECORD_HEADERS = [
  'id',
  'userName',
  'userEmail',
  'department',
  'technician',
  'createdById',
  'createdByName',
  'category',
  'problemDescription',
  'status',
  'officialTicketId',
  'attendedAt',
  'createdAt',
  'updatedAt',
  'remindersJson',
  'deletedAt',
  'deletedById',
  'deletedByName',
];

const USER_HEADERS = ['id', 'name', 'email', 'password', 'role', 'canManageUsers', 'createdAt'];

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  } else if (headers && headers.length) {
    // Auto-migrate missing columns on existing sheets
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    headers.forEach((h) => {
      if (!existingHeaders.includes(h)) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(h).setFontWeight('bold');
      }
    });
  }
  return sheet;
}

function ensureStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss, SHEETS.RECORDS, RECORD_HEADERS);
  const usersSheet = getOrCreateSheet(ss, SHEETS.USERS, USER_HEADERS);
  getOrCreateSheet(ss, SHEETS.AUDIT, [
    'id',
    'timestamp',
    'actorId',
    'actorName',
    'action',
    'summary',
  ]);
  const configSheet = getOrCreateSheet(ss, SHEETS.CONFIG, ['key', 'value']);

  const now = new Date().toISOString();
  const existingUsers = sheetToObjects(usersSheet);

  if (existingUsers.length === 0) {
    usersSheet.appendRow(['tech-admin', 'Administrador TI', 'admin@empresa.com.br', 'admin123', 'admin', true, now]);
    usersSheet.appendRow(['tech-joao', 'João Silva', 'joao.silva@empresa.com.br', '123456', 'tech', false, now]);
    usersSheet.appendRow(['tech-ana', 'Ana Paula', 'ana.paula@empresa.com.br', '123456', 'tech', false, now]);
  } else if (!existingUsers.some((u) => String(u.role) === 'admin')) {
    // Ensure at least one admin exists if migrating from an older sheet
    usersSheet.appendRow(['tech-admin', 'Administrador TI', 'admin@empresa.com.br', 'admin123', 'admin', true, now]);
  }

  if (configSheet.getLastRow() <= 1) {
    configSheet.appendRow([
      'settings',
      JSON.stringify({
        portalUrl: 'https://suporte.empresa.com.br/chamados',
        senderEmail: '',
        emailProvider: 'outlook_office',
        technicians: ['João Silva', 'Ana Paula', 'Carlos Eduardo'],
        reminderTemplate:
          'Olá {nome}! 👋\n\nRealizamos um atendimento de TI para você em {data} (Técnico responsável: {tecnico}).\n\n📌 Problema resolvido:\n{problema}\n\nPor favor, lembre-se de registrar o chamado formal no portal de TI para controle da equipe:\n🔗 {link_portal}\n\nMuito obrigado!',
      }),
    ]);
  }

  return ss;
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(String);
  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
}

function sanitizeUser(u) {
  const role = String(u.role || 'tech') === 'admin' ? 'admin' : 'tech';
  const canManage = role === 'admin' || u.canManageUsers === true || String(u.canManageUsers).toLowerCase() === 'true';
  return {
    id: String(u.id),
    name: String(u.name || ''),
    email: String(u.email || ''),
    role: role,
    canManageUsers: canManage,
    createdAt: String(u.createdAt || ''),
  };
}

function verifyActor(ss, actorId, actorPassword) {
  const users = sheetToObjects(ss.getSheetByName(SHEETS.USERS));
  const found = users.find((u) => String(u.id) === String(actorId));
  if (!found || String(found.password) !== String(actorPassword || '')) {
    return null;
  }
  return sanitizeUser(found);
}

function readAllState() {
  const ss = ensureStructure();
  const recordsRaw = sheetToObjects(ss.getSheetByName(SHEETS.RECORDS));
  const records = recordsRaw.map((r) => {
    let reminders = [];
    try {
      reminders = r.remindersJson ? JSON.parse(r.remindersJson) : [];
    } catch (e) {
      reminders = [];
    }
    return {
      id: String(r.id),
      userName: String(r.userName || ''),
      userEmail: String(r.userEmail || ''),
      department: String(r.department || ''),
      technician: String(r.technician || ''),
      createdById: String(r.createdById || ''),
      createdByName: String(r.createdByName || ''),
      category: String(r.category || 'Geral'),
      problemDescription: String(r.problemDescription || ''),
      status: String(r.status || 'pending'),
      officialTicketId: r.officialTicketId ? String(r.officialTicketId) : null,
      attendedAt: String(r.attendedAt || r.createdAt || ''),
      createdAt: String(r.createdAt || ''),
      updatedAt: String(r.updatedAt || ''),
      reminders: reminders,
      deletedAt: r.deletedAt ? String(r.deletedAt) : null,
      deletedById: r.deletedById ? String(r.deletedById) : null,
      deletedByName: r.deletedByName ? String(r.deletedByName) : null,
    };
  });

  // Never expose passwords in doGet
  const users = sheetToObjects(ss.getSheetByName(SHEETS.USERS)).map(sanitizeUser);

  const auditLog = sheetToObjects(ss.getSheetByName(SHEETS.AUDIT))
    .map((a) => ({
      id: String(a.id),
      timestamp: String(a.timestamp || ''),
      actorId: String(a.actorId || ''),
      actorName: String(a.actorName || ''),
      action: String(a.action || ''),
      summary: String(a.summary || ''),
    }))
    .reverse()
    .slice(0, 200);

  const cfgRows = sheetToObjects(ss.getSheetByName(SHEETS.CONFIG));
  const settingsRow = cfgRows.find((c) => c.key === 'settings');
  let settings = {};
  if (settingsRow && settingsRow.value) {
    try {
      settings = JSON.parse(settingsRow.value);
    } catch (e) {}
  }

  return { ok: true, records, users, auditLog, settings };
}

function appendAuditRow(ss, actorId, actorName, action, summary) {
  const sheet = ss.getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([
    Utilities.getUuid(),
    new Date().toISOString(),
    actorId || '',
    actorName || 'Sistema',
    action,
    summary,
  ]);
}

function buildRecordRowByHeaders(headers, r) {
  const map = {
    id: r.id,
    userName: r.userName,
    userEmail: r.userEmail,
    department: r.department || '',
    technician: r.technician,
    createdById: r.createdById || '',
    createdByName: r.createdByName || '',
    category: r.category || 'Geral',
    problemDescription: r.problemDescription,
    status: r.status || 'pending',
    officialTicketId: r.officialTicketId || '',
    attendedAt: r.attendedAt,
    createdAt: r.createdAt,
    updatedAt: r.updatedAt,
    remindersJson: JSON.stringify(r.reminders || []),
    deletedAt: r.deletedAt || '',
    deletedById: r.deletedById || '',
    deletedByName: r.deletedByName || '',
  };
  return headers.map((h) => (map[h] !== undefined ? map[h] : ''));
}

function buildUserRowByHeaders(headers, u) {
  const map = {
    id: u.id,
    name: u.name,
    email: u.email,
    password: u.password,
    role: u.role || 'tech',
    canManageUsers: Boolean(u.canManageUsers),
    createdAt: u.createdAt,
  };
  return headers.map((h) => (map[h] !== undefined ? map[h] : ''));
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function doGet() {
  try {
    return jsonOut(readAllState());
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const ss = ensureStructure();
    const action = body.action;

    if (action === 'login') {
      const users = sheetToObjects(ss.getSheetByName(SHEETS.USERS));
      const email = String(body.email || '').trim().toLowerCase();
      const password = String(body.password || '');
      const found = users.find(
        (u) => String(u.email).trim().toLowerCase() === email && String(u.password) === password
      );
      if (!found) {
        return jsonOut({ ok: false, error: 'E-mail ou senha inválidos.' });
      }
      return jsonOut({ ok: true, user: sanitizeUser(found) });
    }

    if (action === 'register') {
      // Public registration ALWAYS creates a standard 'tech' user — never 'admin'
      const usersSheet = ss.getSheetByName(SHEETS.USERS);
      const users = sheetToObjects(usersSheet);
      const headers = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0].map(String);
      const email = String(body.email || '').trim().toLowerCase();
      if (users.some((u) => String(u.email).trim().toLowerCase() === email)) {
        return jsonOut({ ok: false, error: 'Já existe um técnico cadastrado com este e-mail.' });
      }
      const userObj = {
        id: Utilities.getUuid(),
        name: String(body.name || '').trim(),
        email: email,
        password: String(body.password || ''),
        role: 'tech',
        canManageUsers: false,
        createdAt: new Date().toISOString(),
      };
      usersSheet.appendRow(buildUserRowByHeaders(headers, userObj));
      appendAuditRow(ss, userObj.id, userObj.name, 'register', `Técnico ${userObj.name} (${userObj.email}) cadastrado.`);
      return jsonOut({ ok: true, user: sanitizeUser(userObj) });
    }

    if (action === 'create_record') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor) return jsonOut({ ok: false, error: 'Sessão inválida. Faça login novamente.' });

      const r = body.record;
      r.createdById = actor.id;
      r.createdByName = actor.name;
      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      sheet.appendRow(buildRecordRowByHeaders(headers, r));
      appendAuditRow(
        ss,
        actor.id,
        actor.name,
        'create',
        `Registrou atendimento para ${r.userName} (${r.category}) — Técnico: ${r.technician}`
      );
      return jsonOut({ ok: true, record: r });
    }

    if (action === 'update_record') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor) return jsonOut({ ok: false, error: 'Sessão inválida.' });

      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const idIdx = headers.indexOf('id');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });

      const newRow = buildRecordRowByHeaders(headers, body.record);
      sheet.getRange(rowIdx + 1, 1, 1, newRow.length).setValues([newRow]);
      if (body.auditSummary) {
        appendAuditRow(ss, actor.id, actor.name, body.auditAction || 'update', body.auditSummary);
      }
      return jsonOut({ ok: true, record: body.record });
    }

    if (action === 'delete_record') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor) return jsonOut({ ok: false, error: 'Credenciais inválidas para excluir.' });

      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const idIdx = headers.indexOf('id');
      const ownerIdx = headers.indexOf('createdById');
      const nameIdx = headers.indexOf('userName');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });

      const isOwner = String(data[rowIdx][ownerIdx]) === String(actor.id);
      if (actor.role !== 'admin' && !isOwner) {
        return jsonOut({ ok: false, error: 'Apenas o administrador ou quem registrou pode excluir.' });
      }

      const now = new Date().toISOString();
      const delAtIdx = headers.indexOf('deletedAt');
      const delByIdIdx = headers.indexOf('deletedById');
      const delByNameIdx = headers.indexOf('deletedByName');
      sheet.getRange(rowIdx + 1, delAtIdx + 1).setValue(now);
      sheet.getRange(rowIdx + 1, delByIdIdx + 1).setValue(actor.id);
      sheet.getRange(rowIdx + 1, delByNameIdx + 1).setValue(actor.name);

      appendAuditRow(ss, actor.id, actor.name, 'trash', `Moveu atendimento de ${data[rowIdx][nameIdx]} para a lixeira.`);
      return jsonOut({ ok: true, deletedAt: now, deletedById: actor.id, deletedByName: actor.name });
    }

    if (action === 'restore_record') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor) return jsonOut({ ok: false, error: 'Credenciais inválidas.' });

      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const idIdx = headers.indexOf('id');
      const delByIdIdx = headers.indexOf('deletedById');
      const ownerIdx = headers.indexOf('createdById');
      const nameIdx = headers.indexOf('userName');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });

      const canRestore =
        actor.role === 'admin' ||
        String(data[rowIdx][delByIdIdx]) === String(actor.id) ||
        String(data[rowIdx][ownerIdx]) === String(actor.id);
      if (!canRestore) return jsonOut({ ok: false, error: 'Sem permissão para restaurar.' });

      sheet.getRange(rowIdx + 1, headers.indexOf('deletedAt') + 1).setValue('');
      sheet.getRange(rowIdx + 1, delByIdIdx + 1).setValue('');
      sheet.getRange(rowIdx + 1, headers.indexOf('deletedByName') + 1).setValue('');

      appendAuditRow(ss, actor.id, actor.name, 'restore', `Restaurou atendimento de ${data[rowIdx][nameIdx]} da lixeira.`);
      return jsonOut({ ok: true });
    }

    if (action === 'purge_record') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor || actor.role !== 'admin') {
        return jsonOut({ ok: false, error: 'Apenas o administrador pode excluir definitivamente.' });
      }
      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const idIdx = headers.indexOf('id');
      const nameIdx = headers.indexOf('userName');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });

      const targetName = data[rowIdx][nameIdx];
      sheet.deleteRow(rowIdx + 1);
      appendAuditRow(ss, actor.id, actor.name, 'purge', `Excluiu permanentemente o atendimento de ${targetName}.`);
      return jsonOut({ ok: true });
    }

    if (action === 'manage_user') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor || (actor.role !== 'admin' && !actor.canManageUsers)) {
        return jsonOut({ ok: false, error: 'Sem permissão para gerenciar usuários.' });
      }

      const usersSheet = ss.getSheetByName(SHEETS.USERS);
      const data = usersSheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const op = body.op;

      if (op === 'create') {
        const email = String(body.email || '').trim().toLowerCase();
        const existing = sheetToObjects(usersSheet);
        if (existing.some((u) => String(u.email).trim().toLowerCase() === email)) {
          return jsonOut({ ok: false, error: 'E-mail já cadastrado.' });
        }
        // Only admin can grant canManageUsers; never allow creating another admin via API
        const canManage = actor.role === 'admin' ? Boolean(body.canManageUsers) : false;
        const userObj = {
          id: Utilities.getUuid(),
          name: String(body.name || '').trim(),
          email: email,
          password: String(body.password || ''),
          role: 'tech',
          canManageUsers: canManage,
          createdAt: new Date().toISOString(),
        };
        usersSheet.appendRow(buildUserRowByHeaders(headers, userObj));
        appendAuditRow(ss, actor.id, actor.name, 'user_create', `Criou usuário ${userObj.name} (${userObj.email}).`);
        return jsonOut({ ok: true, user: sanitizeUser(userObj) });
      }

      const idIdx = headers.indexOf('id');
      const roleIdx = headers.indexOf('role');
      const nameIdx = headers.indexOf('name');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.targetUserId));
      if (rowIdx === -1) return jsonOut({ ok: false, error: 'Usuário não encontrado.' });

      const targetRole = String(data[rowIdx][roleIdx] || 'tech');
      const targetName = String(data[rowIdx][nameIdx] || '');

      // Nobody non-admin can touch an admin account
      if (targetRole === 'admin' && actor.role !== 'admin') {
        return jsonOut({ ok: false, error: 'Você não pode alterar a conta do administrador.' });
      }

      if (op === 'toggle_manage') {
        if (actor.role !== 'admin') {
          return jsonOut({ ok: false, error: 'Apenas o administrador pode conceder permissão de gerenciamento.' });
        }
        if (targetRole === 'admin') {
          return jsonOut({ ok: false, error: 'O administrador já possui todas as permissões.' });
        }
        const canManageIdx = headers.indexOf('canManageUsers');
        const nextVal = Boolean(body.canManageUsers);
        usersSheet.getRange(rowIdx + 1, canManageIdx + 1).setValue(nextVal);
        appendAuditRow(
          ss,
          actor.id,
          actor.name,
          'user_perm',
          `${nextVal ? 'Concedeu' : 'Removeu'} permissão de gerenciar usuários para ${targetName}.`
        );
        return jsonOut({ ok: true });
      }

      if (op === 'reset_password') {
        const passIdx = headers.indexOf('password');
        usersSheet.getRange(rowIdx + 1, passIdx + 1).setValue(String(body.newPassword || ''));
        appendAuditRow(ss, actor.id, actor.name, 'user_pass', `Redefiniu a senha de ${targetName}.`);
        return jsonOut({ ok: true });
      }

      if (op === 'delete') {
        if (targetRole === 'admin' || String(body.targetUserId) === String(actor.id)) {
          return jsonOut({ ok: false, error: 'Não é possível excluir o administrador ou a si mesmo.' });
        }
        usersSheet.deleteRow(rowIdx + 1);
        appendAuditRow(ss, actor.id, actor.name, 'user_delete', `Removeu o usuário ${targetName}.`);
        return jsonOut({ ok: true });
      }

      return jsonOut({ ok: false, error: 'Operação inválida.' });
    }

    if (action === 'save_settings') {
      const actor = verifyActor(ss, body.actorId, body.actorPassword);
      if (!actor) return jsonOut({ ok: false, error: 'Sessão inválida.' });
      const sheet = ss.getSheetByName(SHEETS.CONFIG);
      const data = sheet.getDataRange().getValues();
      const rowIdx = data.findIndex((row, idx) => idx > 0 && row[0] === 'settings');
      const val = JSON.stringify(body.settings || {});
      if (rowIdx === -1) {
        sheet.appendRow(['settings', val]);
      } else {
        sheet.getRange(rowIdx + 1, 2).setValue(val);
      }
      appendAuditRow(ss, actor.id, actor.name, 'settings', 'Atualizou configurações.');
      return jsonOut({ ok: true });
    }

    return jsonOut({ ok: false, error: 'Ação desconhecida.' });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}
