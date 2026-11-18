/**
 * Chamados Express — Backend Serverless no Google Sheets (Google Apps Script)
 *
 * Como usar:
 * 1. Crie uma planilha no Google Sheets (https://sheets.new)
 * 2. Vá em Extensões -> Apps Script
 * 3. Apague o código padrão e cole este arquivo inteiro
 * 4. Clique em Salvar e depois em Implantar -> Nova implantação
 * 5. Tipo: App da Web | Executar como: Eu | Quem pode acessar: Qualquer pessoa
 * 6. Copie a URL gerada (/exec) e cole em "URL do Google Sheets (Apps Script)" no site!
 */

const SHEETS = {
  RECORDS: 'Atendimentos',
  USERS: 'Tecnicos',
  AUDIT: 'Logs',
  CONFIG: 'Config',
};

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function ensureStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordsSheet = getOrCreateSheet(ss, SHEETS.RECORDS, [
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
  ]);

  const usersSheet = getOrCreateSheet(ss, SHEETS.USERS, [
    'id',
    'name',
    'email',
    'password',
    'createdAt',
  ]);

  getOrCreateSheet(ss, SHEETS.AUDIT, [
    'id',
    'timestamp',
    'actorId',
    'actorName',
    'action',
    'summary',
  ]);

  const configSheet = getOrCreateSheet(ss, SHEETS.CONFIG, ['key', 'value']);

  // Seed initial technicians if empty
  if (usersSheet.getLastRow() <= 1) {
    const now = new Date().toISOString();
    usersSheet.appendRow(['tech-joao', 'João Silva', 'joao.silva@empresa.com.br', '123456', now]);
    usersSheet.appendRow(['tech-ana', 'Ana Paula', 'ana.paula@empresa.com.br', '123456', now]);
  }

  // Seed default settings if empty
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
  const headers = values[0];
  return values.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
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
    };
  });

  const users = sheetToObjects(ss.getSheetByName(SHEETS.USERS)).map((u) => ({
    id: String(u.id),
    name: String(u.name || ''),
    email: String(u.email || ''),
    password: String(u.password || ''),
    createdAt: String(u.createdAt || ''),
  }));

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

    if (action === 'register') {
      const usersSheet = ss.getSheetByName(SHEETS.USERS);
      const users = sheetToObjects(usersSheet);
      const email = String(body.email || '').trim().toLowerCase();
      if (users.some((u) => String(u.email).toLowerCase() === email)) {
        return jsonOut({ ok: false, error: 'Já existe um técnico cadastrado com este e-mail.' });
      }
      const user = {
        id: Utilities.getUuid(),
        name: String(body.name || '').trim(),
        email: email,
        password: String(body.password || ''),
        createdAt: new Date().toISOString(),
      };
      usersSheet.appendRow([user.id, user.name, user.email, user.password, user.createdAt]);
      appendAuditRow(ss, user.id, user.name, 'register', `Técnico ${user.name} (${user.email}) cadastrado.`);
      return jsonOut({ ok: true, user });
    }

    if (action === 'create_record') {
      const r = body.record;
      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      sheet.appendRow([
        r.id,
        r.userName,
        r.userEmail,
        r.department || '',
        r.technician,
        r.createdById || '',
        r.createdByName || '',
        r.category || 'Geral',
        r.problemDescription,
        r.status || 'pending',
        r.officialTicketId || '',
        r.attendedAt,
        r.createdAt,
        r.updatedAt,
        JSON.stringify(r.reminders || []),
      ]);
      appendAuditRow(
        ss,
        r.createdById,
        r.createdByName || r.technician,
        'create',
        `Registrou atendimento para ${r.userName} (${r.category}) — Técnico: ${r.technician}`
      );
      return jsonOut({ ok: true, record: r });
    }

    if (action === 'update_record') {
      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf('id');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) {
        return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });
      }
      const updated = body.record;
      const newRow = [
        updated.id,
        updated.userName,
        updated.userEmail,
        updated.department || '',
        updated.technician,
        updated.createdById || '',
        updated.createdByName || '',
        updated.category || 'Geral',
        updated.problemDescription,
        updated.status || 'pending',
        updated.officialTicketId || '',
        updated.attendedAt,
        updated.createdAt,
        updated.updatedAt,
        JSON.stringify(updated.reminders || []),
      ];
      sheet.getRange(rowIdx + 1, 1, 1, newRow.length).setValues([newRow]);
      if (body.auditSummary) {
        appendAuditRow(ss, body.actorId, body.actorName, body.auditAction || 'update', body.auditSummary);
      }
      return jsonOut({ ok: true, record: updated });
    }

    if (action === 'delete_record') {
      const sheet = ss.getSheetByName(SHEETS.RECORDS);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf('id');
      const ownerIdx = headers.indexOf('createdById');
      const nameIdx = headers.indexOf('userName');
      const rowIdx = data.findIndex((row, idx) => idx > 0 && String(row[idIdx]) === String(body.id));
      if (rowIdx === -1) {
        return jsonOut({ ok: false, error: 'Atendimento não encontrado.' });
      }
      if (String(data[rowIdx][ownerIdx]) !== String(body.actorId)) {
        return jsonOut({ ok: false, error: 'Apenas o técnico que registrou pode excluir.' });
      }
      const targetUser = data[rowIdx][nameIdx];
      sheet.deleteRow(rowIdx + 1);
      appendAuditRow(ss, body.actorId, body.actorName, 'delete', `Excluiu atendimento de ${targetUser}`);
      return jsonOut({ ok: true });
    }

    if (action === 'save_settings') {
      const sheet = ss.getSheetByName(SHEETS.CONFIG);
      const data = sheet.getDataRange().getValues();
      const rowIdx = data.findIndex((row, idx) => idx > 0 && row[0] === 'settings');
      const val = JSON.stringify(body.settings || {});
      if (rowIdx === -1) {
        sheet.appendRow(['settings', val]);
      } else {
        sheet.getRange(rowIdx + 1, 2).setValue(val);
      }
      appendAuditRow(ss, body.actorId, body.actorName, 'settings', 'Atualizou configurações.');
      return jsonOut({ ok: true });
    }

    return jsonOut({ ok: false, error: 'Ação desconhecida.' });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}
