const INVITADOS_SHEET_ID = '1XL1UnsIBacWrUnxWaJfN5mcmmwdKk6M6iuEwVMCu47o';
const INVITADOS_TAB      = 'Invitados';
const RESPUESTAS_TAB     = 'RSVPs';

function setCorsHeaders(output) {
  return output.setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJson(e, obj) {
  const json = JSON.stringify(obj);
  const cb   = e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return setCorsHeaders(ContentService.createTextOutput(json));
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'search') {
    const query = (e.parameter.q || '').toLowerCase().trim();
    if (query.length < 2) return jsonpOrJson(e, { results: [] });
    const ss      = SpreadsheetApp.openById(INVITADOS_SHEET_ID);
    const sheet   = ss.getSheetByName(INVITADOS_TAB);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxName  = headers.indexOf('name');
    const idxEmail = headers.indexOf('email');
    const idxWa    = headers.indexOf('whatsapp');
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const name = (data[i][idxName] || '').toString().toLowerCase();
      if (name.includes(query)) {
        results.push({ name: data[i][idxName] || '', email: data[i][idxEmail] || '', whatsapp: data[i][idxWa] || '' });
      }
      if (results.length >= 8) break;
    }
    return jsonpOrJson(e, { results });
  }

  if (action === 'person') {
    const name = (e.parameter.name || '').toLowerCase().trim();
    const ss   = SpreadsheetApp.openById(INVITADOS_SHEET_ID);
    const rsvpSheet = ss.getSheetByName(RESPUESTAS_TAB);
    if (rsvpSheet) {
      const data    = rsvpSheet.getDataRange().getValues();
      const headers = data[0];
      const idxName = headers.indexOf('name');
      for (let i = 1; i < data.length; i++) {
        const n = (data[i][idxName] || '').toString().toLowerCase().trim();
        if (n === name) {
          const row = {};
          headers.forEach((h, j) => { row[h] = data[i][j] !== undefined ? data[i][j].toString() : ''; });
          return jsonpOrJson(e, { found: true, data: row });
        }
      }
    }
    return jsonpOrJson(e, { found: false });
  }

  if (action === 'rsvp') {
    const ss    = SpreadsheetApp.openById(INVITADOS_SHEET_ID);
    const sheet = ss.getSheetByName(RESPUESTAS_TAB);
    if (!sheet) return jsonpOrJson(e, { ok: true, data: [] });
    const values  = sheet.getDataRange().getValues();
    if (values.length < 2) return jsonpOrJson(e, { ok: true, data: [] });
    const headers = values[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
    const data    = values.slice(1)
      .filter(r => r[0])
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = r[i] != null ? String(r[i]) : ''; });
        return obj;
      });
    return jsonpOrJson(e, { ok: true, data });
  }

  return jsonpOrJson(e, { error: 'unknown action' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if      (payload.type === 'rsvp')           handleRSVP(payload);
    else if (payload.type === 'reconfirmation') handleReconfirmation(payload);
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok' })));
  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })));
  }
}

function handleRSVP(p) {
  const ss  = SpreadsheetApp.openById(INVITADOS_SHEET_ID);
  let sheet = ss.getSheetByName(RESPUESTAS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(RESPUESTAS_TAB);
    sheet.appendRow([
      'timestamp','name','email','whatsapp','attendance','attendance2','changes',
      'uk_arrival','wedding_part','menu_starter','menu_main','menu_dessert',
      'diet','springkell_accommodation','rooms','accessibility','other_accommodation',
      'reconf_transp_arrival','reconf_transp_pickup_time','reconf_transp_departure','reconf_diet',
      'reconf_accommodation','reconf_rooms','reconf_notes','reconf_wedding_part',
      'reconf_menu_starter','reconf_menu_main','reconf_menu_dessert'
    ]);
    const r = sheet.getRange(1, 1, 1, 28);
    r.setFontWeight('bold'); r.setBackground('#2e2018'); r.setFontColor('#f2ece0');
    sheet.setFrozenRows(1);
  }
  const data     = sheet.getDataRange().getValues();
  const hdrs     = data[0];
  const idxName  = hdrs.indexOf('name');
  const fieldMap = {
    'timestamp':                p.timestamp || new Date().toISOString(),
    'name':                     p.name || '',
    'email':                    p.email || '',
    'whatsapp':                 p.whatsapp || '',
    'attendance':               p.attendance || '',
    'uk_arrival':               p.uk_arrival || '',
    'wedding_part':             p.wedding_part || '',
    'menu_starter':             p.menu_starter || '',      // ← NUEVO
    'menu_main':                p.menu_main || '',         // ← NUEVO
    'menu_dessert':             p.menu_dessert || '',      // ← NUEVO
    'diet':                     p.diet || '',
    'springkell_accommodation': p.springkell_accommodation || '',
    'rooms':                    p.rooms || '',
    'accessibility':            p.accessibility || '',
    'other_accommodation':      p.other_accommodation || ''
  };
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][idxName] || '').toString().toLowerCase().trim();
    if (n === (p.name || '').toLowerCase().trim()) {
      Object.keys(fieldMap).forEach(function(key) {
        const idx = hdrs.indexOf(key);
        if (idx >= 0) sheet.getRange(i+1, idx+1).setValue(fieldMap[key]);
      });
      return;
    }
  }
  const newRow = hdrs.map(function(h) { return fieldMap[h] !== undefined ? fieldMap[h] : ''; });
  sheet.appendRow(newRow);
}

function handleReconfirmation(p) {
  const ss  = SpreadsheetApp.openById(INVITADOS_SHEET_ID);
  let sheet = ss.getSheetByName(RESPUESTAS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(RESPUESTAS_TAB);
    sheet.appendRow([
      'timestamp','name','email','whatsapp','attendance','attendance2','changes',
      'uk_arrival','wedding_part','menu_starter','menu_main','menu_dessert',
      'diet','springkell_accommodation','rooms','accessibility','other_accommodation',
      'reconf_transp_arrival','reconf_transp_pickup_time','reconf_transp_departure','reconf_diet',
      'reconf_accommodation','reconf_rooms','reconf_notes','reconf_wedding_part',
      'reconf_menu_starter','reconf_menu_main','reconf_menu_dessert'
    ]);
    const r = sheet.getRange(1, 1, 1, 28);
    r.setFontWeight('bold'); r.setBackground('#2e2018'); r.setFontColor('#f2ece0');
    sheet.setFrozenRows(1);
  }
  // Si la hoja ya existía sin esta columna, la agregamos al final
  if (sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('reconf_transp_pickup_time') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('reconf_transp_pickup_time');
  }

  const data     = sheet.getDataRange().getValues();
  const hdrs     = data[0];
  const idxName  = hdrs.indexOf('name');
  const fieldMap = {
    'attendance2':             p.attendance2 || '',
    'changes':                 p.changes || '',
    'reconf_transp_arrival':   p.reconf_transp_arrival || '',
    'reconf_transp_pickup_time': p.reconf_transp_pickup_time || '',
    'reconf_transp_departure': p.reconf_transp_departure || '',
    'reconf_diet':             p.reconf_diet || '',
    'reconf_accommodation':    p.reconf_accommodation || '',
    'reconf_rooms':            p.reconf_rooms || '',
    'reconf_notes':            p.reconf_notes || '',
    'reconf_wedding_part':     p.reconf_wedding_part || '',
    'reconf_menu_starter':     p.reconf_menu_starter || '',   // ← NUEVO
    'reconf_menu_main':        p.reconf_menu_main || '',      // ← NUEVO
    'reconf_menu_dessert':     p.reconf_menu_dessert || ''    // ← NUEVO
  };
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][idxName] || '').toString().toLowerCase().trim();
    if (n === (p.name || '').toLowerCase().trim()) {
      Object.keys(fieldMap).forEach(function(key) {
        const idx = hdrs.indexOf(key);
        if (idx >= 0) sheet.getRange(i+1, idx+1).setValue(fieldMap[key]);
      });
      return;
    }
  }
  const newRow = hdrs.map(function(h) { return fieldMap[h] !== undefined ? fieldMap[h] : ''; });
  sheet.appendRow(newRow);
}
