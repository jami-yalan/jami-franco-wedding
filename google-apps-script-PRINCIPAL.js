/**
 * ═══════════════════════════════════════════════════════════
 *  JAMI & FRANCO WEDDING — Google Apps Script PRINCIPAL
 *  Pega este código en: script.google.com → Nuevo proyecto
 *  Llámalo: "JF Wedding - Principal"
 * ═══════════════════════════════════════════════════════════
 *
 *  ESTE SCRIPT MANEJA:
 *    · RSVPs
 *    · Reconfirmaciones
 *    · Regalos por meta (con comprobante → Drive)
 *    · Contribuciones libres (con comprobante → Drive)
 *    · Mensajes del Livestream
 *
 *  CONFIGURACIÓN:
 *  1. Ve a script.google.com → Nuevo proyecto
 *  2. Pega este código completo
 *  3. Cambia SHEET_ID por el ID de tu Google Sheet
 *  4. Cambia DRIVE_FOLDER_ID por el ID de la carpeta de Drive
 *     donde quieres guardar los comprobantes
 *     (abre la carpeta en Drive → la URL tiene /folders/ESTE_ES_EL_ID)
 *  5. Implementar → Nueva implementación → Aplicación web
 *     · Ejecutar como: Yo
 *     · Quién tiene acceso: Cualquier persona
 *  6. Copia la URL y pégala en jami-franco-wedding.html y
 *     livestream.html donde dice YOUR_APPS_SCRIPT_URL_HERE
 *     También en rsvp.html y rsvp-reconfirmacion.html
 * ═══════════════════════════════════════════════════════════
 */

const SHEET_ID       = 'TU_GOOGLE_SHEET_ID_AQUI';       // ← Cambia esto
const DRIVE_FOLDER_ID = 'TU_CARPETA_DRIVE_ID_AQUI';      // ← Cambia esto (carpeta para comprobantes)

const SHEET_RSVP    = 'RSVPs';
const SHEET_RECONF  = 'Reconfirmaciones';
const SHEET_REGALOS = 'Regalos';

// ─────────────────────────────────────────────
//  CORS
// ─────────────────────────────────────────────
function setCorsHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ─────────────────────────────────────────────
//  GET — Búsqueda de nombres y datos de RSVP
// ─────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  // Búsqueda por nombre (autocomplete del RSVP2)
  if (action === 'search') {
    const query = (e.parameter.q || '').toLowerCase().trim();
    if (query.length < 2) {
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ results: [] })));
    }
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_RSVP);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxNombre   = headers.indexOf('nombre');
    const idxApellido = headers.indexOf('apellido');
    const idxEmail    = headers.indexOf('email');
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const nombre   = (row[idxNombre]   || '').toString().toLowerCase();
      const apellido = (row[idxApellido] || '').toString().toLowerCase();
      const full     = nombre + ' ' + apellido;
      if (full.includes(query) || nombre.includes(query) || apellido.includes(query)) {
        results.push({ nombre: row[idxNombre], apellido: row[idxApellido], email: row[idxEmail] || '' });
      }
      if (results.length >= 8) break;
    }
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ results })));
  }

  // Datos completos de una persona para pre-cargar reconfirmación
  if (action === 'person') {
    const nombre   = (e.parameter.nombre   || '').toLowerCase().trim();
    const apellido = (e.parameter.apellido || '').toLowerCase().trim();
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_RSVP);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxNombre       = headers.indexOf('nombre');
    const idxApellido     = headers.indexOf('apellido');
    const idxAlojamiento  = headers.indexOf('alojamiento_springkell');
    const idxHabitaciones = headers.indexOf('habitaciones');
    for (let i = 1; i < data.length; i++) {
      const n = (data[i][idxNombre]   || '').toString().toLowerCase().trim();
      const a = (data[i][idxApellido] || '').toString().toLowerCase().trim();
      if (n === nombre && a === apellido) {
        return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({
          found: true,
          alojamiento_springkell: idxAlojamiento  >= 0 ? (data[i][idxAlojamiento]  || '').toString() : '',
          habitaciones:           idxHabitaciones >= 0 ? (data[i][idxHabitaciones] || '').toString() : ''
        })));
      }
    }
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ found: false })));
  }

  // Totales confirmados por meta (para las barras de progreso)
  // Llamada: ?action=totales
  if (action === 'totales') {
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_REGALOS);
    if (!sheet) {
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ totales: {} })));
    }
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxMeta   = headers.indexOf('meta');
    const idxMonto  = headers.indexOf('monto_usd');
    const idxOrigen = headers.indexOf('origen');

    // Mapeo de nombre de meta → id del gift en el HTML
    const metaMap = {
      'Pasajes de luna de miel a Grecia': 1,
      'Honeymoon flights to Greece':      1,
      'Estadía en Grecia':                2,
      'Accommodation in Greece':          2,
      'Museos en Grecia':                 3,
      'Museums in Greece':                3,
      'Pasajes para visitar islas en Grecia': 4,
      'Flights to Greek islands':         4,
      'Visita Ítaca':                     5,
      'Visit Ithaca':                     5,
      'Viáticos':                         6,
      'Daily expenses':                   6,
      'Tours':                            7
    };

    const totales = {};
    for (let i = 1; i < data.length; i++) {
      const meta   = (data[i][idxMeta]   || '').toString().trim();
      const monto  = parseFloat(data[i][idxMonto]) || 0;
      const origen = (data[i][idxOrigen] || '').toString().trim();
      // Solo contar filas que vienen de 'home' (no livestream) y tienen monto numérico
      if (origen !== 'home' || monto <= 0) continue;
      const giftId = metaMap[meta];
      if (giftId) {
        totales[giftId] = (totales[giftId] || 0) + monto;
      }
    }
    // Soporte JSONP: si viene ?callback=nombre, envolver la respuesta
    const callback = e.parameter.callback;
    const json = JSON.stringify({ totales });
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return setCorsHeaders(ContentService.createTextOutput(json));
  }

  return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' })));
}

// ─────────────────────────────────────────────
//  POST — Router principal
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if      (payload.type === 'rsvp')           handleRSVP(payload);
    else if (payload.type === 'reconfirmation') handleReconfirmation(payload);
    else if (payload.type === 'regalo')         handleRegalo(payload);
    else if (payload.type === 'livestream_msg') handleLivestreamMsg(payload);
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok' })));
  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })));
  }
}

// ─────────────────────────────────────────────
//  RSVP
// ─────────────────────────────────────────────
function handleRSVP(p) {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_RSVP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RSVP);
    sheet.appendRow([
      'timestamp','nombre','apellido','email','whatsapp',
      'asistencia','llegada_uk','dieta','alojamiento_springkell',
      'habitaciones','accesibilidad','alojamiento_otro','notas'
    ]);
    const r = sheet.getRange(1, 1, 1, 13);
    r.setFontWeight('bold'); r.setBackground('#2e2018'); r.setFontColor('#f2ece0');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    p.timestamp || new Date().toISOString(),
    p.nombre, p.apellido, p.email, p.whatsapp,
    p.asistencia, p.llegada, p.dieta,
    p.alojamiento_springkell, p.habitaciones,
    p.accesibilidad, p.alojamiento_otro, p.notas
  ]);
}

// ─────────────────────────────────────────────
//  RECONFIRMACIÓN
// ─────────────────────────────────────────────
function handleReconfirmation(p) {
  const ss        = SpreadsheetApp.openById(SHEET_ID);
  const rsvpSheet = ss.getSheetByName(SHEET_RSVP);

  if (rsvpSheet) {
    const data    = rsvpSheet.getDataRange().getValues();
    const headers = data[0];
    const idxNombre   = headers.indexOf('nombre');
    const idxApellido = headers.indexOf('apellido');
    const lastCol     = headers.length;

    let idxReconfAsist   = headers.indexOf('reconf_asistencia');
    let idxReconfLlegada = headers.indexOf('reconf_transp_llegada');
    let idxReconfSalida  = headers.indexOf('reconf_transp_salida');
    let idxReconfDieta   = headers.indexOf('reconf_dieta_cambio');
    let idxReconfAloj    = headers.indexOf('reconf_alojamiento');
    let idxReconfHabs    = headers.indexOf('reconf_habitaciones');
    let idxReconfNotas   = headers.indexOf('reconf_notas');
    let idxReconfTS      = headers.indexOf('reconf_timestamp');

    if (idxReconfAsist === -1) {
      const newCols = [
        'reconf_asistencia','reconf_transp_llegada','reconf_transp_salida',
        'reconf_dieta_cambio','reconf_alojamiento','reconf_habitaciones',
        'reconf_notas','reconf_timestamp'
      ];
      newCols.forEach((col, i) => rsvpSheet.getRange(1, lastCol + 1 + i).setValue(col));
      idxReconfAsist   = lastCol;
      idxReconfLlegada = lastCol + 1;
      idxReconfSalida  = lastCol + 2;
      idxReconfDieta   = lastCol + 3;
      idxReconfAloj    = lastCol + 4;
      idxReconfHabs    = lastCol + 5;
      idxReconfNotas   = lastCol + 6;
      idxReconfTS      = lastCol + 7;
    }

    for (let i = 1; i < data.length; i++) {
      const n = (data[i][idxNombre]   || '').toString().toLowerCase().trim();
      const a = (data[i][idxApellido] || '').toString().toLowerCase().trim();
      if (n === p.nombre.toLowerCase().trim() && a === p.apellido.toLowerCase().trim()) {
        rsvpSheet.getRange(i+1, idxReconfAsist+1).setValue(p.asistencia_reconf);
        rsvpSheet.getRange(i+1, idxReconfLlegada+1).setValue(p.transporte_llegada);
        rsvpSheet.getRange(i+1, idxReconfSalida+1).setValue(p.transporte_salida);
        rsvpSheet.getRange(i+1, idxReconfDieta+1).setValue(p.dieta_cambio);
        rsvpSheet.getRange(i+1, idxReconfAloj+1).setValue(p.alojamiento_reconf);
        rsvpSheet.getRange(i+1, idxReconfHabs+1).setValue(p.habitaciones_reconf);
        rsvpSheet.getRange(i+1, idxReconfNotas+1).setValue(p.notas_reconf);
        rsvpSheet.getRange(i+1, idxReconfTS+1).setValue(p.timestamp_reconf);
        break;
      }
    }
  }

  let reconfSheet = ss.getSheetByName(SHEET_RECONF);
  if (!reconfSheet) {
    reconfSheet = ss.insertSheet(SHEET_RECONF);
    reconfSheet.appendRow([
      'timestamp','nombre','apellido','email',
      'asistencia','transporte_llegada','transporte_salida',
      'dieta_cambio','alojamiento','habitaciones','notas'
    ]);
    const hr = reconfSheet.getRange(1, 1, 1, 11);
    hr.setFontWeight('bold'); hr.setBackground('#2e2018'); hr.setFontColor('#f2ece0');
    reconfSheet.setFrozenRows(1);
  }
  reconfSheet.appendRow([
    p.timestamp_reconf || new Date().toISOString(),
    p.nombre, p.apellido, p.email || '',
    p.asistencia_reconf, p.transporte_llegada, p.transporte_salida,
    p.dieta_cambio, p.alojamiento_reconf, p.habitaciones_reconf, p.notas_reconf
  ]);
}

// ─────────────────────────────────────────────
//  REGALOS (metas + contribución libre)
//  Guarda comprobante en Drive si viene adjunto
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
//  Helper: escribir en la hoja Regalos
//  Usada por regalos normales Y mensajes del livestream
// ─────────────────────────────────────────────
function escribirEnRegalos(ss, fila) {
  let sheet = ss.getSheetByName(SHEET_REGALOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REGALOS);
    sheet.appendRow([
      'timestamp','origen','nombre','meta','monto_usd','mensaje',
      'archivo_comprobante','url_comprobante_drive'
    ]);
    const r = sheet.getRange(1, 1, 1, 8);
    r.setFontWeight('bold'); r.setBackground('#2e2018'); r.setFontColor('#f2ece0');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow(fila);
}

// ─────────────────────────────────────────────
//  REGALOS (metas + contribución libre)
//  Guarda comprobante en Drive si viene adjunto
// ─────────────────────────────────────────────
function handleRegalo(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Subir comprobante a Drive si viene en base64
  let driveUrl = '';
  if (p.archivo_data && p.archivo_nombre) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const bytes  = Utilities.base64Decode(p.archivo_data);
      const blob   = Utilities.newBlob(bytes, p.archivo_tipo || 'application/octet-stream', p.archivo_nombre);
      const file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveUrl = file.getUrl();
    } catch(err) {
      driveUrl = 'Error al subir: ' + err.toString();
    }
  }

  escribirEnRegalos(ss, [
    p.timestamp || new Date().toISOString(),
    'home',                   // origen
    p.nombre, p.meta, p.monto_usd, p.mensaje,
    p.archivo_nombre || '',
    driveUrl
  ]);
}

// ─────────────────────────────────────────────
//  MENSAJES DEL LIVESTREAM
//  Van a la misma hoja Regalos con origen=livestream
// ─────────────────────────────────────────────
function handleLivestreamMsg(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const meta = p.tipo === 'regalo'
    ? 'Regalo + Mensaje (Livestream)'
    : 'Mensaje (Livestream)';

  escribirEnRegalos(ss, [
    p.timestamp || new Date().toISOString(),
    'livestream',             // origen
    p.nombre, meta, p.monto_usd || '', p.mensaje,
    '',                       // sin comprobante
    ''
  ]);
}
