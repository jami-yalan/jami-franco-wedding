/**
 * ═══════════════════════════════════════════════════════════
 *  JAMI & FRANCO WEDDING — Google Apps Script FOTOS
 *  Pega este código en: script.google.com → tu proyecto de fotos
 * ═══════════════════════════════════════════════════════════
 *
 *  ESTE SCRIPT MANEJA:
 *    · Fotos subidas por los invitados en fotos.html
 *    · Galería pública en galeria.html (?action=gallery)
 *
 *  DESPUÉS DE PEGAR: Implementar → Nueva implementación
 *  (o "Gestionar implementaciones" → nueva versión si ya existe)
 * ═══════════════════════════════════════════════════════════
 */

const FOTOS_FOLDER_ID = '1ecS8s7oAI5sDMJHElQ5d5iARnxtCs3gn';

// ─────────────────────────────────────────────
//  POST — Recibe una foto y la guarda en Drive
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const payload  = JSON.parse(e.postData.contents);
    const nombre   = (payload.name     || 'Invitado').toString().trim();
    const fileName = (payload.fileName || 'archivo').toString().trim();
    const mimeType = (payload.mimeType || 'application/octet-stream').toString();
    const base64   = payload.data;

    if (!base64) {
      return output({ status: 'error', message: 'No se recibió archivo.' });
    }

    const rootFolder = DriveApp.getFolderById(FOTOS_FOLDER_ID);
    const ext = fileName.includes('.') ? '.' + fileName.split('.').pop() : '';

    let count = 0;
    const prefix = nombre + ' ';
    const allFiles = rootFolder.getFiles();
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      if (f.getName().startsWith(prefix)) count++;
    }

    const newFileName = nombre + ' ' + (count + 1) + ext;
    const bytes = Utilities.base64Decode(base64);
    const blob  = Utilities.newBlob(bytes, mimeType, newFileName);
    const file  = rootFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return output({ status: 'ok', fileId: file.getId(), fileName: newFileName });

  } catch (err) {
    return output({ status: 'error', message: err.toString() });
  }
}

// ─────────────────────────────────────────────
//  GET — ?action=gallery → lista de fotos para galería
//        sin parámetros  → ping de verificación
// ─────────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'gallery') {
    try {
      const folder = DriveApp.getFolderById(FOTOS_FOLDER_ID);
      const files  = folder.getFiles();
      const photos = [];

      while (files.hasNext()) {
        const file = files.next();
        const mime = file.getMimeType();
        if (mime && mime.startsWith('image/')) {
          const id = file.getId();
          const name = file.getName();
          // "María García 3.jpg" → "María García"
          const uploader = name.replace(/\s+\d+\.[^.]+$/, '').replace(/\.[^.]+$/, '') || 'Invitado';
          photos.push({
            id:       id,
            name:     name,
            uploader: uploader,
            thumb:    'https://drive.google.com/thumbnail?id=' + id + '&sz=w800',
            full:     'https://drive.google.com/uc?export=view&id=' + id,
            download: 'https://drive.google.com/uc?export=download&id=' + id
          });
        }
      }

      photos.reverse(); // más recientes primero

      return output({ status: 'ok', photos: photos, total: photos.length });

    } catch (err) {
      return output({ status: 'error', message: err.toString() });
    }
  }

  return output({ status: 'ok', message: 'JF Fotos script activo.' });
}

function output(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
