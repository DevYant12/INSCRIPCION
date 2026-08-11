/**
 * BACKEND para el Panel de Inscripciones del Taller
 * -------------------------------------------------
 * Instrucciones:
 * 1. Ve a https://sheets.google.com y crea una hoja de cálculo nueva
 *    (o usa una que ya tengas para este taller).
 * 2. En el menú, ve a Extensiones > Apps Script.
 * 3. Borra el código de ejemplo que aparece (function myFunction(){...})
 *    y pega TODO el contenido de este archivo.
 * 4. Guarda el proyecto (icono de disquete o Ctrl+S). Ponle un nombre,
 *    por ejemplo "Backend Inscripciones".
 * 5. Arriba a la derecha, haz clic en "Implementar" > "Nueva implementación".
 * 6. Haz clic en el ícono de engranaje junto a "Seleccionar tipo" y elige
 *    "Aplicación web".
 * 7. Configura:
 *      - Descripción: (lo que quieras)
 *      - Ejecutar como: Yo (tu correo)
 *      - Quién tiene acceso: Cualquier usuario
 * 8. Haz clic en "Implementar". Google te pedirá autorizar permisos:
 *    acepta (es tu propio script, es seguro).
 * 9. Copia la URL que aparece bajo "URL de la aplicación web"
 *    (empieza con https://script.google.com/macros/s/.../exec).
 * 10. Pega esa URL en los 3 archivos HTML (index.html, invitaciones.html,
 *     escaner.html), en la constante SCRIPT_URL al inicio del <script>.
 *
 * IMPORTANTE: Si ya tenías este script implementado antes, esta versión
 * es compatible con tu hoja actual: solo agrega 2 columnas nuevas
 * (Estado e IngresoTS) la primera vez que se ejecute.
 *
 * Cada vez que cambies este código, tendrás que hacer
 * "Implementar" > "Gestionar implementaciones" > editar (ícono de lápiz)
 * > Nueva versión > Implementar, para que los cambios se apliquen
 * a la URL ya publicada.
 */

var SHEET_NAME = 'Inscritos';
var PRICE_PER_PERSON = 30;

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Fecha', 'Codigo', 'Nombre', 'Personas', 'Total', 'Estado', 'IngresoTS']);
  } else if (sheet.getLastColumn() < 7) {
    // Migración automática: agrega las columnas nuevas si vienen de una versión anterior
    sheet.getRange(1, 6).setValue('Estado');
    sheet.getRange(1, 7).setValue('IngresoTS');
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e.parameter && e.parameter.action;

  if (action === 'checkin') {
    return jsonOutput_(checkin_(e.parameter.codigo));
  }
  if (action === 'status') {
    return jsonOutput_(status_(e.parameter.codigo));
  }

  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var entries = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[2]) continue; // fila vacía
      entries.push({
        ts: row[0] instanceof Date ? row[0].getTime() : row[0],
        codigo: row[1],
        nombre: row[2],
        personas: row[3],
        total: row[4],
        estado: row[5] || 'Pendiente',
        ingresoTs: row[6] instanceof Date ? row[6].getTime() : (row[6] || null)
      });
    }
    return jsonOutput_({ ok: true, entries: entries });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

// Registra una nueva inscripción
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var nombre = (body.nombre || '').toString().trim();
    var personas = parseInt(body.personas, 10) || 1;

    if (!nombre) {
      return jsonOutput_({ ok: false, error: 'Falta el nombre' });
    }

    var sheet = getSheet_();
    var secuencia = sheet.getLastRow(); // fila 1 es encabezado
    var codigo = 'TC-' + Utilities.formatString('%03d', secuencia) + '-' +
      Math.random().toString(36).slice(2, 6).toUpperCase();
    var total = personas * PRICE_PER_PERSON;
    var ts = new Date();

    sheet.appendRow([ts, codigo, nombre, personas, total, 'Pendiente', '']);

    return jsonOutput_({
      ok: true,
      entry: {
        ts: ts.getTime(), codigo: codigo, nombre: nombre, personas: personas,
        total: total, estado: 'Pendiente'
      }
    });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  }
}

// Busca la fila (1-indexada) que corresponde a un código. -1 si no existe.
function findRowByCodigo_(sheet, codigo) {
  var data = sheet.getDataRange().getValues();
  var target = String(codigo || '').trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() === target) {
      return i + 1;
    }
  }
  return -1;
}

// Marca un código como usado. Usa un candado para que dos escaneos
// simultáneos del MISMO código no puedan pasar los dos a la vez.
function checkin_(codigo) {
  if (!codigo) {
    return { ok: false, error: 'Falta el código' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return { ok: false, error: 'Sistema ocupado, intenta de nuevo en un segundo' };
  }

  try {
    var sheet = getSheet_();
    var rowNum = findRowByCodigo_(sheet, codigo);

    if (rowNum === -1) {
      return { ok: false, error: 'Código no encontrado', notFound: true };
    }

    var row = sheet.getRange(rowNum, 1, 1, 7).getValues()[0];
    var estado = row[5];
    var entry = {
      codigo: row[1],
      nombre: row[2],
      personas: row[3],
      total: row[4]
    };

    if (estado === 'Ingresado') {
      return {
        ok: false,
        error: 'Este código ya fue usado',
        yaUsado: true,
        entry: entry,
        ingresoTs: row[6] instanceof Date ? row[6].getTime() : row[6]
      };
    }

    var now = new Date();
    sheet.getRange(rowNum, 6).setValue('Ingresado');
    sheet.getRange(rowNum, 7).setValue(now);

    return { ok: true, entry: entry, ingresoTs: now.getTime() };
  } finally {
    lock.releaseLock();
  }
}

function status_(codigo) {
  var sheet = getSheet_();
  var rowNum = findRowByCodigo_(sheet, codigo);
  if (rowNum === -1) {
    return { ok: false, error: 'Código no encontrado' };
  }
  var row = sheet.getRange(rowNum, 1, 1, 7).getValues()[0];
  return {
    ok: true,
    entry: {
      codigo: row[1], nombre: row[2], personas: row[3], total: row[4],
      estado: row[5] || 'Pendiente',
      ingresoTs: row[6] instanceof Date ? row[6].getTime() : (row[6] || null)
    }
  };
}
