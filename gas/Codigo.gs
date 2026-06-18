// ══════════════════════════════════════════════════════════════════
// TT Audit RRHH — Apps Script v4
// ──────────────────────────────────────────────────────────────────
// Pasos para desplegar:
//   1. Ve a Extensiones → Apps Script en Google Sheets
//   2. Reemplaza TODO el código con este archivo
//   3. Implementar → Nueva implementación (o "Gestionar implementaciones" para actualizar)
//      · Tipo:      Aplicación web
//      · Ejecutar:  Yo (tu cuenta)
//      · Acceso:    Cualquiera (Anyone)
//   4. Copia la URL y pégala en Configuración → URL Apps Script
//   5. Si es la primera vez, ejecuta initSheets() manualmente desde el editor
// ══════════════════════════════════════════════════════════════════

// ── Configuración ─────────────────────────────────────────────────
const SS_ID              = '1s7r7KjxTYO_PK2obc_yQ9pXI8l-mEl3vbmrrfIqCrio';
const SHEET_PERSONAL     = 'Datos generales personal';
const SHEET_ASISTENCIA   = 'Detalle de asistencia';
const SHEET_VACACIONES   = 'Vacaciones';
const SHEET_RESUMEN      = 'Resumen Mensual';
const SHEET_TARDANZAS    = 'Tardanzas Descuentos';
const SHEET_CONFIG_VAC   = 'Config Vacaciones';

// ── Obtener el Spreadsheet (activo o por ID) ───────────────────────
function getSS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {
    Logger.log('getActiveSpreadsheet falló, usando openById: ' + e.message);
  }
  return SpreadsheetApp.openById(SS_ID);
}

// ── Crear o abrir pestaña; agrega headers si está vacía ────────────
function getOrCreate(ss, name, headers) {
  if (!ss) ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#2c3280')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Inicializar todas las pestañas (ejecutar una sola vez) ─────────
function initSheets() {
  const ss = getSS();
  getOrCreate(ss, SHEET_PERSONAL,   ['DNI','Nombre','Ficha_Buk','Cargo','Area','Fecha_Ingreso','Tipo']);
  getOrCreate(ss, SHEET_ASISTENCIA, ['Documento','Fecha','Ingreso','Salida','Proyecto','DIA']);
  getOrCreate(ss, SHEET_VACACIONES, ['Colaborador','Año','Mes','Nombre_Mes','Dias','Dias_Detalle']);
  getOrCreate(ss, SHEET_RESUMEN,    ['Mes','DNI','Nombre','Horas_DM','Horas_Vacaciones','Objetivo_Horas']);
  getOrCreate(ss, SHEET_TARDANZAS,  ['Mes','DNI','Nombre','Ficha_Buk','Monto_Descuento']);
  getOrCreate(ss, SHEET_CONFIG_VAC, ['Tipo','Clave','Valor']);
  Logger.log('initSheets OK');
  return 'OK — pestañas creadas';
}

// ════════════════════════════════════════════════════════════════════
//  GET — Lectura de datos
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    return handleGet(e);
  } catch (err) {
    Logger.log('doGet error: ' + err.message + '\n' + err.stack);
    return jsonResp({ error: err.message });
  }
}

function handleGet(e) {
  const ss     = getSS();
  const p      = (e && e.parameter) ? e.parameter : {};
  const accion = p.accion || '';

  // Health check
  if (!accion) {
    return jsonResp({ status: 'TT Audit API v4 activa', sheet: ss.getName() });
  }

  // ── GET PERSONAL ────────────────────────────────────────────────
  if (accion === 'getPersonal') {
    const sh = ss.getSheetByName(SHEET_PERSONAL);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada: ' + SHEET_PERSONAL });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1)
      .filter(r => r[0] && r[1])
      .map(r => [
        r[0].toString().trim(),              // DNI
        r[1].toString().trim(),              // Nombre
        r[2] ? r[2].toString().trim() : '',  // Ficha_Buk
        r[3] ? r[3].toString().trim() : '',  // Cargo
        r[4] ? r[4].toString().trim() : '',  // Area
        r[5] ? r[5].toString().trim() : '',  // Fecha_Ingreso
        r[6] ? r[6].toString().trim() : 'Staff', // Tipo
      ]);
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET ASISTENCIA ──────────────────────────────────────────────
  if (accion === 'getAsistencia') {
    const mes = p.mes || '';
    const sh  = ss.getSheetByName(SHEET_ASISTENCIA);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada: ' + SHEET_ASISTENCIA });
    const data = sh.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r   = data[i];
      const dni = (r[0] || '').toString().trim();
      if (!dni) continue;

      let fechaStr = '';
      if (r[1] instanceof Date) {
        fechaStr = pad(r[1].getDate()) + '/' + pad(r[1].getMonth() + 1) + '/' + r[1].getFullYear();
      } else {
        fechaStr = r[1] ? r[1].toString().trim() : '';
      }
      if (!fechaStr) continue;

      if (mes) {
        const parts = fechaStr.split('/');
        if (parts.length === 3) {
          const rowMes = parts[2] + '-' + pad(parseInt(parts[1]));
          if (rowMes !== mes) continue;
        }
      }
      rows.push([
        dni,
        fechaStr,
        formatTime(r[2]),
        formatTime(r[3]),
        r[4] ? r[4].toString().trim() : '',
        r[5] ? r[5].toString().trim() : '',
      ]);
    }
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET VACACIONES ──────────────────────────────────────────────
  if (accion === 'getVacaciones') {
    const sh = ss.getSheetByName(SHEET_VACACIONES);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada: ' + SHEET_VACACIONES });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Colaborador(0), Año(1), Mes(2), Nombre_Mes(3), Dias(4), Dias_Detalle(5)
    const rows = data.slice(1)
      .filter(r => r[0] && r[1] !== '' && r[2] !== '')
      .map(r => ({
        colaborador: r[0].toString().trim(),
        año:         parseInt(r[1]) || 0,
        mes:         parseInt(r[2]) || 0,
        nombre_mes:  r[3] ? r[3].toString().trim() : '',
        dias:        parseInt(r[4]) || 0,   // BUG FIX: era r[3], correcto es r[4]
        detalle:     r[5] ? r[5].toString().trim() : '',
      }))
      .filter(r => r.colaborador && r.año);
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET RESUMEN MENSUAL ─────────────────────────────────────────
  if (accion === 'getResumen') {
    const mes = p.mes || '';
    const sh  = ss.getSheetByName(SHEET_RESUMEN);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada: ' + SHEET_RESUMEN });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Mes(0), DNI(1), Nombre(2), Horas_DM(3), Horas_Vacaciones(4), Objetivo_Horas(5)
    let rows = data.slice(1).filter(r => r[0] && r[1]);
    if (mes) rows = rows.filter(r => r[0].toString().trim() === mes);
    const result = rows.map(r => ({
      mesKey:    r[0].toString().trim(),
      dni:       r[1].toString().trim(),
      nombre:    r[2] ? r[2].toString().trim() : '',
      dm:        parseFloat(r[3]) || 0,
      vac:       parseFloat(r[4]) || 0,
      objetivo:  parseFloat(r[5]) || 0,
    }));
    return jsonResp({ rows: result, count: result.length });
  }

  // ── GET TARDANZAS ───────────────────────────────────────────────
  if (accion === 'getTardanzas') {
    const mes = p.mes || '';
    const sh  = ss.getSheetByName(SHEET_TARDANZAS);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada: ' + SHEET_TARDANZAS });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Mes(0), DNI(1), Nombre(2), Ficha_Buk(3), Monto_Descuento(4)
    let rows = data.slice(1).filter(r => r[0] && r[1]);
    if (mes) rows = rows.filter(r => r[0].toString().trim() === mes);
    const result = rows.map(r => ({
      mesKey:    r[0].toString().trim(),
      dni:       r[1].toString().trim(),
      nombre:    r[2] ? r[2].toString().trim() : '',
      ficha_buk: r[3] ? r[3].toString().trim() : '',
      monto:     parseFloat(r[4]) || 0,
    }));
    return jsonResp({ rows: result, count: result.length });
  }

  // ── GET CONFIG VACACIONES ───────────────────────────────────────
  if (accion === 'getConfigVac') {
    const sh = ss.getSheetByName(SHEET_CONFIG_VAC);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Tipo(0), Clave(1), Valor(2)
    const rows = data.slice(1)
      .filter(r => r[0] && r[1])
      .map(r => ({
        tipo:  r[0].toString().trim(),
        clave: r[1].toString().trim(),
        valor: r[2] !== undefined ? r[2].toString().trim() : '',
      }));
    return jsonResp({ rows, count: rows.length });
  }

  return jsonResp({ status: 'TT Audit API v4 activa', sheet: ss.getName() });
}

// ════════════════════════════════════════════════════════════════════
//  POST — Escritura de datos
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    return handlePost(e);
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResp({ error: err.message });
  }
}

function handlePost(e) {
  const ss = getSS();
  let body;
  try {
    body = JSON.parse(e && e.postData ? e.postData.contents : '{}');
  } catch (err) {
    return jsonResp({ error: 'JSON inválido: ' + err.message });
  }
  const accion = body.accion || '';

  // ── SAVE PERSONAL ───────────────────────────────────────────────
  if (accion === 'savePersonal') {
    const sh = getOrCreate(ss, SHEET_PERSONAL, ['DNI','Nombre','Ficha_Buk','Cargo','Area','Fecha_Ingreso']);
    const { dni, nombre, ficha_buk, cargo, area, fecha_ingreso, tipo } = body;
    if (!dni || !nombre) return jsonResp({ error: 'DNI y Nombre son requeridos' });
    const dniStr = dni.toString().trim();
    const data   = sh.getDataRange().getValues();
    let found    = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === dniStr) { found = i; break; }
    }
    const row = [dniStr, nombre.toString().trim(), ficha_buk||'', cargo||'', area||'', fecha_ingreso||'', tipo||'Staff'];
    if (found > 0) {
      sh.getRange(found + 1, 1, 1, 7).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── DELETE PERSONAL ─────────────────────────────────────────────
  if (accion === 'deletePersonal') {
    const sh = ss.getSheetByName(SHEET_PERSONAL);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada' });
    const { dni } = body;
    if (!dni) return jsonResp({ error: 'DNI requerido' });
    const dniStr = dni.toString().trim();
    const data   = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === dniStr) {
        sh.deleteRow(i + 1);
        return jsonResp({ ok: true, action: 'deleted' });
      }
    }
    return jsonResp({ ok: false, error: 'DNI no encontrado' });
  }

  // ── SAVE ASISTENCIA ─────────────────────────────────────────────
  if (accion === 'saveAsistencia') {
    const sh = getOrCreate(ss, SHEET_ASISTENCIA, ['Documento','Fecha','Ingreso','Salida','Proyecto','DIA']);
    const { documento, fecha, ingreso, salida, proyecto, dia } = body;
    if (!documento || !fecha) return jsonResp({ error: 'Documento y Fecha son requeridos' });
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      const rowFecha = r1ToString(data[i][1]);
      if (data[i][0].toString().trim() === documento.toString().trim() && rowFecha === fecha.toString().trim()) {
        found = i; break;
      }
    }
    const row = [documento, fecha, ingreso||'', salida||'', proyecto||'', dia||''];
    if (found > 0) {
      sh.getRange(found + 1, 1, 1, 6).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE RESUMEN MENSUAL ────────────────────────────────────────
  if (accion === 'saveResumen') {
    const sh = getOrCreate(ss, SHEET_RESUMEN, ['Mes','DNI','Nombre','Horas_DM','Horas_Vacaciones','Objetivo_Horas']);
    const { mesKey, dni, nombre, dm, vac, objetivo } = body;
    if (!mesKey || !dni) return jsonResp({ error: 'mesKey y DNI son requeridos' });
    const dniStr = dni.toString().trim();
    const data   = sh.getDataRange().getValues();
    let found    = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === mesKey && data[i][1].toString().trim() === dniStr) { found = i; break; }
    }
    if (found > 0) {
      sh.getRange(found + 1, 4, 1, 3).setValues([[dm||0, vac||0, objetivo||0]]);
    } else {
      sh.appendRow([mesKey, dniStr, nombre||'', dm||0, vac||0, objetivo||0]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE DESCUENTO TARDANZA ─────────────────────────────────────
  if (accion === 'saveDescuento') {
    const sh = getOrCreate(ss, SHEET_TARDANZAS, ['Mes','DNI','Nombre','Ficha_Buk','Monto_Descuento']);
    const { mesKey, dni, nombre, ficha_buk, monto } = body;
    if (!mesKey || !dni) return jsonResp({ error: 'mesKey y DNI son requeridos' });
    const dniStr = dni.toString().trim();
    const data   = sh.getDataRange().getValues();
    let found    = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === mesKey && data[i][1].toString().trim() === dniStr) { found = i; break; }
    }
    if (found > 0) {
      sh.getRange(found + 1, 5).setValue(monto || 0);
    } else {
      sh.appendRow([mesKey, dniStr, nombre||'', ficha_buk||'', monto||0]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE VACACION ───────────────────────────────────────────────
  if (accion === 'saveVacacion') {
    const sh = getOrCreate(ss, SHEET_VACACIONES, ['Colaborador','Año','Mes','Nombre_Mes','Dias','Dias_Detalle']);
    const { colaborador, año, mes, dias, detalle } = body;
    if (!colaborador || !año || !mes) return jsonResp({ error: 'Faltan campos requeridos: colaborador, año, mes' });
    const mesNombres = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio',
                        'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === colaborador.toString().trim() &&
          parseInt(data[i][1]) === parseInt(año) &&
          parseInt(data[i][2]) === parseInt(mes)) {
        found = i; break;
      }
    }
    if (found > 0) {
      sh.getRange(found + 1, 5).setValue(dias || 0);
      sh.getRange(found + 1, 6).setValue(detalle || '');
    } else {
      sh.appendRow([colaborador, parseInt(año), parseInt(mes),
                    mesNombres[parseInt(mes)] || '', dias || 0, detalle || '']);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE CONFIG VACACIONES ──────────────────────────────────────
  if (accion === 'saveConfigVac') {
    const sh = getOrCreate(ss, SHEET_CONFIG_VAC, ['Tipo','Clave','Valor']);
    const { tipo, clave, valor } = body;
    if (!tipo || !clave) return jsonResp({ error: 'tipo y clave son requeridos' });
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === tipo && data[i][1].toString().trim() === clave) {
        found = i; break;
      }
    }
    if (found > 0) {
      sh.getRange(found + 1, 3).setValue(valor !== undefined ? valor : '');
    } else {
      sh.appendRow([tipo, clave, valor !== undefined ? valor : '']);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  return jsonResp({ error: 'Acción no reconocida: ' + accion });
}

// ════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════
function formatTime(val) {
  if (!val || val === '-') return '';
  if (val instanceof Date) {
    return pad(val.getHours()) + ':' + pad(val.getMinutes()) + ':' + pad(val.getSeconds());
  }
  const s = val.toString().trim();
  if (!s || s === '-') return '';
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) return pad(parseInt(m[1])) + ':' + m[2] + ':' + (m[3] || '00');
  return '';
}

function r1ToString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return pad(val.getDate()) + '/' + pad(val.getMonth() + 1) + '/' + val.getFullYear();
  }
  return val.toString().trim();
}

function pad(n) { return String(n).padStart(2, '0'); }

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
