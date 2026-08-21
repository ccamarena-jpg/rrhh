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
const SHEET_CONFIG_CAMPO = 'Config Campo';        // horario entrada + jornada por subtipo
const SHEET_COSTO        = 'Costo Personal';       // costo por categoria/quincena
const SHEET_INVENTARIO   = 'Inventario';           // activos de la empresa
const SHEET_INV_HIST     = 'Inventario Historial'; // historial de movimientos de activos

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

// ── Asegurar encabezados en columnas concretas (migración no destructiva) ──
// cols: objeto { numeroColumna(1-based): 'Texto encabezado' }
function ensureHeaders(sh, cols) {
  if (!sh) return;
  Object.keys(cols).forEach(function (c) {
    const col = parseInt(c, 10);
    const cell = sh.getRange(1, col);
    const actual = (cell.getValue() || '').toString().trim();
    if (!actual) {
      cell.setValue(cols[c])
          .setFontWeight('bold')
          .setBackground('#2c3280')
          .setFontColor('#ffffff');
    }
  });
}

// ── Inicializar todas las pestañas (ejecutar una sola vez) ─────────
function initSheets() {
  const ss = getSS();
  getOrCreate(ss, SHEET_PERSONAL,   ['DNI','Nombre','Ficha_Buk','Cargo','Area','Fecha_Ingreso','Tipo','Subtipo']);
  // Migración: si la pestaña de personal ya existía con datos, asegura los encabezados Tipo/Subtipo
  ensureHeaders(ss.getSheetByName(SHEET_PERSONAL), { 7: 'Tipo', 8: 'Subtipo' });
  getOrCreate(ss, SHEET_ASISTENCIA, ['Documento','Fecha','Ingreso','Salida','Proyecto','DIA']);
  getOrCreate(ss, SHEET_VACACIONES, ['Colaborador','Año','Mes','Nombre_Mes','Dias','Dias_Detalle']);
  getOrCreate(ss, SHEET_RESUMEN,    ['Mes','DNI','Nombre','Horas_DM','Horas_Vacaciones','Objetivo_Horas']);
  getOrCreate(ss, SHEET_TARDANZAS,  ['Mes','DNI','Nombre','Ficha_Buk','Monto_Descuento']);
  getOrCreate(ss, SHEET_CONFIG_VAC, ['Tipo','Clave','Valor']);

  // Config Campo — horario de entrada, jornada y horas/mes por subtipo de personal de campo
  const shCC = getOrCreate(ss, SHEET_CONFIG_CAMPO, ['Subtipo','Hora_Entrada','Horas_Jornada','Horas_Mes']);
  ensureHeaders(shCC, { 4: 'Horas_Mes' }); // migración si ya existía con 3 columnas
  if (shCC.getLastRow() <= 1) {
    shCC.appendRow(['Merch',        '', '', '']);
    shCC.appendRow(['Promotor',     '', '', '']);
    shCC.appendRow(['Operaciones',  '', '', '']);
    shCC.appendRow(['Practicantes', '', '', '']);
  } else {
    // Asegura que exista la fila Practicantes aunque la pestaña ya tuviera datos
    const vals = shCC.getRange(1, 1, shCC.getLastRow(), 1).getValues().map(r => (r[0]||'').toString().trim().toLowerCase());
    if (vals.indexOf('practicantes') === -1) shCC.appendRow(['Practicantes', '', '', '']);
  }

  getOrCreate(ss, SHEET_COSTO,      ['Anio','Mes','Quincena','Categoria','Monto']);
  getOrCreate(ss, SHEET_INVENTARIO, ['ID','Nombre','Categoria','Estado','Asignado','Fecha_Creacion','Obs']);
  getOrCreate(ss, SHEET_INV_HIST,   ['Fecha','ID_Activo','Nombre','Estado','Asignado','Admin','Comentario']);

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
        r[7] ? r[7].toString().trim() : '',  // Subtipo (Merch/Promotor/Operaciones)
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

  // ── GET CONFIG CAMPO (horarios/jornada por subtipo) ─────────────
  if (accion === 'getConfigCampo') {
    const sh = ss.getSheetByName(SHEET_CONFIG_CAMPO);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Subtipo(0), Hora_Entrada(1), Horas_Jornada(2), Horas_Mes(3)
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({
        subtipo:  r[0].toString().trim(),
        entrada:  formatTime(r[1]) || (r[1] !== undefined && r[1] !== '' ? r[1].toString().trim() : ''),
        jornada:  r[2] !== undefined && r[2] !== '' ? parseFloat(r[2]) || 0 : 0,
        horasMes: r[3] !== undefined && r[3] !== '' ? parseFloat(r[3]) || 0 : 0,
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET COSTO PERSONAL ──────────────────────────────────────────
  if (accion === 'getCostoPersonal') {
    const anio = p.anio || '';
    const sh   = ss.getSheetByName(SHEET_COSTO);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Anio(0), Mes(1), Quincena(2), Categoria(3), Monto(4)
    let rows = data.slice(1).filter(r => r[0] !== '' && r[3]);
    if (anio) rows = rows.filter(r => r[0].toString().trim() === anio.toString().trim());
    const result = rows.map(r => ({
      anio:      parseInt(r[0]) || 0,
      mes:       parseInt(r[1]) || 0,
      quincena:  r[2] ? r[2].toString().trim() : '',
      categoria: r[3].toString().trim(),
      monto:     parseFloat(r[4]) || 0,
    }));
    return jsonResp({ rows: result, count: result.length });
  }

  // ── GET INVENTARIO ──────────────────────────────────────────────
  if (accion === 'getInventario') {
    const sh = ss.getSheetByName(SHEET_INVENTARIO);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: ID(0), Nombre(1), Categoria(2), Estado(3), Asignado(4), Fecha_Creacion(5), Obs(6)
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({
        id:        r[0].toString().trim(),
        nombre:    r[1] ? r[1].toString().trim() : '',
        categoria: r[2] ? r[2].toString().trim() : '',
        estado:    r[3] ? r[3].toString().trim() : 'Disponible',
        asignado:  r[4] ? r[4].toString().trim() : '',
        creacion:  r1ToString(r[5]),
        obs:       r[6] ? r[6].toString().trim() : '',
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET INVENTARIO HISTORIAL ────────────────────────────────────
  if (accion === 'getInvHistorial') {
    const idFiltro = p.id || '';
    const sh = ss.getSheetByName(SHEET_INV_HIST);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Fecha(0), ID_Activo(1), Nombre(2), Estado(3), Asignado(4), Admin(5), Comentario(6)
    let rows = data.slice(1).filter(r => r[1]);
    if (idFiltro) rows = rows.filter(r => r[1].toString().trim() === idFiltro.toString().trim());
    const result = rows.map(r => ({
      fecha:      r1ToStringFull(r[0]),
      id_activo:  r[1].toString().trim(),
      nombre:     r[2] ? r[2].toString().trim() : '',
      estado:     r[3] ? r[3].toString().trim() : '',
      asignado:   r[4] ? r[4].toString().trim() : '',
      admin:      r[5] ? r[5].toString().trim() : '',
      comentario: r[6] ? r[6].toString().trim() : '',
    }));
    return jsonResp({ rows: result, count: result.length });
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
    const sh = getOrCreate(ss, SHEET_PERSONAL, ['DNI','Nombre','Ficha_Buk','Cargo','Area','Fecha_Ingreso','Tipo','Subtipo']);
    const { dni, nombre, ficha_buk, cargo, area, fecha_ingreso, tipo, subtipo } = body;
    if (!dni || !nombre) return jsonResp({ error: 'DNI y Nombre son requeridos' });
    const dniStr = dni.toString().trim();
    const data   = sh.getDataRange().getValues();
    let found    = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === dniStr) { found = i; break; }
    }
    const row = [dniStr, nombre.toString().trim(), ficha_buk||'', cargo||'', area||'', fecha_ingreso||'', tipo||'Staff', subtipo||''];
    if (found > 0) {
      sh.getRange(found + 1, 1, 1, 8).setValues([row]);
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

  // ── SAVE CONFIG CAMPO (horario/jornada por subtipo) ─────────────
  if (accion === 'saveConfigCampo') {
    const sh = getOrCreate(ss, SHEET_CONFIG_CAMPO, ['Subtipo','Hora_Entrada','Horas_Jornada','Horas_Mes']);
    const { subtipo, entrada, jornada, horasMes } = body;
    if (!subtipo) return jsonResp({ error: 'subtipo requerido' });
    const jVal = (jornada !== undefined && jornada !== '') ? jornada : '';
    const hmVal = (horasMes !== undefined && horasMes !== '') ? horasMes : '';
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === subtipo.toString().trim().toLowerCase()) { found = i; break; }
    }
    if (found > 0) {
      sh.getRange(found + 1, 2, 1, 3).setValues([[entrada || '', jVal, hmVal]]);
    } else {
      sh.appendRow([subtipo, entrada || '', jVal, hmVal]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE COSTO PERSONAL ─────────────────────────────────────────
  if (accion === 'saveCostoPersonal') {
    const sh = getOrCreate(ss, SHEET_COSTO, ['Anio','Mes','Quincena','Categoria','Monto']);
    const { anio, mes, quincena, categoria, monto } = body;
    if (!anio || !mes || !quincena || !categoria) return jsonResp({ error: 'anio, mes, quincena y categoria son requeridos' });
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === parseInt(anio) &&
          parseInt(data[i][1]) === parseInt(mes) &&
          data[i][2].toString().trim() === quincena.toString().trim() &&
          data[i][3].toString().trim() === categoria.toString().trim()) {
        found = i; break;
      }
    }
    if (found > 0) {
      sh.getRange(found + 1, 5).setValue(parseFloat(monto) || 0);
    } else {
      sh.appendRow([parseInt(anio), parseInt(mes), quincena, categoria, parseFloat(monto) || 0]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE INVENTARIO (activo) — registra historial si cambia estado/asignado ──
  if (accion === 'saveInventario') {
    const sh  = getOrCreate(ss, SHEET_INVENTARIO, ['ID','Nombre','Categoria','Estado','Asignado','Fecha_Creacion','Obs']);
    const shH = getOrCreate(ss, SHEET_INV_HIST,   ['Fecha','ID_Activo','Nombre','Estado','Asignado','Admin','Comentario']);
    let { id, nombre, categoria, estado, asignado, obs, admin, comentario } = body;
    if (!nombre) return jsonResp({ error: 'nombre requerido' });
    estado = estado || 'Disponible';
    const data = sh.getDataRange().getValues();
    let found  = -1;
    let prevEstado = '', prevAsignado = '';
    if (id) {
      const idStr = id.toString().trim();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === idStr) { found = i; prevEstado = (data[i][3]||'').toString().trim(); prevAsignado = (data[i][4]||'').toString().trim(); break; }
      }
    }
    if (found > 0) {
      // conservar fecha de creación existente
      const creacion = data[found][5] || '';
      sh.getRange(found + 1, 1, 1, 7).setValues([[id, nombre, categoria || '', estado, asignado || '', creacion, obs || '']]);
    } else {
      if (!id) id = nextInvId(data);
      sh.appendRow([id, nombre, categoria || '', estado, asignado || '', r1ToStringFull(new Date()), obs || '']);
    }
    // Historial: registra en alta o cuando cambia estado/asignado
    const cambio = (found <= 0) || (prevEstado !== estado) || (prevAsignado !== (asignado || ''));
    if (cambio) {
      shH.appendRow([r1ToStringFull(new Date()), id, nombre, estado, asignado || '', admin || '', comentario || (found <= 0 ? 'Alta de activo' : '')]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created', id: id });
  }

  // ── DELETE INVENTARIO ───────────────────────────────────────────
  if (accion === 'deleteInventario') {
    const sh = ss.getSheetByName(SHEET_INVENTARIO);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada' });
    const { id } = body;
    if (!id) return jsonResp({ error: 'id requerido' });
    const idStr = id.toString().trim();
    const data  = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === idStr) {
        sh.deleteRow(i + 1);
        return jsonResp({ ok: true, action: 'deleted' });
      }
    }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  return jsonResp({ error: 'Acción no reconocida: ' + accion });
}

// ════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════
// Genera un ID numérico incremental de 3 dígitos para inventario
function nextInvId(data) {
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt((data[i][0] || '').toString().replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return String(max + 1).padStart(3, '0');
}

// Fecha + hora legible (DD/MM/YYYY HH:MM)
function r1ToStringFull(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return pad(val.getDate()) + '/' + pad(val.getMonth() + 1) + '/' + val.getFullYear() +
           ' ' + pad(val.getHours()) + ':' + pad(val.getMinutes());
  }
  return val.toString().trim();
}
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
