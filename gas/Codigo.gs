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
const SHEET_HORAS_CAMPO  = 'Horas Campo';          // horas objetivo al mes por subtipo (mensual)
const SHEET_COSTO        = 'Costo Personal';       // (legado) costo por categoria/quincena
const SHEET_PLANILLA     = 'Planilla';             // base detallada de planilla por persona/quincena
const SHEET_INVENTARIO   = 'Inventario';           // activos de la empresa
const SHEET_INV_HIST     = 'Inventario Historial'; // historial de movimientos de activos
const SHEET_LINEAS       = 'Lineas Celulares';     // líneas móviles / equipos (Entel)
const SHEET_SEG_BASE     = 'Seguros Base';         // base mensual de seguros (todos los meses)
const SHEET_SEG_EPS      = 'Seguros EPS Staff';    // relación EPS staff con costo de prima

// Encabezados de Inventario (extendido con columnas de laptop)
const INV_HEADERS = ['ID','Nombre','Categoria','Estado','Asignado','Fecha_Creacion','Obs',
                     'Cuenta','Marca','Procesador','RAM','ID_Dispositivo','Fecha_Compra','Mantenimiento','Mouse'];
const LINEAS_HEADERS  = ['Telefono','Modelo','SIM','IMEI','Plan','Estado','Inicio_Adenda','Fin_Adenda','Penalidad','Nombre','Posicion','Proyecto','Cuota'];
const SEG_BASE_HEADERS = ['Mes','Empresa','Aseguradora','Tipo_Seguro','Monto','Status'];
const SEG_EPS_HEADERS  = ['Mes','Seguro','Empresa','Contrato','Afiliado','Dependientes','Costo_Prima','Costo_Titular'];

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

  // Config Campo — horario de entrada y jornada (h/día) por subtipo de personal de campo
  const shCC = getOrCreate(ss, SHEET_CONFIG_CAMPO, ['Subtipo','Hora_Entrada','Horas_Jornada']);
  if (shCC.getLastRow() <= 1) {
    shCC.appendRow(['Merch',        '', '']);
    shCC.appendRow(['Promotor',     '', '']);
    shCC.appendRow(['Operaciones',  '', '']);
    shCC.appendRow(['Practicantes', '', '']);
  } else {
    // Asegura que exista la fila Practicantes aunque la pestaña ya tuviera datos
    const vals = shCC.getRange(1, 1, shCC.getLastRow(), 1).getValues().map(r => (r[0]||'').toString().trim().toLowerCase());
    if (vals.indexOf('practicantes') === -1) shCC.appendRow(['Practicantes', '', '']);
  }

  // Horas Campo — horas objetivo al mes por subtipo, editable mes a mes
  getOrCreate(ss, SHEET_HORAS_CAMPO, ['Anio','Mes','Subtipo','Horas']);

  getOrCreate(ss, SHEET_COSTO,      ['Anio','Mes','Quincena','Categoria','Monto']);
  getOrCreate(ss, SHEET_PLANILLA,   ['ID','Razon_Social','Tipo_Personal','Anio','Mes','Quincena','Cuenta','Proyecto','Gerente','Supervisor','Nombres','DNI','Ubigeo','Ciudad','Banco','Num_Cuenta','Num_CCI','Total_Pagar']);
  getOrCreate(ss, SHEET_INVENTARIO, INV_HEADERS);
  // Migración: agrega los encabezados de columnas de laptop si la pestaña ya tenía datos
  ensureHeaders(ss.getSheetByName(SHEET_INVENTARIO),
    { 8:'Cuenta', 9:'Marca', 10:'Procesador', 11:'RAM', 12:'ID_Dispositivo', 13:'Fecha_Compra', 14:'Mantenimiento', 15:'Mouse' });
  getOrCreate(ss, SHEET_INV_HIST,   ['Fecha','ID_Activo','Nombre','Estado','Asignado','Admin','Comentario']);

  // Gestión — Líneas celulares y Seguros
  getOrCreate(ss, SHEET_LINEAS,   LINEAS_HEADERS);
  getOrCreate(ss, SHEET_SEG_BASE, SEG_BASE_HEADERS);
  getOrCreate(ss, SHEET_SEG_EPS,  SEG_EPS_HEADERS);

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
    // Columnas: Subtipo(0), Hora_Entrada(1), Horas_Jornada(2)
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({
        subtipo:  r[0].toString().trim(),
        entrada:  formatTime(r[1]) || (r[1] !== undefined && r[1] !== '' ? r[1].toString().trim() : ''),
        jornada:  r[2] !== undefined && r[2] !== '' ? parseFloat(r[2]) || 0 : 0,
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET HORAS CAMPO (horas objetivo al mes por subtipo) ─────────
  if (accion === 'getHorasCampo') {
    const anio = p.anio || '';
    const sh = ss.getSheetByName(SHEET_HORAS_CAMPO);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: Anio(0), Mes(1), Subtipo(2), Horas(3)
    let rows = data.slice(1).filter(r => r[0] !== '' && r[2]);
    if (anio) rows = rows.filter(r => r[0].toString().trim() === anio.toString().trim());
    const result = rows.map(r => ({
      anio:    parseInt(r[0]) || 0,
      mes:     parseInt(r[1]) || 0,
      subtipo: r[2].toString().trim(),
      horas:   parseFloat(r[3]) || 0,
    }));
    return jsonResp({ rows: result, count: result.length });
  }

  // ── GET PLANILLA (base detallada de costo de personal) ──────────
  if (accion === 'getPlanilla') {
    const anio = p.anio || '';
    const sh   = ss.getSheetByName(SHEET_PLANILLA);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Columnas: ID(0),Razon_Social(1),Tipo_Personal(2),Anio(3),Mes(4),Quincena(5),Cuenta(6),
    //           Proyecto(7),Gerente(8),Supervisor(9),Nombres(10),DNI(11),Ubigeo(12),Ciudad(13),
    //           Banco(14),Num_Cuenta(15),Num_CCI(16),Total_Pagar(17)
    let rows = data.slice(1).filter(r => r[0]);
    if (anio) rows = rows.filter(r => r[3].toString().trim() === anio.toString().trim());
    const result = rows.map(r => ({
      id:         r[0].toString().trim(),
      razon:      r[1] ? r[1].toString().trim() : '',
      tipo:       r[2] ? r[2].toString().trim() : '',
      anio:       parseInt(r[3]) || 0,
      mes:        parseInt(r[4]) || 0,
      quincena:   r[5] ? r[5].toString().trim() : '',
      cuenta:     r[6] ? r[6].toString().trim() : '',
      proyecto:   r[7] ? r[7].toString().trim() : '',
      gerente:    r[8] ? r[8].toString().trim() : '',
      supervisor: r[9] ? r[9].toString().trim() : '',
      nombres:    r[10] ? r[10].toString().trim() : '',
      dni:        r[11] ? r[11].toString().trim() : '',
      ubigeo:     r[12] ? r[12].toString().trim() : '',
      ciudad:     r[13] ? r[13].toString().trim() : '',
      banco:      r[14] ? r[14].toString().trim() : '',
      numCuenta:  r[15] ? r[15].toString().trim() : '',
      numCCI:     r[16] ? r[16].toString().trim() : '',
      total:      parseFloat(r[17]) || 0,
    }));
    return jsonResp({ rows: result, count: result.length });
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
    // Columnas: ID(0),Nombre(1),Categoria(2),Estado(3),Asignado(4),Fecha_Creacion(5),Obs(6),
    //           Cuenta(7),Marca(8),Procesador(9),RAM(10),ID_Dispositivo(11),Fecha_Compra(12),Mantenimiento(13),Mouse(14)
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
        cuenta:    r[7]  ? r[7].toString().trim()  : '',
        marca:     r[8]  ? r[8].toString().trim()  : '',
        procesador:r[9]  ? r[9].toString().trim()  : '',
        ram:       r[10] ? r[10].toString().trim() : '',
        dispositivo:r[11]? r[11].toString().trim() : '',
        fecha_compra:r1ToString(r[12]),
        mantenimiento:r1ToString(r[13]),
        mouse:     r[14] ? r[14].toString().trim() : '',
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET LINEAS CELULARES ────────────────────────────────────────
  if (accion === 'getLineas') {
    const sh = ss.getSheetByName(SHEET_LINEAS);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Telefono(0),Modelo(1),SIM(2),IMEI(3),Plan(4),Estado(5),Inicio(6),Fin(7),Penalidad(8),Nombre(9),Posicion(10),Proyecto(11),Cuota(12)
    const rows = data.slice(1)
      .filter(r => r.some(c => c !== '' && c != null))
      .map((r, i) => ({
        rid:       String(i + 2),
        telefono:  r[0] != null ? r[0].toString().trim() : '',
        modelo:    r[1] ? r[1].toString().trim() : '',
        sim:       r[2] != null ? r[2].toString().trim() : '',
        imei:      r[3] != null ? r[3].toString().trim() : '',
        plan:      r[4] ? r[4].toString().trim() : '',
        estado:    r[5] ? r[5].toString().trim() : '',
        inicio:    r1ToString(r[6]),
        fin:       r1ToString(r[7]),
        penalidad: r[8] != null ? r[8].toString().trim() : '',
        nombre:    r[9] ? r[9].toString().trim() : '',
        posicion:  r[10] ? r[10].toString().trim() : '',
        proyecto:  r[11] ? r[11].toString().trim() : '',
        cuota:     r[12] ? r[12].toString().trim() : '',
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET SEGUROS BASE (todos los meses) ──────────────────────────
  if (accion === 'getSegurosBase') {
    const sh = ss.getSheetByName(SHEET_SEG_BASE);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Mes(0),Empresa(1),Aseguradora(2),Tipo_Seguro(3),Monto(4),Status(5)
    const rows = data.slice(1)
      .filter(r => r[0])
      .map((r, i) => ({
        rid:         String(i + 2),
        mes:         r[0] ? r[0].toString().trim() : '',
        empresa:     r[1] ? r[1].toString().trim() : '',
        aseguradora: r[2] ? r[2].toString().trim() : '',
        tipo:        r[3] ? r[3].toString().trim() : '',
        monto:       parseFloat(r[4]) || 0,
        status:      r[5] ? r[5].toString().trim() : '',
      }));
    return jsonResp({ rows, count: rows.length });
  }

  // ── GET SEGUROS EPS STAFF ───────────────────────────────────────
  if (accion === 'getSegurosEps') {
    const sh = ss.getSheetByName(SHEET_SEG_EPS);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Mes(0),Seguro(1),Empresa(2),Contrato(3),Afiliado(4),Dependientes(5),Costo_Prima(6),Costo_Titular(7)
    const rows = data.slice(1)
      .filter(r => r[4])
      .map((r, i) => ({
        rid:         String(i + 2),
        mes:         r[0] ? r[0].toString().trim() : '',
        seguro:      r[1] ? r[1].toString().trim() : '',
        empresa:     r[2] ? r[2].toString().trim() : '',
        contrato:    r[3] != null ? r[3].toString().trim() : '',
        afiliado:    r[4] ? r[4].toString().trim() : '',
        dependientes:r[5] != null && r[5] !== '' ? (parseInt(r[5]) || 0) : '',
        costo:       parseFloat(r[6]) || 0,
        costo_titular: r[7] != null && r[7] !== '' ? (parseFloat(r[7]) || 0) : '',
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
    const sh = getOrCreate(ss, SHEET_CONFIG_CAMPO, ['Subtipo','Hora_Entrada','Horas_Jornada']);
    const { subtipo, entrada, jornada } = body;
    if (!subtipo) return jsonResp({ error: 'subtipo requerido' });
    const jVal = (jornada !== undefined && jornada !== '') ? jornada : '';
    const data  = sh.getDataRange().getValues();
    let found   = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === subtipo.toString().trim().toLowerCase()) { found = i; break; }
    }
    if (found > 0) {
      sh.getRange(found + 1, 2, 1, 2).setValues([[entrada || '', jVal]]);
    } else {
      sh.appendRow([subtipo, entrada || '', jVal]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE HORAS CAMPO (horas objetivo al mes por subtipo) ────────
  if (accion === 'saveHorasCampo') {
    const sh = getOrCreate(ss, SHEET_HORAS_CAMPO, ['Anio','Mes','Subtipo','Horas']);
    const { anio, mes, subtipo, horas } = body;
    if (!anio || !mes || !subtipo) return jsonResp({ error: 'anio, mes y subtipo son requeridos' });
    const data = sh.getDataRange().getValues();
    let found  = -1;
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === parseInt(anio) &&
          parseInt(data[i][1]) === parseInt(mes) &&
          data[i][2].toString().trim().toLowerCase() === subtipo.toString().trim().toLowerCase()) { found = i; break; }
    }
    if (found > 0) {
      sh.getRange(found + 1, 4).setValue(parseFloat(horas) || 0);
    } else {
      sh.appendRow([parseInt(anio), parseInt(mes), subtipo, parseFloat(horas) || 0]);
    }
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  // ── SAVE PLANILLA (una fila; upsert por ID) ─────────────────────
  if (accion === 'savePlanilla') {
    const sh = getOrCreate(ss, SHEET_PLANILLA, ['ID','Razon_Social','Tipo_Personal','Anio','Mes','Quincena','Cuenta','Proyecto','Gerente','Supervisor','Nombres','DNI','Ubigeo','Ciudad','Banco','Num_Cuenta','Num_CCI','Total_Pagar']);
    let b = body;
    const data = sh.getDataRange().getValues();
    let id = b.id ? b.id.toString().trim() : '';
    let found = -1;
    if (id) { for (let i = 1; i < data.length; i++) { if (data[i][0].toString().trim() === id) { found = i; break; } } }
    if (!id) id = nextPlanillaId(data);
    const row = planillaRow(id, b);
    if (found > 0) sh.getRange(found + 1, 1, 1, 18).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created', id: id });
  }

  // ── SAVE PLANILLA BULK (carga masiva) ───────────────────────────
  if (accion === 'savePlanillaBulk') {
    const sh = getOrCreate(ss, SHEET_PLANILLA, ['ID','Razon_Social','Tipo_Personal','Anio','Mes','Quincena','Cuenta','Proyecto','Gerente','Supervisor','Nombres','DNI','Ubigeo','Ciudad','Banco','Num_Cuenta','Num_CCI','Total_Pagar']);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    const data = sh.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) { const n = parseInt((data[i][0]||'').toString().replace(/\D/g,''),10); if (!isNaN(n) && n>maxId) maxId=n; }
    const nuevas = items.map(b => planillaRow(String(++maxId).padStart(5,'0'), b));
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 18).setValues(nuevas);
    return jsonResp({ ok: true, count: nuevas.length });
  }

  // ── DELETE PLANILLA ─────────────────────────────────────────────
  if (accion === 'deletePlanilla') {
    const sh = ss.getSheetByName(SHEET_PLANILLA);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada' });
    const { id } = body;
    if (!id) return jsonResp({ error: 'id requerido' });
    const idStr = id.toString().trim();
    const data  = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === idStr) { sh.deleteRow(i + 1); return jsonResp({ ok: true, action: 'deleted' }); }
    }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
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
    const sh  = getOrCreate(ss, SHEET_INVENTARIO, INV_HEADERS);
    ensureHeaders(sh, { 8:'Cuenta', 9:'Marca', 10:'Procesador', 11:'RAM', 12:'ID_Dispositivo', 13:'Fecha_Compra', 14:'Mantenimiento', 15:'Mouse' });
    const shH = getOrCreate(ss, SHEET_INV_HIST,   ['Fecha','ID_Activo','Nombre','Estado','Asignado','Admin','Comentario']);
    let { id, nombre, categoria, estado, asignado, obs, admin, comentario,
          cuenta, marca, procesador, ram, dispositivo, fecha_compra, mantenimiento, mouse } = body;
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
      sh.getRange(found + 1, 1, 1, 15).setValues([[id, nombre, categoria || '', estado, asignado || '', creacion, obs || '',
        cuenta || '', marca || '', procesador || '', ram || '', dispositivo || '', fecha_compra || '', mantenimiento || '', mouse || '']]);
    } else {
      if (!id) id = nextInvId(data);
      sh.appendRow([id, nombre, categoria || '', estado, asignado || '', r1ToStringFull(new Date()), obs || '',
        cuenta || '', marca || '', procesador || '', ram || '', dispositivo || '', fecha_compra || '', mantenimiento || '', mouse || '']);
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

  // ── SAVE INVENTARIO BULK (carga masiva de activos) ──────────────
  if (accion === 'saveInventarioBulk') {
    const sh = getOrCreate(ss, SHEET_INVENTARIO, INV_HEADERS);
    ensureHeaders(sh, { 8:'Cuenta', 9:'Marca', 10:'Procesador', 11:'RAM', 12:'ID_Dispositivo', 13:'Fecha_Compra', 14:'Mantenimiento', 15:'Mouse' });
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    if (body.replace) clearBody(sh);
    const data = sh.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) { const n = parseInt((data[i][0]||'').toString().replace(/\D/g,''),10); if (!isNaN(n) && n>maxId) maxId=n; }
    const now = r1ToStringFull(new Date());
    const nuevas = items.map(b => [
      String(++maxId).padStart(3,'0'), b.nombre||'', b.categoria||'', b.estado||'Disponible', b.asignado||'', now, b.obs||'',
      b.cuenta||'', b.marca||'', b.procesador||'', b.ram||'', b.dispositivo||'', b.fecha_compra||'', b.mantenimiento||'', b.mouse||''
    ]);
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 15).setValues(nuevas);
    return jsonResp({ ok: true, count: nuevas.length });
  }

  // ── SAVE LINEA (una fila; upsert por fila rid) ──────────────────
  if (accion === 'saveLinea') {
    const sh = getOrCreate(ss, SHEET_LINEAS, LINEAS_HEADERS);
    const b  = body;
    const row = lineaRow(b);
    const rid = b.rid ? parseInt(b.rid, 10) : 0;
    if (rid && rid >= 2 && rid <= sh.getLastRow()) {
      sh.getRange(rid, 1, 1, LINEAS_HEADERS.length).setValues([row]);
      return jsonResp({ ok: true, action: 'updated', rid: String(rid) });
    }
    sh.appendRow(row);
    return jsonResp({ ok: true, action: 'created', rid: String(sh.getLastRow()) });
  }

  // ── SAVE LINEAS BULK ────────────────────────────────────────────
  if (accion === 'saveLineasBulk') {
    const sh = getOrCreate(ss, SHEET_LINEAS, LINEAS_HEADERS);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    if (body.replace) clearBody(sh);
    const nuevas = items.map(lineaRow);
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, LINEAS_HEADERS.length).setValues(nuevas);
    return jsonResp({ ok: true, count: nuevas.length });
  }

  // ── DELETE LINEA ────────────────────────────────────────────────
  if (accion === 'deleteLinea') {
    const sh = ss.getSheetByName(SHEET_LINEAS);
    if (!sh) return jsonResp({ error: 'Pestaña no encontrada' });
    const rid = parseInt(body.rid, 10);
    if (rid && rid >= 2 && rid <= sh.getLastRow()) { sh.deleteRow(rid); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'Fila no encontrada' });
  }

  // ── SAVE SEGUROS BASE BULK ──────────────────────────────────────
  if (accion === 'saveSegBaseBulk') {
    const sh = getOrCreate(ss, SHEET_SEG_BASE, SEG_BASE_HEADERS);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    if (body.replace) clearBody(sh);
    const nuevas = items.map(b => [ b.mes||'', b.empresa||'', b.aseguradora||'', b.tipo||'', parseFloat(b.monto)||0, b.status||'' ]);
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, SEG_BASE_HEADERS.length).setValues(nuevas);
    return jsonResp({ ok: true, count: nuevas.length });
  }

  // ── SAVE SEGUROS EPS BULK ───────────────────────────────────────
  if (accion === 'saveSegEpsBulk') {
    const sh = getOrCreate(ss, SHEET_SEG_EPS, SEG_EPS_HEADERS);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    if (body.replace) clearBody(sh);
    const nuevas = items.map(b => [ b.mes||'', b.seguro||'', b.empresa||'', b.contrato||'', b.afiliado||'',
      (b.dependientes===''||b.dependientes==null)?'':(parseInt(b.dependientes)||0),
      parseFloat(b.costo)||0,
      (b.costo_titular===''||b.costo_titular==null)?'':(parseFloat(b.costo_titular)||0) ]);
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, SEG_EPS_HEADERS.length).setValues(nuevas);
    return jsonResp({ ok: true, count: nuevas.length });
  }

  return jsonResp({ error: 'Acción no reconocida: ' + accion });
}

// ════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════
// Construye la fila de Planilla (18 columnas) desde el objeto del front
function planillaRow(id, b) {
  return [
    id,
    b.razon || '', b.tipo || '', parseInt(b.anio) || '', parseInt(b.mes) || '', b.quincena || '',
    b.cuenta || '', b.proyecto || '', b.gerente || '', b.supervisor || '', b.nombres || '',
    b.dni || '', b.ubigeo || '', b.ciudad || '', b.banco || '', b.numCuenta || '', b.numCCI || '',
    parseFloat(b.total) || 0,
  ];
}
// ID incremental de 5 dígitos para planilla
function nextPlanillaId(data) {
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt((data[i][0] || '').toString().replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return String(max + 1).padStart(5, '0');
}

// Construye la fila de Líneas Celulares (13 columnas) desde el objeto del front
function lineaRow(b) {
  return [
    b.telefono || '', b.modelo || '', b.sim || '', b.imei || '', b.plan || '', b.estado || '',
    b.inicio || '', b.fin || '', b.penalidad || '', b.nombre || '', b.posicion || '', b.proyecto || '', b.cuota || '',
  ];
}

// Borra todas las filas de datos (deja el encabezado) de una pestaña
function clearBody(sh) {
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
}

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
