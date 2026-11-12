/**
 * Mis Vehículos — Backend Google Apps Script (v3)
 * Hoja de cálculo: "seguimiento_servicios_vehiculos"
 *   Hoja "Vehículos": encabezados en fila 2, datos desde fila 3
 *     A Patente | B Propietario | C Marca | D Modelo | E Año | F Color | G Notas | H Tipo | I Km actual | J Fecha km
 *   Hoja "Servicios": encabezados en fila 2, datos desde fila 3
 *     A Patente | B Propietario | C Tipo de Servicio | D Fecha Servicio | E Km en Servicio | F Próx. Fecha
 *     G Próx. Km | H Días Restantes (fórmula) | I Estado (fórmula) | J Notas | K Costo | L Taller
 *
 * Las columnas I-J (Vehículos) y K-L (Servicios) se crean solas si no existen.
 *
 * INSTALACIÓN
 *  1. Extensiones > Apps Script > reemplazar Code.gs por este archivo > Guardar.
 *  2. Ejecutar una vez la función  instalarAlertas  (pide permisos: aceptar).
 *     Crea un aviso diario por mail a las 8 hs con lo vencido / por vencer.
 *  3. Implementar > Gestionar implementaciones > editar (lápiz) > Versión: "Nueva versión" > Implementar.
 *     (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario). La URL /exec no cambia.
 *  4. (Opcional) Seguridad: Configuración del proyecto > Propiedades del script >
 *     agregar API_KEY = una clave cualquiera, y poner la misma clave en Ajustes de la app.
 */

const SH_VEH = 'Vehículos';
const SH_SVC = 'Servicios';
const FIRST_ROW = 2 + 1;           // fila 2 = encabezados
const VEH_HDR = ['Patente','Propietario','Marca','Modelo','Año','Color','Notas','Tipo','Km actual','Fecha km'];
const SVC_HDR = ['Patente','Propietario','Tipo de Servicio','Fecha Servicio','Km en Servicio','Próx. Fecha','Próx. Km','Días Restantes','Estado','Notas','Costo','Taller'];

const DIAS_AVISO_MAIL = 30;        // días de anticipación para el mail diario
const MARGEN_KM_MAIL  = 1000;      // km de anticipación
const MARGEN_HS_MAIL  = 20;        // horas de anticipación (embarcaciones / aeronaves)

// ─────────────────────────────────────────────────────────────
function doGet(e)  { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
  return handle_(p);
}

function handle_(p) {
  let out;
  try {
    checkKey_(p);
    out = route_(p);
  } catch (err) {
    out = { error: String(err && err.message ? err.message : err) };
  }
  const json = JSON.stringify(out);
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (key && p.key !== key) throw new Error('Clave de acceso inválida');
}

function route_(p) {
  const a = p.action;
  if (a === 'ping')   return { ok: true, version: 3, tz: tz_() };
  if (a === 'getAll') return getAll_();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSchema_();
    switch (a) {
      case 'addVehiculo':    return addVehiculo_(p);
      case 'updateVehiculo': return updateVehiculo_(p);
      case 'deleteVehiculo': return deleteVehiculo_(p);
      case 'updateKm':       return updateKm_(p);
      case 'addServicio':    return addServicio_(p);
      case 'updateServicio': return updateServicio_(p);
      case 'deleteServicio': return deleteServicio_(p);
    }
  } finally {
    lock.releaseLock();
  }
  throw new Error('Acción desconocida: ' + a);
}

// ── Helpers ──────────────────────────────────────────────────
function ss_()  { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_()  { return ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
function sh_(n) {
  const s = ss_().getSheetByName(n);
  if (!s) throw new Error('No existe la hoja "' + n + '"');
  return s;
}
function norm_(p) { return String(p || '').replace(/\s+/g, '').toUpperCase(); }
function fmtD_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.substring(0, 10) : '';
}
function toDate_(s) {
  s = String(s || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '';
  return Utilities.parseDate(s + ' 12:00', tz_(), 'yyyy-MM-dd HH:mm');
}
function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? '' : n;
}

function ensureSchema_() {
  const v = sh_(SH_VEH), s = sh_(SH_SVC);
  for (let i = 8; i < VEH_HDR.length; i++) {
    const c = v.getRange(2, i + 1);
    if (!c.getValue()) c.setValue(VEH_HDR[i]).setFontWeight('bold');
  }
  for (let i = 10; i < SVC_HDR.length; i++) {
    const c = s.getRange(2, i + 1);
    if (!c.getValue()) c.setValue(SVC_HDR[i]).setFontWeight('bold');
  }
}

function rows_(sheet, ncols) {
  const last = sheet.getLastRow();
  if (last < FIRST_ROW) return [];
  return sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, ncols).getValues();
}

/** Primera fila libre (columna A vacía) — respeta filas con fórmulas pre-cargadas */
function nextRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < FIRST_ROW) return FIRST_ROW;
  const col = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === '') return FIRST_ROW + i;
  return last + 1;
}

function findVehRow_(pat) {
  const data = rows_(sh_(SH_VEH), 1);
  for (let i = 0; i < data.length; i++) if (norm_(data[i][0]) === norm_(pat)) return FIRST_ROW + i;
  return -1;
}

// ── Lectura ──────────────────────────────────────────────────
function getAll_() {
  const vehiculos = rows_(sh_(SH_VEH), VEH_HDR.length)
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      patente: String(r[0]).trim(), propietario: String(r[1] || ''),
      marca: String(r[2] || ''), modelo: String(r[3] || ''),
      anio: r[4] ? String(Math.round(Number(r[4])) || r[4]) : '',
      color: String(r[5] || ''), notas: String(r[6] || ''),
      tipoVeh: String(r[7] || 'Automóvil'),
      kmActual: r[8] === '' ? '' : r[8], fechaKm: fmtD_(r[9])
    }));

  const servicios = [];
  rows_(sh_(SH_SVC), SVC_HDR.length).forEach((r, i) => {
    if (String(r[0]).trim() === '') return;
    servicios.push({
      idRow: FIRST_ROW + i,
      patente: String(r[0]).trim(), propietario: String(r[1] || ''),
      tipo: String(r[2] || ''), fecha: fmtD_(r[3]), km: r[4] === '' ? '' : r[4],
      proxFecha: fmtD_(r[5]), proxKm: r[6] === '' ? '' : r[6],
      notas: String(r[9] || ''), costo: r[10] === '' ? '' : r[10], taller: String(r[11] || '')
    });
  });
  return { ok: true, vehiculos, servicios, sheetUrl: ss_().getUrl(), ts: new Date().toISOString() };
}

// ── Vehículos ────────────────────────────────────────────────
function vehRow_(p) {
  return [norm_(p.patente), p.propietario || '', p.marca || '', p.modelo || '',
          num_(p.anio), p.color || '', p.notas || '', p.tipoVeh || 'Automóvil',
          num_(p.kmActual), toDate_(p.fechaKm)];
}
function addVehiculo_(p) {
  if (!p.patente) throw new Error('Falta la patente');
  if (findVehRow_(p.patente) > 0) throw new Error('Ya existe ese vehículo');
  const sh = sh_(SH_VEH);
  sh.getRange(nextRow_(sh), 1, 1, VEH_HDR.length).setValues([vehRow_(p)]);
  return { ok: true };
}
function updateVehiculo_(p) {
  const row = findVehRow_(p.patente);
  if (row < 0) throw new Error('No se encontró el vehículo');
  const sh = sh_(SH_VEH);
  const cur = sh.getRange(row, 1, 1, VEH_HDR.length).getValues()[0];
  const nv = vehRow_(p);
  if (p.kmActual === undefined) { nv[8] = cur[8]; nv[9] = cur[9]; }
  sh.getRange(row, 1, 1, VEH_HDR.length).setValues([nv]);
  // propaga propietario a servicios
  const ss = sh_(SH_SVC), data = rows_(ss, 2);
  data.forEach((r, i) => { if (norm_(r[0]) === norm_(p.patente) && r[1] !== nv[1]) ss.getRange(FIRST_ROW + i, 2).setValue(nv[1]); });
  return { ok: true };
}
function updateKm_(p) {
  const row = findVehRow_(p.patente);
  if (row < 0) throw new Error('No se encontró el vehículo');
  sh_(SH_VEH).getRange(row, 9, 1, 2).setValues([[num_(p.km), toDate_(p.fecha) || new Date()]]);
  return { ok: true };
}
function deleteVehiculo_(p) {
  const row = findVehRow_(p.patente);
  if (row < 0) throw new Error('No se encontró el vehículo');
  const ss = sh_(SH_SVC), data = rows_(ss, 1);
  for (let i = data.length - 1; i >= 0; i--) if (norm_(data[i][0]) === norm_(p.patente)) ss.deleteRow(FIRST_ROW + i);
  sh_(SH_VEH).deleteRow(row);
  return { ok: true };
}

// ── Servicios ────────────────────────────────────────────────
function svcVals_(p) {
  return [norm_(p.patente), p.propietario || '', p.tipo || '', toDate_(p.fecha), num_(p.km),
          toDate_(p.proxFecha), num_(p.proxKm)];
}
function writeFormulas_(sh, row) {
  sh.getRange(row, 8, 1, 2).setFormulas([[
    '=IF(F' + row + '="", "", INT(F' + row + ')-INT(TODAY()))',
    '=IF(H' + row + '="", "Sin fecha", IF(H' + row + '<0, "VENCIDO", IF(H' + row + '<=30, "⚠️ Próximo", "✅ Al día")))'
  ]]);
}
function addServicio_(p) {
  if (!p.patente || !p.tipo || !p.fecha) throw new Error('Faltan datos obligatorios');
  const sh = sh_(SH_SVC), row = nextRow_(sh);
  sh.getRange(row, 1, 1, 7).setValues([svcVals_(p)]);
  writeFormulas_(sh, row);
  sh.getRange(row, 10, 1, 3).setValues([[p.notas || '', num_(p.costo), p.taller || '']]);
  autoKm_(p);
  return { ok: true, idRow: row };
}
function updateServicio_(p) {
  const sh = sh_(SH_SVC), row = Number(p.idRow);
  if (!row || row < FIRST_ROW || row > sh.getLastRow()) throw new Error('Servicio inválido');
  const cur = norm_(sh.getRange(row, 1).getValue());
  if (p.origPatente && cur !== norm_(p.origPatente)) throw new Error('La planilla cambió, actualizá y reintentá');
  sh.getRange(row, 1, 1, 7).setValues([svcVals_(p)]);
  writeFormulas_(sh, row);
  sh.getRange(row, 10, 1, 3).setValues([[p.notas || '', num_(p.costo), p.taller || '']]);
  autoKm_(p);
  return { ok: true };
}
function deleteServicio_(p) {
  const sh = sh_(SH_SVC), row = Number(p.idRow);
  if (!row || row < FIRST_ROW || row > sh.getLastRow()) throw new Error('Servicio inválido');
  if (p.patente && norm_(sh.getRange(row, 1).getValue()) !== norm_(p.patente)) {
    throw new Error('La planilla cambió, actualizá y reintentá');
  }
  sh.deleteRow(row);
  return { ok: true };
}
/** Si el km del servicio es mayor al km actual registrado, lo actualiza */
function autoKm_(p) {
  const km = num_(p.km); if (km === '') return;
  const row = findVehRow_(p.patente); if (row < 0) return;
  const sh = sh_(SH_VEH), cur = num_(sh.getRange(row, 9).getValue());
  if (cur === '' || km > cur) sh.getRange(row, 9, 1, 2).setValues([[km, toDate_(p.fecha) || new Date()]]);
}

// ── Alertas por mail ─────────────────────────────────────────
function instalarAlertas() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'revisarVencimientos')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisarVencimientos').timeBased().everyDays(1).atHour(8).create();
  revisarVencimientos();
}

function revisarVencimientos() {
  const data = getAll_();
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const kmAct = {};
  const tipos = {};
  data.vehiculos.forEach(v => { kmAct[norm_(v.patente)] = num_(v.kmActual) || 0; tipos[norm_(v.patente)] = v.tipoVeh; });
  data.servicios.forEach(s => { const k = num_(s.km) || 0; const p = norm_(s.patente); if (k > (kmAct[p] || 0)) kmAct[p] = k; });

  // un servicio sigue "activo" si es el más reciente para al menos uno de sus ítems
  const ult = {};
  data.servicios.forEach(s => {
    s.tipo.split('+').map(x => x.trim()).filter(String).forEach(c => {
      const key = norm_(s.patente) + '|' + c.toLowerCase();
      if (!ult[key] || s.fecha > ult[key].fecha || (s.fecha === ult[key].fecha && s.idRow > ult[key].idRow)) ult[key] = s;
    });
  });
  const activos = {};
  Object.keys(ult).forEach(k => activos[ult[k].idRow] = ult[k]);

  const items = [];
  Object.keys(activos).forEach(k => {
    const s = activos[k], p = norm_(s.patente);
    const esHs = tipos[p] === 'Embarcación' || tipos[p] === 'Aeronave';
    let dias = null, rest = null;
    if (s.proxFecha) dias = Math.round((toDate_(s.proxFecha) - hoy) / 86400000);
    if (num_(s.proxKm) !== '' && Number(s.proxKm) > 0 && kmAct[p]) rest = Number(s.proxKm) - kmAct[p];
    const venc = (dias !== null && dias < 0) || (rest !== null && rest <= 0);
    const prox = (dias !== null && dias <= DIAS_AVISO_MAIL) || (rest !== null && rest <= (esHs ? MARGEN_HS_MAIL : MARGEN_KM_MAIL));
    if (venc || prox) items.push({ s, dias, rest, venc, u: esHs ? 'hs' : 'km' });
  });
  if (!items.length) return;

  items.sort((a, b) => (b.venc - a.venc) || ((a.dias ?? 9999) - (b.dias ?? 9999)));
  const filas = items.map(it => {
    const col = it.venc ? '#C0392B' : '#856404';
    const det = [];
    if (it.dias !== null) det.push(it.dias < 0 ? 'vencido hace ' + (-it.dias) + ' días' : 'faltan ' + it.dias + ' días');
    if (it.rest !== null) det.push(it.rest <= 0 ? 'excedido ' + (-it.rest).toLocaleString('es-AR') + ' ' + it.u : 'faltan ' + it.rest.toLocaleString('es-AR') + ' ' + it.u);
    return '<tr><td style="padding:6px;border-bottom:1px solid #eee"><b>' + it.s.patente + '</b><br><small>' + it.s.propietario + '</small></td>' +
      '<td style="padding:6px;border-bottom:1px solid #eee">' + it.s.tipo + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #eee;color:' + col + ';font-weight:bold">' + det.join(' · ') + '</td></tr>';
  }).join('');
  const nV = items.filter(i => i.venc).length;
  const asunto = '🔧 Mis Vehículos: ' + (nV ? nV + ' vencido(s), ' : '') + (items.length - nV) + ' próximo(s)';
  const html = '<div style="font-family:Arial,sans-serif;max-width:600px"><h2 style="color:#1A6B3C">Mantenimientos a atender</h2>' +
    '<table style="border-collapse:collapse;width:100%;font-size:13px"><tr style="background:#D6F0E2"><th align="left" style="padding:6px">Vehículo</th><th align="left" style="padding:6px">Servicio</th><th align="left" style="padding:6px">Estado</th></tr>' +
    filas + '</table><p style="color:#888;font-size:11px">Aviso automático diario · ' + ss_().getName() + '</p></div>';
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: asunto, htmlBody: html });
}
