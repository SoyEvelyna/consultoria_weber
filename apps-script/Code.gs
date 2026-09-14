/**
 * Weber Tracker — backend de Google Apps Script.
 *
 * Se pega en un proyecto INDEPENDIENTE de Apps Script (script.google.com),
 * creado desde una cuenta Gmail personal con permiso de edición sobre el
 * Sheet, y se publica como Web App. La web del tracker le habla por HTTP.
 * Va en una cuenta personal porque el Workspace de soyevelyna.com bloquea
 * la publicación de Web Apps por política de la organización.
 *
 * DISEÑO:
 * - "01 I Plan de trabajo" y "02 I Proceso de trabajo" son la base: la web
 *   las lee en cada pedido y ESCRIBE ahí mismo. Las tareas nuevas o editadas
 *   van a su fila de 02 (INICIATIVAS, o FINALIZADOS al completarse) y las
 *   reuniones se agregan a la tabla de 01. Se puede seguir editando a mano.
 * - Lo que no tiene columna en 01/02 (link de la tarea, notas) se guarda en
 *   "WebApp - Overrides" y "WebApp - Notas". "WebApp - Tareas nuevas" y
 *   "WebApp - Reuniones" quedan de la versión anterior.
 * - Autenticación: un token compartido (ver ACCESS_TOKEN abajo), el mismo
 *   que se configura en la web. No hay login individual.
 */

var ACCESS_TOKEN = "OKSiDEIeEQ55k3kCXS3cOL0f53TNBR9U";

/* Id del Google Sheet "Loopa I Weber I Plan de trabajo".
   El script es independiente (no está pegado al Sheet), así que lo abre
   por id. La cuenta que publica el Web App necesita permiso de edición
   sobre el Sheet — abrahanevelyn@gmail.com ya lo tiene. */
var SHEET_ID = "1rD90bzuYCm4HBJNMYvJDe_0dEaY25ED4zIL4xYKDhd0";
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

var SHEET_ETAPA1 = "01 I Plan de trabajo";
var SHEET_PROCESO = "02 I Proceso de trabajo";
var SHEET_OVERRIDES = "WebApp - Overrides";
var SHEET_NOTES = "WebApp - Notas";
var SHEET_CUSTOM_TASKS = "WebApp - Tareas nuevas";
var SHEET_MEETINGS = "WebApp - Reuniones";

/**
 * EJECUTAR A MANO UNA SOLA VEZ.
 * Crea las 4 pestañas que usa la web (si no existen) con sus encabezados.
 * En el editor: elegir "crearPestanasWebApp" en el desplegable y Ejecutar.
 * No toca "01 I Plan de trabajo" ni "02 I Proceso de trabajo".
 */
function crearPestanasWebApp() {
  ensureSheets_();
  var nombres = Object.keys(SHEET_SCHEMAS).join(", ");
  Logger.log("Listo. Pestañas verificadas/creadas: " + nombres);
}

/* =====================================================================
   ENTRY POINTS
   ===================================================================== */

function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    if (e.parameter.action !== "read") {
      return jsonOut_({ ok: false, error: "acción GET no soportada: " + e.parameter.action });
    }
    ensureSheets_();
    return jsonOut_({
      ok: true,
      seed: readSeed_(),
      overrides: readOverrides_(),
      notes: readSimpleRows_(SHEET_NOTES, ["id", "text", "author", "createdAt"]),
      customTasks: readSimpleRows_(SHEET_CUSTOM_TASKS,
        ["id", "tarea", "area", "tema", "responsable", "fase", "inicio", "cierre", "estado", "obs", "link", "createdAt"],
        ["inicio", "cierre"]),
      meetings: [] // las reuniones ahora viven en la tabla de "01" (seed.etapa1)
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents || "{}");
    checkToken_(body.token);
    ensureSheets_();
    var result = handleAction_(body.action, body.payload || {});
    return jsonOut_({ ok: true, result: result });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handleAction_(action, p) {
  switch (action) {
    case "setOverride": return setOverride_(p.taskId, p.patch || {});
    case "addNote": return addRow_(SHEET_NOTES, ["id", "text", "author", "createdAt"],
      Object.assign({ id: "n" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteNote": return deleteRow_(SHEET_NOTES, p.id);
    case "addCustomTask": return addRow_(SHEET_CUSTOM_TASKS,
      ["id", "tarea", "area", "tema", "responsable", "fase", "inicio", "cierre", "estado", "obs", "link", "createdAt"],
      Object.assign({ id: "c" + Date.now(), createdAt: nowIso_() }, p));
    case "updateCustomTask": return updateRow_(SHEET_CUSTOM_TASKS, p.id, p.patch || {});
    case "deleteCustomTask": return deleteRow_(SHEET_CUSTOM_TASKS, p.id);
    case "addMeeting": return addRow_(SHEET_MEETINGS, ["id", "fecha", "responsable", "duracion", "resumen", "createdAt"],
      Object.assign({ id: "m" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteMeeting": return deleteRow_(SHEET_MEETINGS, p.id);
    case "updateTask": return updateTask_(p.id, p.fields || {});
    case "addTask": return addTask_(p.fields || {});
    case "deleteTask": return deleteTask_(p.id);
    case "addMeetingSheet": return addMeetingSheet_(p);
    case "deleteMeetingSheet": return deleteMeetingSheet_(p);
    case "migrarWebApp": return migrarWebAppAlSheet_();
    default: throw new Error("acción POST no soportada: " + action);
  }
}

/* =====================================================================
   AUTH / UTILS
   ===================================================================== */

function checkToken_(token) {
  if (!token || token !== ACCESS_TOKEN) throw new Error("token inválido");
}
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function nowIso_() { return new Date().toISOString(); }

function toIsoDate_(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(value).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var d = m[1], mo = m[2], y = m[3];
    if (y.length === 2) y = "20" + y;
    return y + "-" + ("0" + mo).slice(-2) + "-" + ("0" + d).slice(-2);
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s; // no reconocido: se deja como texto para no perder el dato
}

function cell_(row, idx) {
  var v = row[idx];
  if (v === null || v === undefined) return null;
  if (typeof v === "string") { v = v.trim(); return v === "" ? null : v; }
  return v;
}

/* =====================================================================
   LECTURA DE "01 I Plan de trabajo" Y "02 I Proceso de trabajo"
   (nunca se escriben — solo lectura, respetando su estructura actual)
   ===================================================================== */

function readSeed_() {
  var proceso = readProceso_();
  proceso.etapa1 = readEtapa1_();
  return proceso;
}

function readEtapa1_() {
  var sheet = ss_().getSheetByName(SHEET_ETAPA1);
  if (!sheet) throw new Error("No encuentro la hoja '" + SHEET_ETAPA1 + "'");
  var values = sheet.getDataRange().getValues();
  var headerRow = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === "Fecha") { headerRow = i; break; }
  }
  if (headerRow === -1) throw new Error("No encuentro la fila de encabezado ('Fecha') en '" + SHEET_ETAPA1 + "'");
  var out = [];
  for (var r = headerRow + 1; r < values.length; r++) {
    var row = values[r];
    var fecha = cell_(row, 0);
    if (!fecha) break; // fin de la tabla
    out.push({
      fecha: toIsoDate_(fecha),
      hs: cell_(row, 1),
      tarea: cell_(row, 2),
      responsable: cell_(row, 3),
      estado: cell_(row, 4),
      resultado: cell_(row, 5),
      obs: cell_(row, 6)
    });
  }
  return out;
}

function readProceso_() {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  if (!sheet) throw new Error("No encuentro la hoja '" + SHEET_PROCESO + "'");
  var values = sheet.getDataRange().getValues();

  function findRow(colIdx, text) {
    var target = text.toUpperCase();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][colIdx] || "").trim().toUpperCase() === target) return i;
    }
    return -1;
  }

  var objetivoRow = findRow(1, "OBJETIVO 1");
  var prioridadesRow = findRow(1, "PRIORIDADES");
  var iniciativasRow = findRow(1, "INICIATIVAS");
  var finalizadosRow = findRow(1, "FINALIZADOS");
  if ([objetivoRow, prioridadesRow, iniciativasRow, finalizadosRow].indexOf(-1) !== -1) {
    throw new Error("No encuentro alguna de las secciones (OBJETIVO 1 / PRIORIDADES / INICIATIVAS / FINALIZADOS) en '" + SHEET_PROCESO + "'. ¿Se movieron o renombraron?");
  }
  // PENDIENTES es opcional: se eliminó del Sheet, pero si vuelve se lee igual.
  var pendientesRow = findRow(1, "PENDIENTES");
  var finIniciativas = pendientesRow !== -1 ? pendientesRow : finalizadosRow;

  // Objetivo: 2 filas después de "OBJETIVO 1" (la pregunta, y la respuesta)
  var objetivo = cell_(values[objetivoRow + 2], 1) || "";

  // Prioridades: 3 filas después de "PRIORIDADES" (la pregunta), cada una: [n, "titulo\t...\tdesc"]
  var prioridades = [];
  for (var pr = prioridadesRow + 2; pr < iniciativasRow; pr++) {
    var n = cell_(values[pr], 1);
    var raw = cell_(values[pr], 2);
    if (n === null || raw === null) continue;
    var parts = String(raw).split(/\t+/).map(function (s) { return s.trim(); }).filter(Boolean);
    prioridades.push({ n: n, tt: parts[0] || "", desc: parts.slice(1).join(" ") });
  }

  // Iniciativas: encabezado (ÁREA..OBSERVACIONES) en columna C, 2 filas después de "INICIATIVAS"
  var iniciativasHeaderRow = iniciativasRow + 2;
  var iniciativas = [];
  for (var ir = iniciativasHeaderRow + 1; ir < finIniciativas; ir++) {
    var row = values[ir];
    var area = cell_(row, 2), tarea = cell_(row, 4);
    if (!area && !tarea) continue;
    iniciativas.push({
      area: area, tema: cell_(row, 3), tarea: tarea, responsable: cell_(row, 5),
      inicio: toIsoDate_(cell_(row, 6)), tiempo: toIsoDate_(cell_(row, 7)), cierre: toIsoDate_(cell_(row, 8)),
      estado: cell_(row, 9), obs: cell_(row, 10)
    });
  }

  // Pendientes / backlog: solo texto libre en columna E (vacío si no está la sección)
  var backlog = [];
  if (pendientesRow !== -1) {
    for (var brow = pendientesRow + 1; brow < finalizadosRow; brow++) {
      var txt = cell_(values[brow], 4);
      if (txt) backlog.push(txt);
    }
  }

  // Finalizados: misma estructura de columnas que iniciativas, hasta el final de la hoja
  var finalizados = [];
  for (var fr = finalizadosRow + 1; fr < values.length; fr++) {
    var frow = values[fr];
    var farea = cell_(frow, 2), ftarea = cell_(frow, 4);
    if (!farea && !ftarea) continue;
    finalizados.push({
      area: farea, tema: cell_(frow, 3), tarea: ftarea, responsable: cell_(frow, 5),
      inicio: toIsoDate_(cell_(frow, 6)), tiempo: toIsoDate_(cell_(frow, 7)), cierre: toIsoDate_(cell_(frow, 8)),
      estado: cell_(frow, 9), obs: cell_(frow, 10)
    });
  }

  return {
    objetivo: objetivo,
    prioridades: prioridades,
    iniciativas: iniciativas,
    finalizados: finalizados,
    backlog: backlog
  };
}

/* =====================================================================
   PESTAÑAS DE LA WEB (WebApp - *) — creadas si no existen
   ===================================================================== */

var SHEET_SCHEMAS = {};
SHEET_SCHEMAS[SHEET_OVERRIDES] = ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"];
SHEET_SCHEMAS[SHEET_NOTES] = ["id", "text", "author", "createdAt"];
SHEET_SCHEMAS[SHEET_CUSTOM_TASKS] = ["id", "tarea", "area", "tema", "responsable", "fase", "inicio", "cierre", "estado", "obs", "link", "createdAt"];
SHEET_SCHEMAS[SHEET_MEETINGS] = ["id", "fecha", "responsable", "duracion", "resumen", "createdAt"];

function ensureSheets_() {
  var ss = ss_();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
      sheet.setFrozenRows(1);
    }
  });
}

/* dateCols: columnas que Sheets puede convertir sola a tipo fecha. Se
   normalizan a "yyyy-MM-dd", que es lo que espera la web. */
function readSimpleRows_(sheetName, cols, dateCols) {
  var sheet = ss_().getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!cell_(row, 0)) continue;
    var obj = {};
    cols.forEach(function (c, i) {
      var v = cell_(row, i);
      if (dateCols && dateCols.indexOf(c) !== -1) v = toIsoDate_(v);
      obj[c] = v;
    });
    out.push(obj);
  }
  return out;
}

function readOverrides_() {
  var rows = readSimpleRows_(SHEET_OVERRIDES, ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"],
    ["inicio", "cierre"]);
  var out = {};
  rows.forEach(function (r) {
    out[r.task_id] = { estado: r.estado, link: r.link, inicio: r.inicio, cierre: r.cierre, hidden: r.hidden === true || r.hidden === "true", updatedAt: r.updated_at };
  });
  return out;
}

function findRowIndexById_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) return r + 1; // 1-based para getRange
  }
  return -1;
}

function addRow_(sheetName, cols, obj) {
  var sheet = ss_().getSheetByName(sheetName);
  sheet.appendRow(cols.map(function (c) { return obj[c] === undefined ? "" : obj[c]; }));
  return obj;
}

function updateRow_(sheetName, id, patch) {
  var sheet = ss_().getSheetByName(sheetName);
  var cols = SHEET_SCHEMAS[sheetName];
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx === -1) throw new Error("No encuentro id " + id + " en " + sheetName);
  Object.keys(patch).forEach(function (key) {
    var colIdx = cols.indexOf(key);
    if (colIdx === -1) return;
    sheet.getRange(rowIdx, colIdx + 1).setValue(patch[key] === null ? "" : patch[key]);
  });
  return { id: id, patch: patch };
}

function deleteRow_(sheetName, id) {
  var sheet = ss_().getSheetByName(sheetName);
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx !== -1) sheet.deleteRow(rowIdx);
  return { id: id, deleted: rowIdx !== -1 };
}

function setOverride_(taskId, patch) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var rowIdx = findRowIndexById_(sheet, taskId);
  var updatedAt = nowIso_();
  if (rowIdx === -1) {
    var row = { task_id: taskId, estado: "", link: "", inicio: "", cierre: "", hidden: false, updated_at: updatedAt };
    Object.keys(patch).forEach(function (k) { row[k] = patch[k] === null ? "" : patch[k]; });
    addRow_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES], row);
  } else {
    var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
    Object.keys(patch).forEach(function (key) {
      var colIdx = cols.indexOf(key);
      if (colIdx === -1) return;
      sheet.getRange(rowIdx, colIdx + 1).setValue(patch[key] === null ? "" : patch[key]);
    });
    sheet.getRange(rowIdx, cols.indexOf("updated_at") + 1).setValue(updatedAt);
  }
  return { taskId: taskId, patch: patch };
}

/* =====================================================================
   ESCRITURA DIRECTA EN "01 I Plan de trabajo" Y "02 I Proceso de trabajo"
   ===================================================================== */

var ESTADO_A_SHEET = { "Por hacer": "Pendiente", "En proceso": "En proceso", "En revisión": "Revisar", "Completado": "Finalizada", "Bloqueado": "Bloqueado" };
var COLS_02 = { area: 3, tema: 4, tarea: 5, responsable: 6, inicio: 7, cierre: 9, estado: 10, obs: 11 };
var TITULO_REUNION = "Encuentro I Estado del proceso de trabajo";

/* Mismo id que calcula la web: hash del contenido + contador de repetidas. */
function hashId_(str) {
  var h = 5381;
  for (var i = 0; i < str.length; i++) h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  return "t" + h.toString(36);
}

/* Ubica cada tarea de 02 (id, fila, sección) con el mismo recorrido que readProceso_. */
function procesoLayout_() {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  var values = sheet.getDataRange().getValues();
  function findRow(text) {
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][1] || "").trim().toUpperCase() === text) return i;
    }
    return -1;
  }
  var iniciativasRow = findRow("INICIATIVAS"), finalizadosRow = findRow("FINALIZADOS"), pendientesRow = findRow("PENDIENTES");
  if (iniciativasRow === -1 || finalizadosRow === -1) {
    throw new Error("No encuentro INICIATIVAS / FINALIZADOS en '" + SHEET_PROCESO + "'");
  }
  var finIniciativas = pendientesRow !== -1 ? pendientesRow : finalizadosRow;
  var tasks = [], seen = {};
  function scan(fromRow, toRow, source) {
    var last = fromRow;
    for (var r = fromRow + 1; r < toRow; r++) {
      var row = values[r];
      var area = cell_(row, 2), tarea = cell_(row, 4);
      if (!area && !tarea) continue;
      var key = [source, area || "", cell_(row, 3) || "", tarea || ""].join("|");
      seen[key] = (seen[key] || 0) + 1;
      tasks.push({ id: hashId_(key + "#" + seen[key]), row: r + 1, source: source });
      last = r;
    }
    return last + 1; // 1-based
  }
  var lastIniciativaRow = scan(iniciativasRow + 2, finIniciativas, "iniciativa");
  var lastFinalizadoRow = scan(finalizadosRow, values.length, "finalizado");
  return { sheet: sheet, tasks: tasks, lastIniciativaRow: lastIniciativaRow, lastFinalizadoRow: lastFinalizadoRow };
}

function findTask_(layout, id) {
  for (var i = 0; i < layout.tasks.length; i++) if (layout.tasks[i].id === id) return layout.tasks[i];
  return null;
}

function idAtRow_(row) {
  var L = procesoLayout_();
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].row === row) return L.tasks[i].id;
  return null;
}

/* "yyyy-mm-dd" -> fecha a mediodía, para que no se corra de día por zona horaria. */
function toSheetDate_(iso) {
  if (!iso) return "";
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return iso;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function writeTaskCells_(sheet, row, fields) {
  Object.keys(COLS_02).forEach(function (k) {
    if (fields[k] === undefined) return;
    var v = fields[k];
    if (k === "inicio" || k === "cierre") v = toSheetDate_(v);
    else if (k === "estado") v = ESTADO_A_SHEET[v] || v || "";
    else if (v === null) v = "";
    sheet.getRange(row, COLS_02[k]).setValue(v);
  });
}

/* Mueve una fila (valores y formato) debajo de afterRow. Devuelve su fila final. */
function moveRow_(sheet, fromRow, afterRow) {
  sheet.insertRowAfter(afterRow);
  var to = afterRow + 1;
  if (fromRow >= to) fromRow++;
  var width = sheet.getMaxColumns();
  sheet.getRange(fromRow, 1, 1, width).copyTo(sheet.getRange(to, 1, 1, width));
  sheet.getRange(to, 2).setValue(""); // el n° de FINALIZADOS no aplica a la fila movida
  sheet.deleteRow(fromRow);
  return fromRow < to ? to - 1 : to;
}

/* El id cambia si cambia el texto o la sección: el link guardado pasa al id nuevo.
   Estado y fechas ya quedaron en 02, así que se limpian del override. */
function migrateOverride_(oldId, newId, fields) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
  var rowIdx = findRowIndexById_(sheet, oldId);
  if (rowIdx !== -1) {
    var link = sheet.getRange(rowIdx, cols.indexOf("link") + 1).getValue();
    if (fields.link !== undefined) link = fields.link || "";
    sheet.getRange(rowIdx, 1, 1, cols.length).setValues([[newId, "", link, "", "", false, nowIso_()]]);
  } else if (fields.link) {
    setOverride_(newId, { link: fields.link });
  }
}

function updateTask_(id, fields) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en '" + SHEET_PROCESO + "'. Puede haber cambiado en el Sheet: recargá la página.");
  writeTaskCells_(L.sheet, t.row, fields);
  var row = t.row;
  if (fields.estado !== undefined) {
    var done = fields.estado === "Completado";
    if (done && t.source === "iniciativa") row = moveRow_(L.sheet, row, L.lastFinalizadoRow);
    else if (!done && t.source === "finalizado") row = moveRow_(L.sheet, row, L.lastIniciativaRow);
  }
  var newId = idAtRow_(row);
  migrateOverride_(id, newId, fields);
  return { id: newId };
}

function addTask_(fields) {
  var L = procesoLayout_();
  var after = fields.estado === "Completado" ? L.lastFinalizadoRow : L.lastIniciativaRow;
  L.sheet.insertRowAfter(after);
  var row = after + 1;
  L.sheet.getRange(row, 1, 1, L.sheet.getMaxColumns()).clearContent();
  writeTaskCells_(L.sheet, row, Object.assign({ estado: "Por hacer" }, fields));
  var newId = idAtRow_(row);
  if (fields.link) setOverride_(newId, { link: fields.link });
  return { id: newId };
}

function deleteTask_(id) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en '" + SHEET_PROCESO + "'. Recargá la página.");
  L.sheet.deleteRow(t.row);
  deleteRow_(SHEET_OVERRIDES, id);
  return { deleted: id };
}

function etapa1Table_() {
  var sheet = ss_().getSheetByName(SHEET_ETAPA1);
  var values = sheet.getDataRange().getValues();
  var header = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === "Fecha") { header = i; break; }
  }
  if (header === -1) throw new Error("No encuentro la tabla de reuniones en '" + SHEET_ETAPA1 + "'");
  var last = header;
  while (last + 1 < values.length && cell_(values[last + 1], 0)) last++;
  return { sheet: sheet, values: values, header: header, last: last };
}

/* "2 hs" -> 2 ; "2.30" se deja como texto, igual que en la hoja. */
function horas_(duracion) {
  var m = String(duracion || "").match(/\d+(?:[.,]\d+)?/);
  if (!m) return "";
  return /[.,]/.test(m[0]) ? m[0] : Number(m[0]);
}

/* La tabla de 01 tiene desplegables/chips (Responsable, Estado...) que
   descartan en silencio lo que no está en su lista. Por eso se copia la
   última reunión (formato y desplegables incluidos), se escribe celda por
   celda y se verifica: si una celda rechaza el valor, queda el valor válido
   copiado y el dato va a Observaciones para no perderlo. */
function addMeetingSheet_(p) {
  var T = etapa1Table_();
  var sheet = T.sheet;
  var src = T.last + 1; // última reunión (1-based)
  sheet.insertRowAfter(src);
  var row = src + 1;
  sheet.getRange(src, 1, 1, 7).copyTo(sheet.getRange(row, 1, 1, 7));

  var wanted = [
    { col: 1, label: "Fecha", value: toSheetDate_(p.fecha), keepCopy: false },
    { col: 2, label: "Hs", value: horas_(p.duracion), keepCopy: false },
    { col: 3, label: "Tarea", value: p.tarea || TITULO_REUNION, keepCopy: false },
    { col: 4, label: "Responsable", value: p.responsable || "", keepCopy: true },
    { col: 5, label: "Estado", value: "Finalizado", keepCopy: true },
    { col: 6, label: "Resultado", value: p.resumen || "", keepCopy: false }
  ];
  var perdidos = [];
  wanted.forEach(function (w) {
    var cell = sheet.getRange(row, w.col);
    var copied = cell.getValue();
    if (w.value === "" && w.keepCopy) return; // sin dato: queda el valor copiado
    var got = "";
    try {
      cell.setValue(w.value);
      SpreadsheetApp.flush();
      got = cell.getValue();
    } catch (err) {
      got = ""; // la validación de la celda rechazó el valor (ej. Responsable: Loopa/Equipo/Weber)
    }
    if (w.value !== "" && (got === "" || got === null)) {
      cell.setValue(w.keepCopy ? copied : "");
      perdidos.push(w.label + ": " + (w.value instanceof Date ? p.fecha : w.value));
    }
  });
  sheet.getRange(row, 7).setValue(perdidos.join(" | "));
  SpreadsheetApp.flush();
  return { row: row, guardado: sheet.getRange(row, 1, 1, 7).getDisplayValues()[0], enObservaciones: perdidos };
}

function deleteMeetingSheet_(p) {
  var T = etapa1Table_();
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    if (toIsoDate_(row[0]) === p.fecha &&
        String(cell_(row, 2) || "") === String(p.tarea || "") &&
        String(cell_(row, 5) || "") === String(p.resultado || "")) {
      T.sheet.deleteRow(r + 1);
      return { deleted: true };
    }
  }
  throw new Error("No encuentro esa reunión en '" + SHEET_ETAPA1 + "'. Recargá la página.");
}

/* Pasa a 02 las tareas que quedaron en "WebApp - Tareas nuevas" (versión anterior). */
function migrarWebAppAlSheet_() {
  var rows = readSimpleRows_(SHEET_CUSTOM_TASKS, SHEET_SCHEMAS[SHEET_CUSTOM_TASKS], ["inicio", "cierre"]);
  var log = [];
  rows.forEach(function (c) {
    var fields = { tarea: c.tarea, area: c.area, tema: c.tema, responsable: c.responsable,
      inicio: c.inicio, cierre: c.cierre, estado: c.estado || "Por hacer", obs: c.obs };
    if (c.link) fields.link = c.link;
    var id = String(c.id), res;
    if (id.charAt(0) === "x" && findTask_(procesoLayout_(), id.slice(1))) res = updateTask_(id.slice(1), fields);
    else res = addTask_(fields);
    deleteRow_(SHEET_CUSTOM_TASKS, id);
    log.push(id + " -> " + res.id);
  });
  return log;
}
