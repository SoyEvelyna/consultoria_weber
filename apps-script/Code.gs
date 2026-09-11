/**
 * Weber Tracker — backend de Google Apps Script.
 *
 * Se pega en el editor de Apps Script del Google Sheet
 * "Loopa I Weber I Plan de trabajo" (Extensiones → Apps Script) y se
 * publica como Web App. La web del tracker le habla por HTTP (fetch).
 *
 * DISEÑO:
 * - Las hojas "01 I Plan de trabajo" y "02 I Proceso de trabajo" se leen
 *   TAL CUAL están, en cada pedido — nunca se escriben. Podés seguir
 *   editándolas a mano y descargando el Excel sin que la web interfiera.
 * - Los cambios que se hacen desde la web (estado/fecha/link de una tarea,
 *   tareas nuevas, notas, reuniones) se guardan en pestañas nuevas que este
 *   script crea solo si no existen: "WebApp - Overrides", "WebApp - Notas",
 *   "WebApp - Tareas nuevas", "WebApp - Reuniones".
 * - Autenticación: un token compartido (ver ACCESS_TOKEN abajo), el mismo
 *   que se configura en la web. No hay login individual.
 */

var ACCESS_TOKEN = "OKSiDEIeEQ55k3kCXS3cOL0f53TNBR9U";

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
        ["id", "tarea", "area", "tema", "responsable", "fase", "inicio", "cierre", "estado", "obs", "link", "createdAt"]),
      meetings: readSimpleRows_(SHEET_MEETINGS, ["id", "fecha", "responsable", "duracion", "resumen", "createdAt"])
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
      { id: "n" + Date.now(), text: p.text, author: p.author, createdAt: nowIso_() });
    case "deleteNote": return deleteRow_(SHEET_NOTES, p.id);
    case "addCustomTask": return addRow_(SHEET_CUSTOM_TASKS,
      ["id", "tarea", "area", "tema", "responsable", "fase", "inicio", "cierre", "estado", "obs", "link", "createdAt"],
      Object.assign({ id: "c" + Date.now(), createdAt: nowIso_() }, p));
    case "updateCustomTask": return updateRow_(SHEET_CUSTOM_TASKS, p.id, p.patch || {});
    case "deleteCustomTask": return deleteRow_(SHEET_CUSTOM_TASKS, p.id);
    case "addMeeting": return addRow_(SHEET_MEETINGS, ["id", "fecha", "responsable", "duracion", "resumen", "createdAt"],
      Object.assign({ id: "m" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteMeeting": return deleteRow_(SHEET_MEETINGS, p.id);
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
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ETAPA1);
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
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PROCESO);
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
  var pendientesRow = findRow(1, "PENDIENTES");
  var finalizadosRow = findRow(1, "FINALIZADOS");
  if ([objetivoRow, prioridadesRow, iniciativasRow, pendientesRow, finalizadosRow].indexOf(-1) !== -1) {
    throw new Error("No encuentro alguna de las secciones (OBJETIVO 1 / PRIORIDADES / INICIATIVAS / PENDIENTES / FINALIZADOS) en '" + SHEET_PROCESO + "'. ¿Se movieron o renombraron?");
  }

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
  for (var ir = iniciativasHeaderRow + 1; ir < pendientesRow; ir++) {
    var row = values[ir];
    var area = cell_(row, 2), tarea = cell_(row, 4);
    if (!area && !tarea) continue;
    iniciativas.push({
      area: area, tema: cell_(row, 3), tarea: tarea, responsable: cell_(row, 5),
      inicio: toIsoDate_(cell_(row, 6)), tiempo: toIsoDate_(cell_(row, 7)), cierre: toIsoDate_(cell_(row, 8)),
      estado: cell_(row, 9), obs: cell_(row, 10)
    });
  }

  // Pendientes / backlog: solo texto libre en columna E
  var backlog = [];
  for (var brow = pendientesRow + 1; brow < finalizadosRow; brow++) {
    var txt = cell_(values[brow], 4);
    if (txt) backlog.push(txt);
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
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
      sheet.setFrozenRows(1);
    }
  });
}

function readSimpleRows_(sheetName, cols) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!cell_(row, 0)) continue;
    var obj = {};
    cols.forEach(function (c, i) { obj[c] = cell_(row, i); });
    out.push(obj);
  }
  return out;
}

function readOverrides_() {
  var rows = readSimpleRows_(SHEET_OVERRIDES, ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"]);
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
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  sheet.appendRow(cols.map(function (c) { return obj[c] === undefined ? "" : obj[c]; }));
  return obj;
}

function updateRow_(sheetName, id, patch) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
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
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx !== -1) sheet.deleteRow(rowIdx);
  return { id: id, deleted: rowIdx !== -1 };
}

function setOverride_(taskId, patch) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_OVERRIDES);
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
