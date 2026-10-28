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

/* Para ejecutar a mano desde el editor (el menú no muestra funciones con "_").
   Van primero porque el editor corre la primera función del archivo. */
function probarResumenAhora() { Logger.log(JSON.stringify(probarResumen_())); }
function quienEnvia() { Logger.log(JSON.stringify({ efectivo: Session.getEffectiveUser().getEmail(), activo: Session.getActiveUser().getEmail() })); }
function activarResumenDiario() { Logger.log(JSON.stringify(activarResumen_())); }
function desactivarResumenDiario() { Logger.log(JSON.stringify(desactivarResumen_())); }

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
var SHEET_LOGS = "WebApp - Logs";
var SHEET_AGENDA = "WebApp - Agenda";
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
      logs: readLogs_(),
      agenda: readAgenda_(),
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
    case "addLog": return addRow_(SHEET_LOGS, SHEET_SCHEMAS[SHEET_LOGS], p);
    case "addAgenda": return addRow_(SHEET_AGENDA, SHEET_SCHEMAS[SHEET_AGENDA], p);
    case "updateAgenda": return updateRow_(SHEET_AGENDA, p.id, p.patch || {});
    case "deleteAgenda": return deleteRow_(SHEET_AGENDA, p.id);
    case "syncCalendario": return sincronizarCalendario();
    case "quitarOpcion": return quitarOpcion_(p.campo, p.valor);
    case "setObjetivo": return setObjetivo_(p.texto);
    case "setPrioridad": return setPrioridad_(p.n, p.texto, p.desc);
    case "agregarPrioridad": return agregarPrioridad_(p.texto, p.desc);
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
    case "migrarEstados": return migrarEstados_();
    case "migrarPrioridad": return migrarPrioridad_();
    case "migrarLinks": return migrarLinks_();
    case "updateTask": return updateTask_(p.id, p.fields || {});
    case "addTask": return addTask_(p.fields || {});
    case "deleteTask": return deleteTask_(p.id);
    case "addMeetingSheet": return addMeetingSheet_(p);
    case "deleteMeetingSheet": return deleteMeetingSheet_(p);
    case "migrarWebApp": return migrarWebAppAlSheet_();
    case "probarResumen": return probarResumen_();
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
    return Utilities.formatDate(value, tzHoja_(), "yyyy-MM-dd");
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
  proceso.opciones = readOpciones_();
  return proceso;
}

/* Opciones que ofrecen los desplegables de la hoja, para que la web proponga
   las mismas y no se rechace lo que se escribe. */
function readOpciones_() {
  var out = { area: [], responsable: [], responsableReuniones: [] };
  try {
    var L = procesoLayout_();
    var row = L.tasks.length ? L.tasks[0].row : null;
    if (row) {
      ["area", "responsable"].forEach(function (k) {
        if (L.cols[k] !== undefined) out[k] = dropdownValues_(L.sheet.getRange(row, L.cols[k]));
      });
    }
  } catch (err) {}
  try {
    var T = etapa1Table_();
    out.responsableReuniones = dropdownValues_(T.sheet.getRange(T.header + 2, 4));
  } catch (err2) {}
  return out;
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

/* Ubica las columnas de la tabla de tareas por el texto del encabezado, como
   en Noctis. TAREA no tiene encabezado propio (ahí va el cartel de la etapa):
   es la columna siguiente a TEMA. Índices 0-based. */
function colsProceso_(values) {
  var NOMBRES = { "PRIORIDAD": "prioridad", "AREA": "area", "TEMA": "tema", "RESPONSABLE": "responsable",
    "INICIO": "inicio", "TIEMPO": "tiempo", "CIERRE": "cierre", "ESTADO": "estado", "OBSERVACIONES": "obs" };
  for (var r = 0; r < values.length; r++) {
    var fila = values[r], map = null;
    for (var c = 0; c < fila.length; c++) {
      var k = NOMBRES[sinAcentos_(fila[c])];
      if (!k) continue;
      map = map || {};
      if (map[k] === undefined) map[k] = c;
    }
    if (map && map.area !== undefined && map.tema !== undefined && map.estado !== undefined) {
      map.tarea = map.tema + 1;
      map.header = r;
      return map;
    }
  }
  throw new Error("No encuentro el encabezado (AREA / TEMA / ESTADO) en '" + SHEET_PROCESO + "'");
}

function sinAcentos_(v) {
  return String(v == null ? "" : v).trim().toUpperCase()
    .replace(/[ÁÀÄÂ]/g, "A").replace(/[ÉÈËÊ]/g, "E").replace(/[ÍÌÏÎ]/g, "I")
    .replace(/[ÓÒÖÔ]/g, "O").replace(/[ÚÙÜÛ]/g, "U");
}

/* Columnas escribibles, 1-based, para writeTaskCells_. */
function colsEscritura_(C) {
  var out = {};
  ["prioridad", "area", "tema", "tarea", "responsable", "inicio", "cierre", "estado", "obs"].forEach(function (k) {
    if (C[k] !== undefined) out[k] = C[k] + 1;
  });
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
  var C = colsProceso_(values);
  var richObs = C.obs === undefined ? null : sheet.getRange(1, C.obs + 1, values.length, 1).getRichTextValues();

  // Prioridades: 3 filas después de "PRIORIDADES" (la pregunta), cada una: [n, "titulo\t...\tdesc"]
  var prioridades = [];
  for (var pr = prioridadesRow + 2; pr < iniciativasRow; pr++) {
    var n = cell_(values[pr], 1);
    var raw = cell_(values[pr], 2) || cell_(values[pr], 3) || cell_(values[pr], 4);
    if (n === null || raw === null) continue;
    var parts = String(raw).split(/\t+/).map(function (s) { return s.trim(); }).filter(Boolean);
    prioridades.push({ n: n, tt: parts[0] || "", desc: parts.slice(1).join(" ") });
  }

  // Iniciativas: encabezado (ÁREA..OBSERVACIONES) en columna C, 2 filas después de "INICIATIVAS"
  var iniciativasHeaderRow = iniciativasRow + 2;
  var iniciativas = [];
  for (var ir = iniciativasHeaderRow + 1; ir < finIniciativas; ir++) {
    if (!tieneTarea_(values[ir], C)) continue;
    iniciativas.push(leerTarea_(values[ir], C, richObs ? richObs[ir][0] : null));
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
    if (!tieneTarea_(values[fr], C)) continue;
    finalizados.push(leerTarea_(values[fr], C, richObs ? richObs[fr][0] : null));
  }

  return {
    objetivo: objetivo,
    prioridades: prioridades,
    iniciativas: iniciativas,
    finalizados: finalizados,
    backlog: backlog
  };
}

function tieneTarea_(row, C) {
  return !!(cell_(row, C.area) || cell_(row, C.tarea));
}

function leerTarea_(row, C, richObs) {
  function v(k) { return C[k] === undefined ? null : cell_(row, C[k]); }
  var obsText = v("obs");
  var links = obsLinks_(richObs, obsText);
  return {
    link: links.length ? links.join("\n") : null,
    prioridad: prioridadNorm_(v("prioridad")), area: v("area"), tema: v("tema"), tarea: v("tarea"),
    responsable: v("responsable"), inicio: toIsoDate_(v("inicio")), tiempo: toIsoDate_(v("tiempo")),
    cierre: toIsoDate_(v("cierre")), estado: v("estado"), obs: obsSinLinks_(obsText, links)
  };
}

var MAX_LINKS = 3;

/* Links de OBSERVACIONES (hasta 3): textos enlazados en la celda y URLs escritas. */
function obsLinks_(richValue, text) {
  var out = [];
  function add(u) { if (u && out.indexOf(u) === -1 && out.length < MAX_LINKS) out.push(u); }
  if (richValue) {
    add(richValue.getLinkUrl());
    richValue.getRuns().forEach(function (run) { add(run.getLinkUrl()); });
  }
  (String(text || "").match(/https?:\/\/\S+/g) || []).forEach(add);
  return out;
}

/* Texto de OBSERVACIONES sin las líneas que son solo un link. */
function obsSinLinks_(text, links) {
  if (!text) return null;
  var lines = String(text).split(/\n/).filter(function (line) {
    var l = line.trim();
    return l && links.indexOf(l) === -1 && !/^https?:\/\/\S+$/.test(l);
  });
  var t = lines.join("\n").trim();
  return t || null;
}

/* Deja en la celda de OBSERVACIONES el texto y, debajo, un link por línea.
   Una línea que trae texto además del link (ej. "Reposteo de Ger: https://…")
   se guarda como texto, y la URL de adentro igual queda clickeable. */
function escribirObsYLinks_(cell, fields, perdidos) {
  var currentText = String(cell.getValue() || "");
  var currentLinks = obsLinks_(cell.getRichTextValue(), currentText);
  var partes = fields.link !== undefined
    ? String(fields.link || "").split(/\s*\n\s*|\s+\|\s+/).map(function (x) { return x.trim(); }).filter(Boolean)
    : currentLinks;
  var links = [], conTexto = [];
  partes.forEach(function (x) {
    if (/^https?:\/\/\S+$/.test(x)) { if (links.length < MAX_LINKS) links.push(x); }
    else conTexto.push(x);
  });
  var base = fields.obs !== undefined ? (fields.obs || "") : (obsSinLinks_(currentText, currentLinks) || "");
  base = [base].concat(conTexto).concat(perdidos || []).filter(Boolean).join("\n");
  var text = [base].concat(links).filter(Boolean).join("\n");
  if (!text) { cell.setValue(""); return; }
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var re = /https?:\/\/[^\s]+/g, m;
  while ((m = re.exec(text)) !== null) {
    try { builder = builder.setLinkUrl(m.index, m.index + m[0].length, m[0]); } catch (err) {}
  }
  cell.setRichTextValue(builder.build());
}

/* "1", "p1", "P1 " -> "P1". Cualquier otra cosa se deja como está. */
function prioridadNorm_(v) {
  if (v === null || v === undefined || v === "") return null;
  var m = String(v).trim().match(/^p?\s*([123])$/i);
  return m ? "P" + m[1] : String(v).trim();
}

/* =====================================================================
   PESTAÑAS DE LA WEB (WebApp - *) — creadas si no existen
   ===================================================================== */

var SHEET_SCHEMAS = {};
SHEET_SCHEMAS[SHEET_OVERRIDES] = ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at", "urgencia"];
SHEET_SCHEMAS[SHEET_NOTES] = ["id", "text", "author", "createdAt"];
SHEET_SCHEMAS[SHEET_LOGS] = ["id", "ts", "autor", "accion", "tarea", "detalle"];
SHEET_SCHEMAS[SHEET_AGENDA] = ["id", "fecha", "hora", "tema", "convocados", "link", "creadoPor", "createdAt"];
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
    asegurarColumnas_(name);
  });
}
/* Las pestañas de la web las crea ensureSheets_, pero una que ya existía se
   quedó con las columnas de cuando se creó. Esto agrega al final las que
   falten, sin tocar ni mover nada de lo que ya hay. */
function asegurarColumnas_(nombre) {
  var sheet = ss_().getSheetByName(nombre);
  if (!sheet) return;
  var cols = SHEET_SCHEMAS[nombre];
  var ancho = Math.max(sheet.getLastColumn(), 1);
  var actuales = sheet.getRange(1, 1, 1, ancho).getValues()[0]
    .map(function (v) { return String(v === null || v === undefined ? "" : v).trim(); });
  while (actuales.length && !actuales[actuales.length - 1]) actuales.pop();
  for (var i = actuales.length; i < cols.length; i++) {
    sheet.getRange(1, i + 1).setValue(cols[i]);
  }
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
  var rows = readSimpleRows_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES],
    ["inicio", "cierre"]);
  var out = {};
  rows.forEach(function (r) {
    out[r.task_id] = { estado: r.estado, link: r.link, inicio: r.inicio, cierre: r.cierre, hidden: r.hidden === true || r.hidden === "true", updatedAt: r.updated_at, urgencia: r.urgencia || null };
  });
  return out;
}

/* =====================================================================
   LOGS · registro de ediciones
   Cada cambio hecho desde la web agrega una fila. Se leen sólo los últimos
   movimientos: la hoja crece sin tope y el tablero muestra los días recientes.
   ===================================================================== */
var LOGS_MAX = 500;

/* =====================================================================
   AGENDA · las reuniones que todavía no pasaron
   Viven en su propia pestaña de la web. Las que ya se hicieron siguen
   viniendo de la hoja 01, que es donde las escribe Fathom: así una cosa no
   pisa la otra y la agenda no ensucia el historial.
   ===================================================================== */
function readAgenda_() {
  var sheet = ss_().getSheetByName(SHEET_AGENDA);
  if (!sheet) return [];
  var cols = SHEET_SCHEMAS[SHEET_AGENDA];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!cell_(values[r], 0)) continue;
    var obj = {};
    for (var c = 0; c < cols.length; c++) {
      var v = cell_(values[r], c);
      if (cols[c] === "fecha") v = toIsoDate_(v);
      else if (cols[c] === "hora" && v instanceof Date) {
        v = Utilities.formatDate(v, tzHoja_(), "HH:mm");
      }
      obj[cols[c]] = (v === null || v === undefined) ? "" : String(v);
    }
    out.push(obj);
  }
  out.sort(function (a, b) {
    return String(a.fecha + " " + a.hora).localeCompare(String(b.fecha + " " + b.hora));
  });
  return out;
}

function readLogs_() {
  var sheet = ss_().getSheetByName(SHEET_LOGS);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var cols = SHEET_SCHEMAS[SHEET_LOGS];
  var desde = Math.max(2, last - LOGS_MAX + 1);
  var values = sheet.getRange(desde, 1, last - desde + 1, cols.length).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    if (!cell_(values[r], 0)) continue;
    var obj = {};
    for (var c = 0; c < cols.length; c++) {
      var v = cell_(values[r], c);
      if (cols[c] === "ts" && v instanceof Date) v = v.toISOString();
      obj[cols[c]] = (v === null || v === undefined) ? "" : String(v);
    }
    out.push(obj);
  }
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

var ESTADO_A_SHEET = { "Por hacer": "Pendiente", "En proceso": "Proceso", "En revisión": "Revisar", "Testear": "Testear", "Completado": "Finalizada" };
var ESTADOS_SHEET = ["Pendiente", "Proceso", "Revisar", "Testear", "Finalizada"];
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
  var C = colsProceso_(values);
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
      if (!tieneTarea_(row, C)) continue;
      var key = [source, cell_(row, C.area) || "", cell_(row, C.tema) || "", cell_(row, C.tarea) || ""].join("|");
      seen[key] = (seen[key] || 0) + 1;
      tasks.push({ id: hashId_(key + "#" + seen[key]), row: r + 1, source: source });
      last = r;
    }
    return last + 1; // 1-based
  }
  var lastIniciativaRow = scan(iniciativasRow + 2, finIniciativas, "iniciativa");
  var lastFinalizadoRow = scan(finalizadosRow, values.length, "finalizado");
  return { sheet: sheet, values: values, C: C, cols: colsEscritura_(C), tasks: tasks,
    lastIniciativaRow: lastIniciativaRow, lastFinalizadoRow: lastFinalizadoRow };
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

function norm_(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ").trim().toUpperCase();
}

/* Opciones de un desplegable que NO acepta otros valores (null si acepta cualquiera). */
function listOptions_(cell) {
  var dv = cell.getDataValidation();
  if (!dv || dv.getAllowInvalid()) return null;
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return crit[0];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return crit[0].getValues().map(function (r) { return r[0]; }).filter(function (v) { return v !== ""; });
  }
  return null;
}

/* Valores de un desplegable (acepte o no otros), para ofrecerlos en la web. */
function dropdownValues_(cell) {
  var dv = cell.getDataValidation();
  if (!dv) return [];
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  var vals = [];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) vals = crit[0];
  else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    vals = crit[0].getValues().map(function (r) { return r[0]; });
  }
  return vals.map(function (v) { return String(v).trim(); }).filter(Boolean);
}

/* Escribe el primer candidato que la celda acepte. false si el desplegable los rechaza. */
function setSafe_(cell, candidates) {
  var options = listOptions_(cell);
  var value = candidates[0];
  if (options) {
    var hit = null;
    candidates.forEach(function (c) {
      if (hit !== null) return;
      options.forEach(function (o) { if (hit === null && norm_(o) === norm_(c)) hit = o; });
    });
    if (hit === null) return false;
    value = hit;
  }
  try { cell.setValue(value); return true; } catch (err) { return false; }
}

/* Escribe los campos en su columna. Lo que un desplegable rechaza no tira abajo
   el guardado: se anota en OBSERVACIONES y se avisa a la web. */
function writeTaskCells_(sheet, row, fields, cols) {
  var perdidos = [];
  Object.keys(cols).forEach(function (k) {
    if (k === "obs" || fields[k] === undefined) return;
    var v = fields[k];
    if (v === null || v === "") { try { sheet.getRange(row, cols[k]).setValue(""); } catch (err) {} return; }
    var candidatos = k === "inicio" || k === "cierre" ? [toSheetDate_(v)]
      : k === "estado" ? [ESTADO_A_SHEET[v] || v, v]
      : k === "prioridad" ? [prioridadNorm_(v) || v]
      : [v];
    if (!setSafe_(sheet.getRange(row, cols[k]), candidatos)) {
      perdidos.push(k.charAt(0).toUpperCase() + k.slice(1) + ": " + v);
    }
  });
  // OBSERVACIONES guarda el texto, lo que no entró en su columna, y los links.
  if (cols.obs !== undefined && (fields.obs !== undefined || fields.link !== undefined || perdidos.length)) {
    escribirObsYLinks_(sheet.getRange(row, cols.obs), fields, perdidos);
  }
  return perdidos;
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
    // Los links ahora van en OBSERVACIONES: esta fila ya no los guarda.
    var link = sheet.getRange(rowIdx, cols.indexOf("link") + 1).getValue();
    if (fields.link !== undefined) link = "";
    sheet.getRange(rowIdx, 1, 1, cols.length).setValues([[newId, "", link, "", "", false, nowIso_()]]);
  }
}

function updateTask_(id, fields) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en '" + SHEET_PROCESO + "'. Puede haber cambiado en el Sheet: recargá la página.");
  var perdidos = writeTaskCells_(L.sheet, t.row, fields, L.cols);
  var row = t.row;
  if (fields.estado !== undefined) {
    var done = fields.estado === "Completado";
    if (done && t.source === "iniciativa") row = moveRow_(L.sheet, row, L.lastFinalizadoRow);
    else if (!done && t.source === "finalizado") row = moveRow_(L.sheet, row, L.lastIniciativaRow);
  }
  var newId = idAtRow_(row);
  migrateOverride_(id, newId, fields);
  return { id: newId, enObservaciones: perdidos };
}

function addTask_(fields) {
  var L = procesoLayout_();
  var after = fields.estado === "Completado" ? L.lastFinalizadoRow : L.lastIniciativaRow;
  L.sheet.insertRowAfter(after);
  var row = after + 1;
  L.sheet.getRange(row, 1, 1, L.sheet.getMaxColumns()).clearContent();
  var perdidos = writeTaskCells_(L.sheet, row, Object.assign({ estado: "Por hacer" }, fields), L.cols);
  return { id: idAtRow_(row), enObservaciones: perdidos };
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

/* Estado viejo de la hoja -> estado nuevo (Pendiente, Proceso, Revisar, Testear, Finalizada). */
function estadoNuevo_(v) {
  var s = String(v || "").trim().toLowerCase();
  if (s.indexOf("final") === 0) return "Finalizada";
  if (s.indexOf("en proceso") === 0 || s.indexOf("proceso") === 0 || s.indexOf("actualiz") === 0) return "Proceso";
  if (s.indexOf("revis") === 0 || s.indexOf("propuesta") === 0) return "Revisar";
  if (s.indexOf("test") === 0) return "Testear";
  return "Pendiente";
}

/* Una vez: pasa los links guardados en "WebApp - Overrides" a la columna
   OBSERVACIONES de la hoja 02 y los borra de la pestaña de la web. */
function migrarLinks_() {
  var L = procesoLayout_();
  if (L.cols.obs === undefined) throw new Error("No encuentro la columna OBSERVACIONES en '" + SHEET_PROCESO + "'");
  var overrides = readOverrides_();
  var movidos = 0;
  L.tasks.forEach(function (t) {
    var ov = overrides[t.id];
    if (!ov || !ov.link) return;
    escribirObsYLinks_(L.sheet.getRange(t.row, L.cols.obs), { link: ov.link });
    setOverride_(t.id, { link: "" });
    movidos++;
  });
  return { tareas: L.tasks.length, linksMovidos: movidos };
}

var PRIORIDADES_SHEET = ["P1", "P2", "P3"];

/* Prioridad sugerida para una tarea que todavía no la tiene. Misma
   clasificación por palabras clave que usaba el tablero para agrupar. */
function prioridadSugerida_(row, C) {
  function t(k) { return C[k] === undefined ? "" : String(cell_(row, C[k]) || ""); }
  var s = (t("tema") + " " + t("tarea") + " " + t("area")).toLowerCase();
  function tiene(lista) {
    for (var i = 0; i < lista.length; i++) if (s.indexOf(lista[i]) !== -1) return true;
    return false;
  }
  if (tiene(["cliente potenc", "evento", "oilgas", "chatgpt", "expo-logisti", "leads", "comexnp", "competencia",
    "vender", "funnel", "vaca muerta", "infopymes", "seguidores", "contactar"])) return "P3";
  if (tiene(["brand nuevo", "logo", "redes sociales", "whatsapp", "email marketing", "google my business",
    "instagram", "registro de marca", "carpeta institucional", "membretada", "firma de email", "web nueva",
    "web:", "web i", "lanzamiento", "landing", "identidad visual"])) return "P1";
  return "P2";
}

/* Una vez: agrega la columna PRIORIDAD a 02 (antes de AREA), le pone
   desplegable P1/P2/P3 y completa las tareas que estén sin prioridad. */
function migrarPrioridad_() {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  var C = colsProceso_(sheet.getDataRange().getValues());
  var creada = false;
  if (C.prioridad === undefined) {
    var col = C.area + 1; // 1-based: la nueva columna queda donde estaba AREA
    sheet.insertColumnBefore(col);
    var filaHeader = C.header + 1;
    sheet.getRange(filaHeader, col + 1).copyTo(sheet.getRange(filaHeader, col), { formatOnly: true });
    sheet.getRange(filaHeader, col).setValue("PRIORIDAD");
    sheet.setColumnWidth(col, 92);
    creada = true;
    SpreadsheetApp.flush();
  }
  var L = procesoLayout_();
  if (L.cols.prioridad === undefined) throw new Error("No pude crear la columna PRIORIDAD en '" + SHEET_PROCESO + "'");
  var regla = SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES_SHEET, true).setAllowInvalid(true).build();
  var puestas = 0, conteo = {};
  L.tasks.forEach(function (t) {
    var cell = L.sheet.getRange(t.row, L.cols.prioridad);
    cell.setDataValidation(regla);
    var actual = prioridadNorm_(cell.getValue());
    if (actual) { conteo[actual] = (conteo[actual] || 0) + 1; return; }
    var sugerida = prioridadSugerida_(L.values[t.row - 1], L.C);
    cell.setValue(sugerida);
    conteo[sugerida] = (conteo[sugerida] || 0) + 1;
    puestas++;
  });
  return { columnaCreada: creada, tareas: L.tasks.length, completadas: puestas, conteo: conteo };
}

/* Una vez: pasa todas las tareas de 02 a los estados nuevos y, si la columna
   ESTADO tiene desplegable, lo deja con esas 5 opciones. */
function migrarEstados_() {
  var L = procesoLayout_();
  var col = L.cols.estado;
  var cambios = {};
  L.tasks.forEach(function (t) {
    var cell = L.sheet.getRange(t.row, col);
    var antes = String(cell.getValue() || "");
    var despues = estadoNuevo_(antes);
    var dv = cell.getDataValidation();
    if (dv) cell.setDataValidation(dv.copy().requireValueInList(ESTADOS_SHEET, true).build());
    if (antes !== despues) cell.setValue(despues);
    var k = (antes || "(vacío)") + " -> " + despues;
    cambios[k] = (cambios[k] || 0) + 1;
  });
  return cambios;
}

/* =====================================================================
   RESUMEN DIARIO POR EMAIL (18 h)
   Compara las tareas y reuniones de hoy con la "foto" guardada en el último
   resumen y manda un email con lo que cambió (desde la web o a mano en la
   Hoja). Si no cambió nada, no manda nada.
   El envío corre en un proyecto aparte de hola@soyevelyna.com (sin Web App),
   porque el remitente es la cuenta que autoriza el script.
   ===================================================================== */

var RESUMEN_DESTINATARIOS = [
  "hola@soyevelyna.com",
  "gweber@webersi.com.ar",
  "marketingdigital.push@gmail.com",
  "juliana@loopa.io",
  "leo@loopa.io"
];
var RESUMEN_PRUEBA = "hola@soyevelyna.com,abrahanevelyn@gmail.com";  /* solo para las pruebas */
var RESUMEN_ZONA = "America/Argentina/Buenos_Aires";
var RESUMEN_HORA = 18;
var TABLERO_URL = "https://soyevelyna.github.io/consultoria_weber/";
var SHEET_RESUMEN = "WebApp - Resumen";

function fotoActual_() {
  var seed = readSeed_();
  function tarea(t, seccion) {
    return { seccion: seccion, area: t.area, tema: t.tema, tarea: t.tarea, responsable: t.responsable,
      inicio: t.inicio, cierre: t.cierre, estado: t.estado, obs: t.obs };
  }
  var tareas = seed.iniciativas.map(function (t) { return tarea(t, "iniciativas"); })
    .concat(seed.finalizados.map(function (t) { return tarea(t, "finalizados"); }));
  return {
    fecha: nowIso_(),
    tareas: tareas,
    reuniones: seed.etapa1.map(function (m) {
      return { fecha: m.fecha, tarea: m.tarea, responsable: m.responsable, resultado: m.resultado };
    })
  };
}

function leerFoto_() {
  var sh = ss_().getSheetByName(SHEET_RESUMEN);
  if (!sh) return null;
  var v = sh.getRange(1, 1).getValue();
  try { return v ? JSON.parse(v) : null; } catch (err) { return null; }
}

function guardarFoto_(foto) {
  var texto = JSON.stringify(foto);
  if (texto.length > 49000) throw new Error("La foto del resumen es demasiado grande para una celda");
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_RESUMEN) || ss.insertSheet(SHEET_RESUMEN);
  sh.getRange(1, 1).setValue(texto);
  sh.getRange(2, 1).setValue("Foto que usa el resumen diario por email. No editar. Última: " + (foto.fecha || nowIso_()));
}

function claveTarea_(t) { return String(t.tarea || "").trim().toLowerCase(); }

function difResumen_(antes, ahora) {
  var r = { nuevas: [], finalizadas: [], estados: [], otros: [], eliminadas: [], reuniones: [] };
  var mapaAntes = {};
  (antes.tareas || []).forEach(function (t) { mapaAntes[claveTarea_(t)] = t; });
  var vistas = {};
  var campos = [["responsable", "Responsable"], ["inicio", "Inicio"], ["cierre", "Cierre"],
    ["area", "Área"], ["tema", "Tema"], ["obs", "Observaciones"]];
  (ahora.tareas || []).forEach(function (t) {
    var k = claveTarea_(t);
    vistas[k] = true;
    var p = mapaAntes[k];
    if (!p) { r.nuevas.push(t); return; }
    var eAntes = String(p.estado || ""), eAhora = String(t.estado || "");
    var pasoAFinalizados = p.seccion !== "finalizados" && t.seccion === "finalizados";
    if (pasoAFinalizados || (eAntes !== eAhora && /^final/i.test(eAhora))) r.finalizadas.push(t);
    else if (eAntes !== eAhora) r.estados.push({ t: t, de: eAntes || "(sin estado)", a: eAhora || "(sin estado)" });
    var cambios = [];
    campos.forEach(function (c) {
      var a = p[c[0]] == null ? "" : String(p[c[0]]);
      var b = t[c[0]] == null ? "" : String(t[c[0]]);
      if (a !== b) cambios.push(c[1] + ": " + (a || "—") + " → " + (b || "—"));
    });
    if (cambios.length) r.otros.push({ t: t, cambios: cambios });
  });
  (antes.tareas || []).forEach(function (p) { if (!vistas[claveTarea_(p)]) r.eliminadas.push(p); });
  /* Las reuniones quedan fuera del resumen a propósito: se cargan en el
     tablero y ahí se ven. Si una reunión nueva es el único cambio del día,
     no sale ningún mail. */
  r.total = r.nuevas.length + r.finalizadas.length + r.estados.length + r.otros.length + r.eliminadas.length;
  return r;
}

function htmlResumen_(r, hoy, prueba) {
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function corto(s) { s = String(s == null ? "" : s); return s.length > 110 ? s.slice(0, 107) + "…" : s; }
  function linea(t, extra) {
    return "<li style='margin-bottom:6px'><b>" + esc(t.tarea) + "</b>" + (t.responsable ? " · " + esc(t.responsable) : "") +
      (extra ? "<br><span style='color:#555'>" + extra + "</span>" : "") + "</li>";
  }
  function seccion(titulo, color, items) {
    if (!items.length) return "";
    return "<h3 style='margin:20px 0 6px;font-family:Arial,sans-serif;font-size:16px;color:" + color + "'>" + titulo + " (" + items.length + ")</h3>" +
      "<ul style='margin:0;padding-left:18px;font-family:Arial,sans-serif;font-size:14px;line-height:1.45'>" + items.join("") + "</ul>";
  }
  var html = "<div style='max-width:640px'>";
  if (prueba) {
    html += "<p style='background:#FFF3CD;padding:8px 10px;font-family:Arial,sans-serif;font-size:13px'>Email de prueba: así se va a ver el resumen diario de las 18 h. Por ahora solo te llega a vos.</p>";
  }
  html += "<h2 style='font-family:Arial,sans-serif;margin:0 0 4px'>Weber · Resumen del " + esc(hoy) + "</h2>";
  html += "<p style='font-family:Arial,sans-serif;color:#555;margin:0'>Cambios en el tablero desde el último resumen.</p>";
  html += seccion("Tareas nuevas", "#212121", r.nuevas.map(function (t) { return linea(t, esc(t.estado || "")); }));
  html += seccion("Finalizadas", "#23914A", r.finalizadas.map(function (t) { return linea(t); }));
  html += seccion("Cambios de estado", "#7447B8", r.estados.map(function (x) { return linea(x.t, esc(x.de) + " → " + esc(x.a)); }));
  html += seccion("Otros cambios", "#1F7FB5", r.otros.map(function (x) {
    return linea(x.t, x.cambios.map(function (c) { return esc(corto(c)); }).join("<br>"));
  }));
  html += seccion("Eliminadas", "#B9282D", r.eliminadas.map(function (t) { return linea(t); }));
  if (!r.total) html += "<p style='font-family:Arial,sans-serif;margin-top:16px'>No hubo cambios.</p>";
  html += "<p style='font-family:Arial,sans-serif;margin-top:24px'><a href='" + TABLERO_URL + "'>Abrir el tablero de Weber</a></p></div>";
  return html;
}

/* Lo ejecuta el activador diario (18 h). También guarda la foto nueva. */
function enviarResumenDiario() {
  var ahora = fotoActual_();
  var antes = leerFoto_();
  if (!antes) { guardarFoto_(ahora); return { enviado: false, motivo: "primera foto guardada" }; }
  var r = difResumen_(antes, ahora);
  if (r.total) {
    var hoy = Utilities.formatDate(new Date(), RESUMEN_ZONA, "dd/MM/yyyy");
    MailApp.sendEmail({
      to: RESUMEN_DESTINATARIOS.join(","),
      subject: "WEBER I Tablero de trabajo: Resumen del día " + hoy + " (" + r.total + (r.total === 1 ? " cambio)" : " cambios)"),
      htmlBody: htmlResumen_(r, hoy, false),
      name: "WEBER I Tablero de trabajo"
    });
  }
  guardarFoto_(ahora);
  return { enviado: !!r.total, cambios: r.total };
}

/* Prueba: manda el resumen solo a RESUMEN_PRUEBA, sin tocar la foto guardada. */
function probarResumen_() {
  var ahora = fotoActual_();
  var antes = leerFoto_() || { tareas: [], reuniones: [] };
  var r = difResumen_(antes, ahora);
  var hoy = Utilities.formatDate(new Date(), RESUMEN_ZONA, "dd/MM/yyyy");
  MailApp.sendEmail({
    to: RESUMEN_PRUEBA,
    subject: "[Prueba] WEBER I Tablero de trabajo: Resumen del día " + hoy + " (" + r.total + (r.total === 1 ? " cambio)" : " cambios)"),
    htmlBody: htmlResumen_(r, hoy, true),
    name: "WEBER I Tablero de trabajo"
  });
  return { enviadoA: RESUMEN_PRUEBA, cambios: r.total, cuotaRestante: MailApp.getRemainingDailyQuota() };
}

/* Crea (o recrea) el activador diario y arranca desde la foto de ahora. */
function activarResumen_() {
  desactivarResumen_();
  ScriptApp.newTrigger("enviarResumenDiario").timeBased().everyDays(1).atHour(RESUMEN_HORA).inTimezone(RESUMEN_ZONA).create();
  guardarFoto_(fotoActual_());
  return { activado: true, hora: RESUMEN_HORA, zona: RESUMEN_ZONA, destinatarios: RESUMEN_DESTINATARIOS };
}

function desactivarResumen_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === "enviarResumenDiario") { ScriptApp.deleteTrigger(tr); n++; }
  });
  return { desactivados: n };
}

/* =====================================================================
   DESTRABAR LOS DESPLEGABLES DE RESPONSABLE Y ÁREA
   Vienen en modo "Rechazar": si la web manda "Eve, Ger" en RESPONSABLE, o un
   área que todavía no está en la lista, la celda lo rechaza y el valor
   termina anotado en OBSERVACIONES. Esto deja las mismas listas pero en modo
   aviso, así una celda puede llevar varios nombres y se puede sumar un área
   nueva. No borra ni cambia ningún dato: sólo deja de bloquear.
   EJECUTAR A MANO UNA SOLA VEZ: elegir "permitirVariosResponsables" en el
   desplegable del editor y Ejecutar.
   ===================================================================== */
function permitirVariosResponsables() {
  function limpio_(v) {
    return String(v === null || v === undefined ? "" : v).trim().toUpperCase()
      .replace(/[\u00C1\u00C0\u00C4\u00C2]/g, "A").replace(/[\u00C9\u00C8\u00CB\u00CA]/g, "E")
      .replace(/[\u00CD\u00CC\u00CF\u00CE]/g, "I").replace(/[\u00D3\u00D2\u00D6\u00D4]/g, "O")
      .replace(/[\u00DA\u00D9\u00DC\u00DB]/g, "U");
  }
  var out = [];
  ss_().getSheets().forEach(function (sheet) {
    var nombre = sheet.getName();
    if (nombre.indexOf("WebApp") === 0) return;
    var values = sheet.getDataRange().getValues();
    var fila = -1, columnas = {};
    for (var r = 0; r < values.length && fila === -1; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cab = limpio_(values[r][c]);
        if (cab === "RESPONSABLE" || cab === "AREA") {
          if (columnas[cab] === undefined) columnas[cab] = c;
          fila = r;
        }
      }
    }
    if (fila === -1) return;
    var desde = fila + 2;
    var filas = sheet.getMaxRows() - desde + 1;
    if (filas < 1) { out.push({ hoja: nombre, estado: "sin filas" }); return; }
    Object.keys(columnas).forEach(function (campo) {
    var col = columnas[campo];
    var rango = sheet.getRange(desde, col + 1, filas, 1);
    var vistos = {}, nombres = [];
    function sumar(v) {
      var x = String(v === null || v === undefined ? "" : v).trim();
      if (!x || vistos[x.toUpperCase()]) return;
      vistos[x.toUpperCase()] = true;
      nombres.push(x);
    }
    rango.getDataValidations().forEach(function (f) {
      var dv = f[0];
      if (!dv) return;
      try {
        var tipo = dv.getCriteriaType(), crit = dv.getCriteriaValues();
        if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          crit[0].forEach(sumar);
        } else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          crit[0].getValues().forEach(function (f2) { sumar(f2[0]); });
        }
      } catch (e) {}
    });
    // Lo ya escrito también entra en la lista, partido por coma.
    rango.getValues().forEach(function (f) {
      String(f[0] === null || f[0] === undefined ? "" : f[0]).split(/\s*[,\/]\s*/).forEach(sumar);
    });
    if (!nombres.length) { out.push({ hoja: nombre, campo: campo, estado: "sin valores" }); return; }
    nombres.sort(function (a, b) { return a.localeCompare(b); });
    rango.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true).setAllowInvalid(true).build());
    out.push({ hoja: nombre, campo: campo, columna: col + 1, desdeFila: desde, valores: nombres });
    });
  });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/* =====================================================================
   QUITAR UN NOMBRE DE LA LISTA DEL DESPLEGABLE
   Saca un valor de la lista que ofrece RESPONSABLE (o ÁREA) sin tocar una
   sola celda escrita: la lista es lo que el desplegable sugiere, no el dato.
   Lo llama la web desde la ficha de la tarea.
   ===================================================================== */
function quitarOpcion_(campo, valor) {
  function limpio_(v) {
    return String(v === null || v === undefined ? "" : v).trim().toUpperCase()
      .replace(/[\u00C1\u00C0\u00C4\u00C2]/g, "A").replace(/[\u00C9\u00C8\u00CB\u00CA]/g, "E")
      .replace(/[\u00CD\u00CC\u00CF\u00CE]/g, "I").replace(/[\u00D3\u00D2\u00D6\u00D4]/g, "O")
      .replace(/[\u00DA\u00D9\u00DC\u00DB]/g, "U");
  }
  var buscado = limpio_(campo || "RESPONSABLE");
  var quitar = String(valor === null || valor === undefined ? "" : valor).trim();
  if (!quitar) throw new Error("No me dijiste qué valor quitar.");
  var objetivo = limpio_(quitar);
  var hojas = [], enUso = 0;

  ss_().getSheets().forEach(function (sheet) {
    var nombre = sheet.getName();
    if (nombre.indexOf("WebApp") === 0) return;
    var values = sheet.getDataRange().getValues();
    var fila = -1, col = -1;
    for (var r = 0; r < values.length && fila === -1; r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (limpio_(values[r][c]) === buscado) { fila = r; col = c; break; }
      }
    }
    if (fila === -1) return;
    var desde = fila + 2;
    var filas = sheet.getMaxRows() - desde + 1;
    if (filas < 1) return;
    var rango = sheet.getRange(desde, col + 1, filas, 1);

    /* Si alguna celda todavía lo tiene escrito, no se saca de la lista: el
       desplegable dejaría de explicar un valor que está ahí a la vista. */
    rango.getValues().forEach(function (f) {
      String(f[0] === null || f[0] === undefined ? "" : f[0]).split(/\s*[,\/]\s*/).forEach(function (x) {
        if (limpio_(x) === objetivo) enUso++;
      });
    });

    var vistos = {}, quedan = [];
    function sumar(v) {
      var x = String(v === null || v === undefined ? "" : v).trim();
      if (!x || limpio_(x) === objetivo || vistos[x.toUpperCase()]) return;
      vistos[x.toUpperCase()] = true;
      quedan.push(x);
    }
    rango.getDataValidations().forEach(function (f) {
      var dv = f[0];
      if (!dv) return;
      try {
        var tipo = dv.getCriteriaType(), crit = dv.getCriteriaValues();
        if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          crit[0].forEach(sumar);
        } else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          crit[0].getValues().forEach(function (f2) { sumar(f2[0]); });
        }
      } catch (e) {}
    });
    if (!quedan.length) return;
    hojas.push({ hoja: nombre, rango: rango, lista: quedan });
  });

  if (enUso) {
    throw new Error("No puedo quitar \"" + quitar + "\" de la lista: todavía está escrito en " +
      enUso + (enUso === 1 ? " celda." : " celdas."));
  }
  hojas.forEach(function (h) {
    h.rango.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(h.lista, true).setAllowInvalid(true).build());
  });
  return { quitado: quitar, campo: buscado, hojas: hojas.map(function (h) { return h.hoja; }) };
}

/* =====================================================================
   OBJETIVO Y PRIORIDADES, EDITABLES DESDE LA WEB
   El bloque vive arriba de la hoja de proceso: OBJETIVO 1 con su respuesta
   dos filas abajo, y PRIORIDADES con una fila numerada por prioridad. El
   título y la descripción comparten celda, separados por un tab, que es como
   readSeed_ los viene leyendo desde siempre.
   ===================================================================== */
function filaEtiqueta_(values, texto) {
  var target = String(texto).toUpperCase();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][1] || "").trim().toUpperCase() === target) return i;
  }
  return -1;
}

/* Las filas numeradas del bloque, con la columna donde está su texto. */
function bloquePrioridades_() {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  var values = sheet.getDataRange().getValues();
  var inicio = filaEtiqueta_(values, "PRIORIDADES");
  var fin = filaEtiqueta_(values, "INICIATIVAS");
  if (inicio === -1 || fin === -1) {
    throw new Error("No encuentro PRIORIDADES / INICIATIVAS en '" + SHEET_PROCESO + "'");
  }
  var filas = [];
  for (var r = inicio + 2; r < fin; r++) {
    var n = cell_(values[r], 1);
    if (n === null || isNaN(Number(n))) continue;
    var col = 2;
    while (col < 6 && cell_(values[r], col) === null) col++;
    if (col >= 6) col = 2;
    filas.push({ fila: r, n: Number(n), col: col });
  }
  return { sheet: sheet, values: values, inicio: inicio, fin: fin, filas: filas };
}

function setObjetivo_(texto) {
  var sheet = ss_().getSheetByName(SHEET_PROCESO);
  var values = sheet.getDataRange().getValues();
  var fila = filaEtiqueta_(values, "OBJETIVO 1");
  if (fila === -1) throw new Error("No encuentro el bloque OBJETIVO 1 en '" + SHEET_PROCESO + "'");
  sheet.getRange(fila + 3, 2).setValue(String(texto === null || texto === undefined ? "" : texto));
  return { objetivo: texto || "" };
}

function setPrioridad_(n, texto, desc) {
  var num = Number(n);
  var B = bloquePrioridades_();
  var fila = null;
  B.filas.forEach(function (f) { if (f.n === num) fila = f; });
  if (!fila) throw new Error("No encuentro la prioridad " + num + " en el bloque PRIORIDADES");
  var partes = [String(texto === null || texto === undefined ? "" : texto).trim()];
  var d = (desc === null || desc === undefined) ? "" : String(desc).trim();
  if (d) partes.push(d);
  B.sheet.getRange(fila.fila + 1, fila.col + 1).setValue(partes.join("\t"));
  return { n: num, texto: partes[0], desc: d };
}

/* Suma una prioridad al final del bloque: copia la fila anterior con su
   formato y sus celdas combinadas, y suma la etiqueta nueva al desplegable
   de la columna PRIORIDAD. */
function agregarPrioridad_(texto, desc) {
  var B = bloquePrioridades_();
  if (!B.filas.length) throw new Error("No encuentro ninguna prioridad para copiar el formato");
  var ultima = B.filas[B.filas.length - 1];
  B.filas.forEach(function (f) { if (f.n > ultima.n) ultima = f; });
  if (ultima.n >= 9) throw new Error("Ya hay 9 prioridades, que es el máximo que maneja el tablero");

  var filaOrigen = ultima.fila + 1;
  var filaNueva = filaOrigen + 1;
  var ancho = Math.max(B.sheet.getLastColumn(), ultima.col + 1);
  B.sheet.insertRowAfter(filaOrigen);
  B.sheet.getRange(filaOrigen, 1, 1, ancho).copyTo(B.sheet.getRange(filaNueva, 1, 1, ancho), { formatOnly: true });
  B.sheet.getRange(filaOrigen, 1, 1, ancho).getMergedRanges().forEach(function (m) {
    if (m.getNumRows() !== 1) return;
    B.sheet.getRange(filaNueva, m.getColumn(), 1, m.getNumColumns()).merge();
  });

  var partes = [String(texto === null || texto === undefined ? "" : texto).trim()];
  var d = (desc === null || desc === undefined) ? "" : String(desc).trim();
  if (d) partes.push(d);
  B.sheet.getRange(filaNueva, 2).setValue(ultima.n + 1);
  B.sheet.getRange(filaNueva, ultima.col + 1).setValue(partes.join("\t"));

  /* La columna PRIORIDAD tiene que ofrecer la etiqueta nueva. */
  var etiquetas = [];
  for (var i = 1; i <= ultima.n + 1; i++) etiquetas.push("P" + i);
  var values = B.sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  if (C.prioridad !== undefined) {
    var desde = C.header + 2;
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(etiquetas, true).setAllowInvalid(true).build();
    B.sheet.getRange(desde, C.prioridad + 1, Math.max(B.sheet.getMaxRows() - desde + 1, 1), 1)
      .setDataValidation(regla);
  }
  return { n: ultima.n + 1, texto: partes[0], fila: filaNueva, opciones: etiquetas };
}

/* =====================================================================
   CALENDARIO · lo que ya está agendado en Google Calendar
   Una vez por hora, las reuniones de Weber que están en el calendario
   de trabajo entran solas en la pestaña de Agenda, con su link de Meet.
   Es de una sola mano: manda el calendario. Lo que escribas en el tablero
   y el calendario no tenga (un tema propio, por ejemplo) no se pisa.

   PARA DEJARLO ANDANDO, UNA SOLA VEZ:
   1. En el editor, "Servicios" (+) → "Google Calendar API" → Agregar.
   2. Elegir "conectarCalendario" en el desplegable y Ejecutar. Pide
      permiso para ver el calendario: aceptar.
   Después corre solo, cada hora.
   ===================================================================== */

/* El calendario de donde salen las reuniones. */
var CAL_ID = "hola@soyevelyna.com";
/* Cuáles de ese calendario son de este tablero: se mira el título. */
var CAL_FILTRO = /\bweber\b/i;
/* Del título sobra el nombre del cliente: adentro del tablero ya se sabe
   de quién es, así que "Weber I Reunión de proceso" queda en
   "Reunión de proceso". */
var CAL_PREFIJO = /^\s*weber\s*(?:[I|·\-–:]\s*)?/i;
/* El tablero lo lee el equipo, que está en Argentina. El calendario está en
   hora de Europa: la hora se guarda convertida, no como se ve allá. */
var CAL_ZONA = "America/Argentina/Buenos_Aires";
/* Hasta dónde mirar para adelante. */
var CAL_DIAS = 120;
/* Los invitados llegan como mails; en el tablero van por su nombre. Un mail
   que no esté acá no se escribe: la reunión igual se carga, y el nombre se
   puede sumar a mano desde el tablero (o acá, que es para siempre). */
var CAL_NOMBRES = {
  "hola@soyevelyna.com": "Eve",
  "abrahanevelyn@gmail.com": "Eve",
  "marketingdigital.push@gmail.com": "Fer",
  "gweber@webersi.com.ar": "Ger",
  "juliana@loopa.io": "Loopa",
};

/* Para correr a mano desde el editor. */
function conectarCalendario() {
  var viejos = ScriptApp.getProjectTriggers();
  for (var i = 0; i < viejos.length; i++) {
    if (viejos[i].getHandlerFunction() === "sincronizarCalendario") {
      ScriptApp.deleteTrigger(viejos[i]);
    }
  }
  ScriptApp.newTrigger("sincronizarCalendario").timeBased().everyHours(1).create();
  var r = sincronizarCalendario();
  Logger.log("Listo, se sincroniza cada hora. Ahora: " + JSON.stringify(r));
  return r;
}
function sincronizarCalendarioAhora() {
  Logger.log(JSON.stringify(sincronizarCalendario()));
}

function sincronizarCalendario() {
  ensureSheets_();
  var desde = new Date();
  desde.setHours(0, 0, 0, 0);
  var hasta = new Date(desde.getTime() + CAL_DIAS * 24 * 3600 * 1000);

  var items = eventosDelCalendario_(desde, hasta);

  var sheet = ss_().getSheetByName(SHEET_AGENDA);
  var cols = SHEET_SCHEMAS[SHEET_AGENDA];
  var iId = cols.indexOf("id"), iFecha = cols.indexOf("fecha"), iHora = cols.indexOf("hora");
  var iTema = cols.indexOf("tema"), iConv = cols.indexOf("convocados"), iLink = cols.indexOf("link");
  var iPor = cols.indexOf("creadoPor"), iAlta = cols.indexOf("createdAt");
  var filas = sheet.getDataRange().getValues();

  /* Dos maneras de encontrar una reunión que ya está: por su id del
     calendario, y por fecha+hora, que es cómo se reconocen las que
     cargaste a mano antes de que esto existiera. */
  var porId = {}, porCuando = {};
  for (var r = 1; r < filas.length; r++) {
    var id = String(filas[r][iId] || "").trim();
    if (!id) continue;
    porId[id] = r;
    var clave = calFecha_(filas[r][iFecha]) + " " + calHora_(filas[r][iHora]);
    if (!porCuando[clave]) porCuando[clave] = r;
  }

  var vistos = {}, nuevas = [], res = { creadas: 0, actualizadas: 0, adoptadas: 0, borradas: 0 };

  for (var k = 0; k < items.length; k++) {
    var ev = items[k];
    var id = "cal_" + ev.id;
    vistos[id] = true;
    var fecha = ev.fecha, hora = ev.hora;
    var fila = porId[id];
    var adoptada = false;
    if (fila === undefined) {
      fila = porCuando[fecha + " " + hora];
      /* Sólo se adopta una fila que nadie del calendario reclamó todavía. */
      if (fila !== undefined && String(filas[fila][iId] || "").indexOf("cal_") === 0) {
        fila = undefined;
      }
      adoptada = fila !== undefined;
    }

    if (fila === undefined) {
      var alta = [];
      for (var c = 0; c < cols.length; c++) alta.push("");
      alta[iId] = id;
      alta[iFecha] = fecha;
      alta[iHora] = hora;
      alta[iTema] = ev.tema;
      alta[iConv] = ev.convocados;
      alta[iLink] = ev.link;
      alta[iPor] = "Google Calendar";
      alta[iAlta] = nowIso_();
      nuevas.push(alta);
      res.creadas++;
      continue;
    }

    /* Del calendario manda cuándo es y por dónde se entra. El tema y los
       convocados sólo se completan si están vacíos: lo que escribiste vos
       vale más que el título del evento. */
    var cambio = false;
    cambio = escribirSi_(sheet, fila, iId, filas[fila][iId], id) || cambio;
    cambio = escribirSi_(sheet, fila, iFecha, calFecha_(filas[fila][iFecha]), fecha) || cambio;
    cambio = escribirSi_(sheet, fila, iHora, calHora_(filas[fila][iHora]), hora) || cambio;
    if (ev.link) cambio = escribirSi_(sheet, fila, iLink, filas[fila][iLink], ev.link) || cambio;
    if (!String(filas[fila][iTema] || "").trim() && ev.tema) {
      cambio = escribirSi_(sheet, fila, iTema, filas[fila][iTema], ev.tema) || cambio;
    }
    if (!String(filas[fila][iConv] || "").trim() && ev.convocados) {
      cambio = escribirSi_(sheet, fila, iConv, filas[fila][iConv], ev.convocados) || cambio;
    }
    escribirSi_(sheet, fila, iPor, filas[fila][iPor], "Google Calendar");
    if (adoptada) res.adoptadas++;
    else if (cambio) res.actualizadas++;
  }

  if (nuevas.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevas.length, cols.length).setValues(nuevas);
  }

  /* Una reunión que el calendario ya no tiene y que había traído el
     calendario, se va. Las cargadas a mano no se tocan nunca. */
  var hastaIso = Utilities.formatDate(hasta, CAL_ZONA, "yyyy-MM-dd");
  var desdeIso = Utilities.formatDate(desde, CAL_ZONA, "yyyy-MM-dd");
  for (var r2 = filas.length - 1; r2 >= 1; r2--) {
    var idv = String(filas[r2][iId] || "").trim();
    if (idv.indexOf("cal_") !== 0 || vistos[idv]) continue;
    var f = calFecha_(filas[r2][iFecha]);
    if (f >= desdeIso && f <= hastaIso) { sheet.deleteRow(r2 + 1); res.borradas++; }
  }
  return res;
}

/* Se usa el servicio avanzado de Calendar (no CalendarApp) porque es el
   único que trae el link de Meet y los invitados. */
function eventosDelCalendario_(desde, hasta) {
  if (typeof Calendar === "undefined" || !Calendar.Events) {
    throw new Error("Falta activar el servicio 'Google Calendar API' en el editor (Servicios +).");
  }
  var resp = Calendar.Events.list(CAL_ID, {
    timeMin: desde.toISOString(),
    timeMax: hasta.toISOString(),
    singleEvents: true,
    orderBy: "startTime",
    maxResults: 250,
    showDeleted: false
  });
  var items = resp.items || [], out = [];
  for (var i = 0; i < items.length; i++) {
    var ev = items[i];
    if (ev.status === "cancelled") continue;
    var titulo = String(ev.summary || "");
    if (!CAL_FILTRO.test(titulo)) continue;
    var link = calLink_(ev);
    var invitados = ev.attendees || [];
    /* Un recordatorio suelto ("Factura | Noctis") no es una reunión: para
       entrar tiene que tener link o gente invitada. */
    if (!link && !invitados.length) continue;
    var ini = ev.start && ev.start.dateTime
      ? new Date(ev.start.dateTime)
      : (ev.start && ev.start.date ? new Date(ev.start.date + "T12:00:00Z") : null);
    if (!ini) continue;
    out.push({
      id: ev.id,
      fecha: Utilities.formatDate(ini, CAL_ZONA, "yyyy-MM-dd"),
      hora: (ev.start && ev.start.dateTime) ? Utilities.formatDate(ini, CAL_ZONA, "HH:mm") : "",
      tema: titulo.replace(CAL_PREFIJO, "").trim() || titulo.trim(),
      convocados: calConvocados_(invitados),
      link: link
    });
  }
  return out;
}

function calLink_(ev) {
  if (ev.hangoutLink) return ev.hangoutLink;
  var eps = (ev.conferenceData && ev.conferenceData.entryPoints) || [];
  for (var i = 0; i < eps.length; i++) {
    if (eps[i].entryPointType === "video" && eps[i].uri) return eps[i].uri;
  }
  if (ev.location && /^https?:\/\//i.test(String(ev.location))) return String(ev.location).trim();
  return "";
}

function calConvocados_(invitados) {
  var out = [];
  for (var i = 0; i < invitados.length; i++) {
    var mail = String(invitados[i].email || "").toLowerCase().trim();
    var nombre = CAL_NOMBRES[mail];
    if (nombre && out.indexOf(nombre) === -1) out.push(nombre);
  }
  return out.join(", ");
}

/* La zona horaria del Sheet, no la del proyecto. Una hora guardada como
   hora (no como texto) la arma Sheets en SU zona; leerla en otra la corre.
   Eso hacía que Weber mostrara 18:16 donde la hoja dice 13:00. */
function tzHoja_() { return ss_().getSpreadsheetTimeZone(); }

function calFecha_(v) { return toIsoDate_(v) || ""; }

function calHora_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, tzHoja_(), "HH:mm");
  }
  return String(v === null || v === undefined ? "" : v).trim();
}

/* Escribe sólo si cambia: así el Sheet no registra ediciones de más. */
function escribirSi_(sheet, fila, col, actual, valor) {
  var a = Object.prototype.toString.call(actual) === "[object Date]" ? actual : String(actual === null || actual === undefined ? "" : actual).trim();
  if (a === valor) return false;
  sheet.getRange(fila + 1, col + 1).setValue(valor);
  return true;
}
