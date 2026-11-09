/**
 * Forja - backend en Google Apps Script.
 *
 * Se pega en: planilla "Forja - Datos" -> Extensiones -> Apps Script.
 * Antes de publicar, definir la propiedad de script TOKEN
 * (Configuración del proyecto -> Propiedades de la secuencia de comandos).
 *
 * Pestañas esperadas:
 *   Sesiones: sesion_id | fecha | rutina | duracion_seg | volumen | ejercicio |
 *             patron | serie_nro | peso | reps | nota | rutina_id
 *   Estado:   clave | valor   (B2 guarda el JSON de rutinas)
 */

var SESSIONS_SHEET = 'Sesiones';
var STATE_SHEET = 'Estado';
var NCOLS = 12;

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || req.token !== token) throw new Error('unauthorized');
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (req.action === 'pull') out = pull_();
      else if (req.action === 'push') { push_(req.data); out = { ok: true }; }
      else throw new Error('bad action');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Útil para comprobar en el navegador que la Web app está viva.
function doGet() {
  return ContentService.createTextOutput('Forja OK');
}

function pull_() {
  var ss = SpreadsheetApp.getActive();
  var stateRaw = ss.getSheetByName(STATE_SHEET).getRange('B2').getValue();
  var meta = stateRaw ? JSON.parse(stateRaw) : null;

  var sh = ss.getSheetByName(SESSIONS_SHEET);
  var last = sh.getLastRow();
  var sessions = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, NCOLS).getValues();
    var byId = {};
    rows.forEach(function (r) {
      var id = String(r[0]);
      if (!id) return;
      var s = byId[id];
      if (!s) {
        var d = r[1];
        s = byId[id] = {
          id: id,
          date: d instanceof Date ? d.getTime() : Number(d) || 0,
          routineName: String(r[2]),
          duration: Number(r[3]) || 0,
          volume: Number(r[4]) || 0,
          routineId: String(r[11] || ''),
          totalSets: 0,
          exercises: [],
          _ex: {}
        };
        sessions.push(s);
      }
      var name = String(r[5]);
      var ex = s._ex[name];
      if (!ex) {
        ex = s._ex[name] = { name: name, pattern: String(r[6]), sets: [], note: String(r[10]) };
        s.exercises.push(ex);
      }
      var reps = r[9] === '' ? '' : String(r[9]);
      var weight = r[8] === '' ? '' : String(r[8]);
      ex.sets.push({ weight: weight, reps: reps, done: reps !== '' || weight !== '' });
      if (parseFloat(reps) > 0) s.totalSets++;
    });
    sessions.forEach(function (s) { delete s._ex; });
  }
  return { meta: meta, sessions: sessions };
}

function push_(data) {
  var ss = SpreadsheetApp.getActive();

  ss.getSheetByName(STATE_SHEET).getRange('B2').setValue(JSON.stringify({
    routines: data.routines,
    defaultRest: data.defaultRest,
    updatedAt: data.updatedAt
  }));

  var rows = [];
  (data.sessions || []).forEach(function (s) {
    (s.exercises || []).forEach(function (ex) {
      (ex.sets || []).forEach(function (st, i) {
        rows.push([
          s.id, new Date(s.date), s.routineName, s.duration, s.volume,
          ex.name, ex.pattern || '', i + 1,
          st.weight === '' || st.weight == null ? '' : st.weight,
          st.reps === '' || st.reps == null ? '' : st.reps,
          ex.note || '', s.routineId || ''
        ]);
      });
    });
  });

  var sh = ss.getSheetByName(SESSIONS_SHEET);
  sh.getRange(1, 12).setValue('rutina_id');
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, NCOLS).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, NCOLS).setValues(rows);
}
