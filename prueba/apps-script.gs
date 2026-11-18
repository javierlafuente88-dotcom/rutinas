/**
 * Recibe el progreso que mandan los alumnos y lo guarda en la planilla "Seguimiento de alumnos".
 * Se pega UNA vez en Extensiones > Apps Script de esa planilla y se publica como aplicación web.
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // evita que dos envíos simultáneos se pisen
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.accion === "guardarRutina") return guardarRutina(ss, d);   // rutinas con link corto
    const registro = ss.getSheetByName("Registro");
    const ahora = new Date();
    const alumno = String(d.alumno || "").trim() || "(sin nombre)";
    const rutina = String(d.rutina || "").trim();

    (d.ejercicios || []).forEach(function (ej) {
      const nombre = String(ej.nombre || "").trim();
      if (!nombre) return;
      const kg = String(ej.kg == null ? "" : ej.kg).replace(",", ".").trim();
      const fila = registro.getLastRow() + 1;
      registro.getRange(fila, 1, 1, 7).setValues([[
        ahora, alumno, rutina, nombre,
        Number(ej.hechas) || 0, Number(ej.pedidas) || 0,
        kg === "" || isNaN(Number(kg)) ? "" : Number(kg)
      ]]);
      registro.getRange(fila, 8, 1, 3).setFormulas([[
        '=B' + fila + '&"|"&D' + fila,
        '=COUNTIF(H$2:H' + fila + ',H' + fila + ')',
        '=H' + fila + '&"|"&I' + fila
      ]]);
      agregarSiFalta(ss.getSheetByName("Alumnos"), alumno);
      agregarSiFalta(ss.getSheetByName("Ejercicios"), nombre);
    });
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Sin parámetros: sirve para comprobar que la publicación funciona.
// Con ?id=xxxxxx: devuelve la rutina guardada con ese código (lo usa el link corto del alumno).
function doGet(e) {
  const id = e && e.parameter && e.parameter.id;
  if (!id) return ContentService.createTextOutput("La planilla está recibiendo datos correctamente.");
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rutinas");
  if (!hoja) return salida({ ok: false });
  const celda = hoja.getRange("A:A").createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!celda) return salida({ ok: false });
  return salida({ ok: true, data: JSON.parse(hoja.getRange(celda.getRow(), 5).getValue()) });
}

// Guarda una rutina con un código corto. La hoja "Rutinas" se crea sola la primera vez.
function guardarRutina(ss, d) {
  if (!/^[a-z0-9]{4,12}$/.test(String(d.id || ""))) return salida({ ok: false, error: "código inválido" });
  let hoja = ss.getSheetByName("Rutinas");
  if (!hoja) {
    hoja = ss.insertSheet("Rutinas");
    hoja.appendRow(["Código", "Fecha", "Alumno", "Rutina", "Datos"]);
    hoja.setFrozenRows(1);
  }
  hoja.appendRow([d.id, new Date(), String(d.alumno || ""), String(d.rutina || ""), JSON.stringify(d.datos || {})]);
  return salida({ ok: true });
}

function agregarSiFalta(hoja, valor) {
  const col = hoja.getRange("A2:A201").getValues().map(function (f) { return f[0]; });
  if (col.indexOf(valor) !== -1) return;
  const libre = col.indexOf("");
  if (libre !== -1) hoja.getRange(libre + 2, 1).setValue(valor);
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
