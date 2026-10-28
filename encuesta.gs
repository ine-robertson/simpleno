// v2026-10-28 — agrega campo qnombre (pregunta sobre el nombre Simpleno)
var SHEET_ID   = '1zufm3lNJxfpk1IeqVYlHzOMe3zZO7TjoaZ94K13ymyw';
var SHEET_NAME = 'Respuestas';

var HEADERS = [
  'Timestamp', 'Perfil', 'Edad', 'Primera reacción',
  'Q1 — Historias anónimas', 'Q2 — Por qué no con nombre',
  'Q3 — Te guardás historias', 'Qsoltar — Escribir ayuda a soltar',
  'Q4 — Por qué te las guardás', 'Q5 — Si hay evidencia lo harías',
  'Q6 — Ayudar a otros', 'Q10 — Frenos al compartir',
  'Lector — Leerías', 'Lector2 — Alguna vez te movilizó',
  'Motivar — Opciones', 'Motivar — Otro',
  'Abandonar — Opciones', 'Abandonar — Otro',
  'Q7 — Qué cambia anónimo', 'Q8 — Importa saber quién',
  'Q11 — Feedback vs resonancia',
  'Perfil específico', 'El nombre — Simpleno', 'Cierre'
];

function doGet(e) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  var p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.perfil        || '',
    p.edad          || '',
    p.q9            || '',
    p.q1            || '',
    p.q2            || '',
    p.q3            || '',
    p.qsoltar       || '',
    p.q4            || '',
    p.q5            || '',
    p.q6            || '',
    p.q10           || '',
    p.qlector       || '',
    p.qlector2      || '',
    p.motivar       || '',
    p.motivarOtro   || '',
    p.abandonar     || '',
    p.abandonarOtro || '',
    p.q7            || '',
    p.q8            || '',
    p.q11           || '',
    p.qpe           || '',
    p.qnombre       || '',
    p.qcierre       || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
