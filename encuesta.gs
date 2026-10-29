// v2026-10-28 — agrega dashboard con resúmenes Gemini
var SHEET_ID   = '1zufm3lNJxfpk1IeqVYlHzOMe3zZO7TjoaZ94K13ymyw';
var SHEET_NAME = 'Respuestas';
var GEMINI_KEY = 'TU_API_KEY_ACA'; // reemplazá con tu key de Google AI Studio

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
  var p = e.parameter;

  if (p.action === 'dashboard') {
    return getDashboard();
  }

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

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

function getDashboard() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var rows  = sheet.getDataRange().getValues();

  if (rows.length <= 1) {
    return jsonResponse({ total: 0, rows: [], resumenes: {} });
  }

  var headers = rows[0];
  var data = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] || ''; });
    return obj;
  });

  // Campos de texto libre a resumir
  var camposTexto = {
    'Primera reacción':        '¿Cuál fue tu primera reacción al conocer el concepto de Simpleno?',
    'Q4 — Por qué te las guardás': '¿Por qué te guardás cosas que te pasaron?',
    'Q7 — Qué cambia anónimo': '¿Qué cambiaría para vos si pudieras contar algo de forma anónima?',
    'El nombre — Simpleno':    '¿Qué te inspira o dice el nombre Simpleno?',
    'Cierre':                  '¿Qué más querías decir sobre Simpleno?'
  };

  var resumenes = {};
  Object.keys(camposTexto).forEach(function(campo) {
    var textos = data
      .map(function(r) { return (r[campo] || '').toString().trim(); })
      .filter(function(t) { return t.length > 2; });

    if (textos.length > 0) {
      resumenes[campo] = resumirConGemini(textos, camposTexto[campo]);
    }
  });

  return jsonResponse({
    total: data.length,
    ultima: data[data.length - 1]['Timestamp'],
    rows: data,
    resumenes: resumenes
  });
}

function resumirConGemini(textos, pregunta) {
  try {
    var prompt = 'Tenés estas respuestas de personas reales a la pregunta: "' + pregunta + '".\n\n'
      + textos.map(function(t, i) { return (i + 1) + '. ' + t; }).join('\n')
      + '\n\nResumí los patrones principales en 2-3 líneas, en español, sin listar las respuestas una por una. Sé directo y concreto.';

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + GEMINI_KEY;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    return json.candidates[0].content.parts[0].text;
  } catch(err) {
    return '(resumen no disponible)';
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
