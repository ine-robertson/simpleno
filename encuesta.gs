// v2026-11-11 — agrega FODA generado por Gemini en getDashboard
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
  'QIA — La IA como oyente', 'IA nunca — Por qué no', 'IA sí — Por qué sí',
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
    p.qia           || '',
    p.iaNunca       || '',
    p.iaSi          || '',
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
    resumenes: resumenes,
    foda: generarFODA(data)
  });
}

function generarFODA(data) {
  var total = data.length;
  if (total === 0) return null;

  var SEM = { total: 'Totalmente', aveces: 'A veces', puedeser: 'Puede ser', nada: 'Para nada' };
  function normSem(v) { return SEM[v] || v; }

  function pctCampo(campo, valor) {
    var n = data.filter(function(r) { return normSem((r[campo]||'').trim()) === valor; }).length;
    return n + '/' + total + ' (' + Math.round(n/total*100) + '%)';
  }
  function topMulti(campo, top) {
    var c = {};
    data.forEach(function(r) {
      (r[campo]||'').split(' | ').forEach(function(v) {
        v = v.trim(); if (v) c[v] = (c[v]||0) + 1;
      });
    });
    return Object.keys(c).sort(function(a,b){return c[b]-c[a];}).slice(0,top)
      .map(function(k){ return k + ' (' + c[k] + ')'; }).join(' / ');
  }

  var stats = 'Total encuestados: ' + total + '\n'
    + 'Tienen cosas sin contar: ' + data.filter(function(r){ return (r['Q3 — Te guardás historias']||'').indexOf('Hay') !== -1 || (r['Q3 — Te guardás historias']||'').indexOf('Generalmente') !== -1; }).length + '/' + total + '\n'
    + 'Las contarían anónimas (totalmente): ' + pctCampo('Q1 — Historias anónimas', 'Totalmente') + '\n'
    + 'Las contarían anónimas (a veces): ' + pctCampo('Q1 — Historias anónimas', 'A veces') + '\n'
    + 'Escribir ayuda a soltar (totalmente): ' + pctCampo('Qsoltar — Escribir ayuda a soltar', 'Totalmente') + '\n'
    + 'Compartirían si aporta a otros (totalmente): ' + pctCampo('Q6 — Ayudar a otros', 'Totalmente — si aporta o habilita a otros, vale la pena') + '\n'
    + 'Leerían historias de desconocidos (totalmente): ' + pctCampo('Lector — Leerías', 'Totalmente') + '\n'
    + 'No les importa saber quién escribió: ' + pctCampo('Q8 — Importa saber quién', 'No me importa') + '\n'
    + 'Le contaron algo a una IA: ' + data.filter(function(r){ return (r['QIA — La IA como oyente']||'').indexOf('Sí') !== -1; }).length + '/' + total + '\n'
    + 'Frenos al compartir (top opciones): ' + topMulti('Q10 — Frenos al compartir', 5) + '\n'
    + 'Motivadores de lectura (top opciones): ' + topMulti('Motivar — Opciones', 4);

  var prompt = 'Sos estratega de producto. Simpleno es una plataforma donde las personas publican historias/experiencias de forma completamente anónima. No hay autor visible, no hay comentarios, no hay likes. Las únicas acciones son "esto resuena conmigo" y "te entiendo". El propósito es la trascendencia: que lo que viví ayude a otros (Viktor Frankl).\n\n'
    + 'Datos de encuesta de validación:\n' + stats + '\n\n'
    + 'Generá un análisis FODA del concepto basado en estos datos. Respondé SOLO con un JSON válido, sin texto extra:\n'
    + '{"fortalezas":["...","..."],"oportunidades":["...","..."],"debilidades":["...","..."],"amenazas":["...","..."]}\n'
    + 'Máximo 4 items por cuadrante. Breves y accionables, en español.';

  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + GEMINI_KEY;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true,
      deadline: 30
    });
    var json = JSON.parse(res.getContentText());
    var text = json.candidates[0].content.parts[0].text.trim()
      .replace(/^```json\s*/,'').replace(/^```\s*/,'').replace(/\s*```$/,'');
    return JSON.parse(text);
  } catch(err) {
    return null;
  }
}

function resumirConGemini(textos, pregunta) {
  try {
    var prompt = 'Tenés estas respuestas de personas reales a la pregunta: "' + pregunta + '".\n\n'
      + textos.map(function(t, i) { return (i + 1) + '. ' + t; }).join('\n')
      + '\n\nResumí los patrones principales en 2-3 líneas, en español, sin listar las respuestas una por una. Sé directo y concreto.';

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + GEMINI_KEY;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true,
      deadline: 30
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
