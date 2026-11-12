// v2026-11-12i — gemini-3.8-flash sin thinking (thinkingBudget:0), fix parsing parts
var SHEET_ID   = '1zufm3lNJxfpk1IeqVYlHzOMe3zZO7TjoaZ94K13ymyw';
var SHEET_NAME = 'Respuestas';
var GEMINI_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || '';

var HEADERS = [
  'Timestamp',
  'Profesión',
  'Edad',
  'Primera reacción al concepto',
  '¿Las contarías anónimas?',
  '¿Por qué no con tu nombre?',
  '¿Tenés cosas que no contaste?',
  '¿Escribir ayuda a soltarlo o procesarlo?',
  '¿Por qué te las guardás?',
  '¿Si hay evidencia de que es terapéutico, lo harías?',
  '¿Lo harías para ayudar a otros aunque no te sume a vos?',
  '¿Qué te frenaría al compartir con anonimato garantizado?',
  '¿Leerías historias de desconocidos?',
  '¿Alguna vez algo genuino contado por alguien que no conocías te movilizó?',
  '¿Qué te motivaría a empezar a leer algo personal de alguien que no conocés?',
  'Motivar — otro (texto libre)',
  '¿Qué te haría abandonar una historia a la mitad?',
  'Abandonar — otro (texto libre)',
  '¿Qué cambia en una historia cuando no sabés quién la vivió?',
  '¿Cuánto te importa saber quién escribió algo?',
  '¿Alguna vez le contaste algo a una IA buscando que te escuche?',
  'IA: ¿Por qué nunca lo harías?',
  'IA: ¿Por qué sí lo hiciste?',
  '¿Necesitás feedback o te alcanza saber que alguien te leyó?',
  'Pregunta específica por profesión',
  '¿Qué te inspira o dice la palabra Simpleno?',
  'Cierre — todo lo que tengas para decir',
  '¿Qué tipo de experiencia te gustaría leer o compartir?'
];

// Mapeo nuevo nombre → viejo nombre para retrocompatibilidad con datos existentes
var CAMPO_ALIAS = {
  '¿Tenés cosas que no contaste?':                                    'Q3 — Te guardás historias',
  '¿Las contarías anónimas?':                                         'Q1 — Historias anónimas',
  '¿Escribir ayuda a soltarlo o procesarlo?':                         'Qsoltar — Escribir ayuda a soltar',
  '¿Lo harías para ayudar a otros aunque no te sume a vos?':          'Q6 — Ayudar a otros',
  '¿Qué te frenaría al compartir con anonimato garantizado?':         'Q10 — Frenos al compartir',
  '¿Leerías historias de desconocidos?':                              'Lector — Leerías',
  '¿Alguna vez algo genuino contado por alguien que no conocías te movilizó?': 'Lector2 — Alguna vez te movilizó',
  '¿Qué te motivaría a empezar a leer algo personal de alguien que no conocés?': 'Motivar — Opciones',
  '¿Qué te haría abandonar una historia a la mitad?':                 'Abandonar — Opciones',
  '¿Qué tipo de experiencia te gustaría leer o compartir?':           'Qtipo — Tipo de historia',
  '¿Necesitás feedback o te alcanza saber que alguien te leyó?':      'Q11 — Feedback vs resonancia',
  '¿Cuánto te importa saber quién escribió algo?':                    'Q8 — Importa saber quién',
  '¿Alguna vez le contaste algo a una IA buscando que te escuche?':   'QIA — La IA como oyente',
  'IA: ¿Por qué nunca lo harías?':                                    'IA nunca — Por qué no',
  'IA: ¿Por qué sí lo hiciste?':                                      'IA sí — Por qué sí',
  'Primera reacción al concepto':                                     'Primera reacción',
  '¿Por qué te las guardás?':                                         'Q4 — Por qué te las guardás',
  '¿Qué cambia en una historia cuando no sabés quién la vivió?':      'Q7 — Qué cambia anónimo',
  '¿Qué te inspira o dice la palabra Simpleno?':                      'El nombre — Simpleno',
  'Cierre — todo lo que tengas para decir':                           'Cierre',
  'Profesión':                                                        'Perfil'
};

function getR(r, campo) {
  return r[campo] || r[CAMPO_ALIAS[campo]] || '';
}

function doGet(e) {
  var p = e.parameter;

  if (p.action === 'dashboard') {
    return getDashboard();
  }

  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    var existingCols = sheet.getLastColumn();
    for (var i = existingCols; i < HEADERS.length; i++) {
      sheet.getRange(1, i + 1).setValue(HEADERS[i]);
    }
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
    p.qcierre       || '',
    p.qtipo         || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function repararEncabezados() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  // Escribe todos los encabezados en fila 1, columna por columna
  // No toca ninguna fila de datos
  for (var i = 0; i < HEADERS.length; i++) {
    sheet.getRange(1, i + 1).setValue(HEADERS[i]);
  }

  Logger.log('Encabezados actualizados: ' + HEADERS.length + ' columnas.');
}

function getDashboard() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
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

    var camposTexto = {
      'Primera reacción al concepto':                                '¿Cuál fue tu primera reacción al conocer el concepto de Simpleno?',
      '¿Por qué te las guardás?':                                    '¿Por qué te guardás cosas que te pasaron?',
      '¿Qué cambia en una historia cuando no sabés quién la vivió?': '¿Qué cambiaría para vos si pudieras contar algo de forma anónima?',
      '¿Qué te inspira o dice la palabra Simpleno?':                 '¿Qué te inspira o dice el nombre Simpleno?',
      'Cierre — todo lo que tengas para decir':                      '¿Qué más querías decir sobre Simpleno?'
    };

    var resumenes = {};
    Object.keys(camposTexto).forEach(function(campo) {
      var textos = data
        .map(function(r) { return getR(r, campo).toString().trim(); })
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
  } catch(e) {
    return jsonResponse({ error: e.message, total: 0, rows: [], resumenes: {} });
  }
}

function generarFODA(data) {
  var total = data.length;
  if (total === 0) return null;

  var SEM = { total: 'Totalmente', aveces: 'A veces', puedeser: 'Puede ser', nada: 'Para nada' };
  function normSem(v) { return SEM[v] || v; }

  function pctCampo(campo, valor) {
    var n = data.filter(function(r) { return normSem(getR(r, campo).trim()) === valor; }).length;
    return n + '/' + total + ' (' + Math.round(n/total*100) + '%)';
  }
  function topMulti(campo, top) {
    var c = {};
    data.forEach(function(r) {
      getR(r, campo).split(' | ').forEach(function(v) {
        v = v.trim(); if (v) c[v] = (c[v]||0) + 1;
      });
    });
    return Object.keys(c).sort(function(a,b){return c[b]-c[a];}).slice(0,top)
      .map(function(k){ return k + ' (' + c[k] + ')'; }).join(' / ');
  }

  var stats = 'Total encuestados: ' + total + '\n'
    + 'Tienen cosas sin contar: ' + data.filter(function(r){ var v = getR(r,'¿Tenés cosas que no contaste?'); return v.indexOf('Hay') !== -1 || v.indexOf('Generalmente') !== -1; }).length + '/' + total + '\n'
    + 'Las contarían anónimas (totalmente): ' + pctCampo('¿Las contarías anónimas?', 'Totalmente') + '\n'
    + 'Las contarían anónimas (a veces): ' + pctCampo('¿Las contarías anónimas?', 'A veces') + '\n'
    + 'Escribir ayuda a soltar (totalmente): ' + pctCampo('¿Escribir ayuda a soltarlo o procesarlo?', 'Totalmente') + '\n'
    + 'Compartirían si aporta a otros (totalmente): ' + pctCampo('¿Lo harías para ayudar a otros aunque no te sume a vos?', 'Totalmente — si aporta o habilita a otros, vale la pena') + '\n'
    + 'Leerían historias de desconocidos (totalmente): ' + pctCampo('¿Leerías historias de desconocidos?', 'Totalmente') + '\n'
    + 'No les importa saber quién escribió: ' + pctCampo('¿Cuánto te importa saber quién escribió algo?', 'No me importa') + '\n'
    + 'Le contaron algo a una IA: ' + data.filter(function(r){ return getR(r,'¿Alguna vez le contaste algo a una IA buscando que te escuche?').indexOf('Sí') !== -1; }).length + '/' + total + '\n'
    + 'Frenos al compartir (top opciones): ' + topMulti('¿Qué te frenaría al compartir con anonimato garantizado?', 5) + '\n'
    + 'Motivadores de lectura (top opciones): ' + topMulti('¿Qué te motivaría a empezar a leer algo personal de alguien que no conocés?', 4);

  var prompt = 'Sos estratega de producto. Simpleno es una plataforma donde las personas publican historias/experiencias de forma completamente anónima. No hay autor visible, no hay comentarios, no hay likes. Las únicas acciones son "esto resuena conmigo" y "te entiendo". El propósito es la trascendencia: que lo que viví ayude a otros (Viktor Frankl).\n\n'
    + 'Datos de encuesta de validación:\n' + stats + '\n\n'
    + 'Generá un análisis FODA del concepto basado en estos datos. Respondé SOLO con un JSON válido, sin texto extra:\n'
    + '{"fortalezas":["...","..."],"oportunidades":["...","..."],"debilidades":["...","..."],"amenazas":["...","..."]}\n'
    + 'Máximo 4 items por cuadrante. Breves y accionables, en español.';

  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=' + GEMINI_KEY;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true,
      deadline: 45
    });
    var json = JSON.parse(res.getContentText());
    var parts = json.candidates[0].content.parts;
    var textPart = parts.filter(function(p){ return !p.thought && p.text; })[0] || parts[parts.length - 1];
    var text = textPart.text.trim()
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

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=' + GEMINI_KEY;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true,
      deadline: 45
    });
    var json = JSON.parse(res.getContentText());
    var parts = json.candidates[0].content.parts;
    var textPart = parts.filter(function(p){ return !p.thought && p.text; })[0] || parts[parts.length - 1];
    return textPart.text;
  } catch(err) {
    return '(resumen no disponible)';
  }
}

function testGemini() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  Logger.log('Key: ' + (key ? key.substring(0,8) + '... (' + key.length + ' chars)' : 'NO ENCONTRADA'));
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=' + key;
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: 'Decí hola en una línea' }] }],
        generationConfig: { thinkingConfig: { thinkingBudget: 0 } }
      }),
      muteHttpExceptions: true,
      deadline: 45
    });
    Logger.log('Status: ' + res.getResponseCode());
    var json = JSON.parse(res.getContentText());
    var parts = json.candidates[0].content.parts;
    var textPart = parts.filter(function(p){ return !p.thought && p.text; })[0] || parts[parts.length-1];
    Logger.log('Respuesta: ' + textPart.text);
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
