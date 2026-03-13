// ============================================================
// CONFIGURACION
// ============================================================
var VENDEDORES = [
  {id: '1kIO9TlRatBTWP5K1sM5KeguFXNB5qzSu80W41mQkbZw', nombre: 'LINEA 314', filas: [8,10,12,14,16,18,20,22,24,26,28,30], cols: ['D','E']},
  {id: '1hrDYiUGbfwars04Wx_ZImVrrgLZKIzO6bDk-CxGeX-c', nombre: 'OSITO S.R.L', filas: [6,8,10,12,14,16,18,20,22,24,26,28], cols: ['D','E','F','G']},
  {id: '1_EJbkqX7Xp8ui8QCaNaIMv3SteKTApx9S7xWiCbn2Q4', nombre: 'MELY', filas: [8,10,12,14,16,18,20,22,24,26,28,30], cols: ['D','E','F','H','I','J']},
  {id: '1DlKcy7lmn0Yr02fGrEUf8FiB-BEltv7eVw5nKuObTag', nombre: 'GONZA', filas: [6,8,10,12,14,16,18,20,22,24,26,28], cols: ['D','E','F']},
  {id: '1_jCQkl2fBgsWVH326o6VBLq2tciSK6TZw43NCb0UT8Q', nombre: 'TOBIAS', filas: [6,8,10,12,14,16,18,20,22,24,26,28], cols: ['D','E','F']},
  {id: '1KBusYiaUuD4-rQ-JHTTv6kaH27xHyC4p6IFyRoScimM', nombre: 'TINO', filas: [6,8,10,12,14,16,18,20,22,24,26,28], cols: ['D','E','F','G']},
  {id: '1k1Uyphm-df7eN6IyEx77fqOixuROy4p1Cfq8t5tng78', nombre: 'PATO', filas: [6,8,10,12,14,16,18,20,22,24,26,28], cols: ['D','E']},
  {id: '1t42edtRuqh3mJRV7hKnBKaSq_THwd9FwbpqEc8Zn86U', nombre: 'MODULO 1', filas: [8,10,12,14,16,18,20,22,24,26,28,30], cols: ['D','E','F','H','I','J']},
  {id: '1zpIuSEtlmm4fZN-YIy86u_NLp3actVdFj9zy78pCSbA', nombre: 'MODULO 2', filas: [8,10,12,14,16,18,20,22,24,26,28,30], cols: ['D','E','F','H','I','J']}
];
var MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

// ============================================================
// BOTON PRINCIPAL - CONGELAR MES
// ============================================================
function congelarMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = ss.getSheetByName('PANEL DE CONTROL');
  var dash = ss.getSheetByName('DASHBOARD');
  var ui = SpreadsheetApp.getUi();

  // Config planilla principal
  var filasPesos = [17,22,27,32,37,42,47,52,57,62,67,72];
  var filasUSD = [18,23,28,33,38,43,48,53,58,63,68,73];
  var filasCotiz = [12,13,14,15,16,17,18,19,20,21,22,23];

  // Buscar mes activo (el que tiene formulas)
  var mesCongelar = -1;
  for (var m = 0; m < 12; m++) {
    if (panel.getRange('E' + filasPesos[m]).getFormula() ||
        panel.getRange('E' + filasUSD[m]).getFormula() ||
        panel.getRange('I' + filasPesos[m]).getFormula() ||
        panel.getRange('I' + filasUSD[m]).getFormula()) {
      mesCongelar = m;
      break;
    }
  }

  if (mesCongelar < 0) {
    // Buscar en vendedores
    for (var vi = 0; vi < VENDEDORES.length && mesCongelar < 0; vi++) {
      try {
        var extSS = SpreadsheetApp.openById(VENDEDORES[vi].id);
        var vP = extSS.getSheetByName('PANEL DE CONTROL');
        if (!vP) continue;
        for (var m = 0; m < 12; m++) {
          if (vP.getRange(VENDEDORES[vi].cols[0] + VENDEDORES[vi].filas[m]).getFormula()) {
            mesCongelar = m;
            break;
          }
        }
      } catch(e) {}
    }
  }

  if (mesCongelar < 0) {
    ui.alert('No se encontro un mes con formulas para congelar.');
    return;
  }

  if (mesCongelar > 10) {
    ui.alert('Diciembre es el ultimo mes del año.');
    return;
  }

  var mesSig = mesCongelar + 1;

  var resp = ui.alert(
    'CONGELAR ' + MESES[mesCongelar],
    'Se va a:\n' +
    '1. Congelar ' + MESES[mesCongelar] + ' (valores fijos)\n' +
    '2. Crear formulas para ' + MESES[mesSig] + '\n\n' +
    'En: Tu planilla + 7 vendedores\n\n' +
    'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = [];

  // === PLANILLA PRINCIPAL - PANEL DE CONTROL ===
  congelarCelda(panel, 'E' + filasPesos[mesCongelar], 'E' + filasPesos[mesSig], log, 'TECNO Pesos');
  congelarCelda(panel, 'E' + filasUSD[mesCongelar], 'E' + filasUSD[mesSig], log, 'TECNO USD');
  congelarCelda(panel, 'I' + filasPesos[mesCongelar], 'I' + filasPesos[mesSig], log, 'CREDITOS Pesos');
  congelarCelda(panel, 'I' + filasUSD[mesCongelar], 'I' + filasUSD[mesSig], log, 'CREDITOS USD');

  // === PLANILLA PRINCIPAL - DASHBOARD COTIZACION ===
  var cotizCell = 'G' + filasCotiz[mesCongelar];
  var cotizFormula = dash.getRange(cotizCell).getFormula();
  var cotizValor = dash.getRange(cotizCell).getValue();
  if (cotizFormula) {
    dash.getRange(cotizCell).setValue(cotizValor);
    dash.getRange('G' + filasCotiz[mesSig]).setFormula(cotizFormula);
    log.push('COTIZACION: congelado $' + cotizValor);
  } else {
    log.push('COTIZACION: sin formula');
  }

  // === VENDEDORES ===
  VENDEDORES.forEach(function(v) {
    try {
      var extSS = SpreadsheetApp.openById(v.id);
      var vPanel = extSS.getSheetByName('PANEL DE CONTROL');
      if (!vPanel) { log.push(v.nombre + ': ERROR - No PANEL DE CONTROL'); return; }
      v.cols.forEach(function(col) {
        congelarCelda(vPanel, col + v.filas[mesCongelar], col + v.filas[mesSig], log, v.nombre + ' ' + col);
      });
    } catch(e) { log.push(v.nombre + ': ERROR - ' + e.message); }
  });

  ui.alert(
    'LISTO - ' + MESES[mesCongelar] + ' CONGELADO',
    MESES[mesCongelar] + ' congelado\n' +
    MESES[mesSig] + ' configurado\n\n' +
    log.join('\n'),
    ui.ButtonSet.OK
  );
}

// ============================================================
// SIMULACION - muestra que pasaria sin tocar nada
// ============================================================
function simularCongelamiento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = ss.getSheetByName('PANEL DE CONTROL');
  var dash = ss.getSheetByName('DASHBOARD');
  var ui = SpreadsheetApp.getUi();

  var filasPesos = [17,22,27,32,37,42,47,52,57,62,67,72];
  var filasUSD = [18,23,28,33,38,43,48,53,58,63,68,73];
  var filasCotiz = [12,13,14,15,16,17,18,19,20,21,22,23];

  // Buscar mes activo
  var mesCongelar = -1;
  for (var m = 0; m < 12; m++) {
    if (panel.getRange('E' + filasPesos[m]).getFormula() ||
        panel.getRange('E' + filasUSD[m]).getFormula() ||
        panel.getRange('I' + filasPesos[m]).getFormula() ||
        panel.getRange('I' + filasUSD[m]).getFormula()) {
      mesCongelar = m;
      break;
    }
  }
  if (mesCongelar < 0) {
    for (var vi = 0; vi < VENDEDORES.length && mesCongelar < 0; vi++) {
      try {
        var extSS = SpreadsheetApp.openById(VENDEDORES[vi].id);
        var vP = extSS.getSheetByName('PANEL DE CONTROL');
        if (!vP) continue;
        for (var m = 0; m < 12; m++) {
          if (vP.getRange(VENDEDORES[vi].cols[0] + VENDEDORES[vi].filas[m]).getFormula()) {
            mesCongelar = m;
            break;
          }
        }
      } catch(e) {}
    }
  }
  if (mesCongelar < 0) { ui.alert('No se encontro un mes con formulas para simular.'); return; }
  if (mesCongelar > 10) { ui.alert('Diciembre es el ultimo mes.'); return; }
  var mesSig = mesCongelar + 1;

  // Crear o limpiar hoja SIMULACION
  var sim = ss.getSheetByName('SIMULACION');
  if (sim) ss.deleteSheet(sim);
  sim = ss.insertSheet('SIMULACION');

  // Titulo
  sim.getRange('A1').setValue('SIMULACION DE CONGELAMIENTO - ' + MESES[mesCongelar]).setFontSize(14).setFontWeight('bold');
  sim.getRange('A2').setValue('Esta hoja es solo una SIMULACION. No se modifico ninguna celda original.').setFontColor('red').setFontWeight('bold');

  // Encabezados
  var headers = ['PLANILLA','COLUMNA','CELDA ACTUAL','FORMULA ACTUAL','VALOR ACTUAL','SE CONGELA A','CELDA SIGUIENTE','FORMULA SIGUIENTE'];
  for (var h = 0; h < headers.length; h++) {
    sim.getRange(4, h + 1).setValue(headers[h]);
  }
  sim.getRange('A4:H4').setFontWeight('bold').setBackground('#4a86c8').setFontColor('white');

  var row = 5;

  // --- PLANILLA PRINCIPAL ---
  var mainCells = [
    {col: 'E', fila: filasPesos[mesCongelar], label: 'TECNO Pesos', filaSig: filasPesos[mesSig]},
    {col: 'E', fila: filasUSD[mesCongelar], label: 'TECNO USD', filaSig: filasUSD[mesSig]},
    {col: 'I', fila: filasPesos[mesCongelar], label: 'CREDITOS Pesos', filaSig: filasPesos[mesSig]},
    {col: 'I', fila: filasUSD[mesCongelar], label: 'CREDITOS USD', filaSig: filasUSD[mesSig]}
  ];
  mainCells.forEach(function(c) {
    var cellRef = c.col + c.fila;
    var formula = panel.getRange(cellRef).getFormula();
    var valor = panel.getRange(cellRef).getValue();
    var nextFormula = '';
    var nextCell = c.col + c.filaSig;
    if (formula) {
      var match = formula.match(/^='?([^'!]+)'?!([A-Z]+\d+)-(.+)$/);
      if (match) { nextFormula = "='" + match[1] + "'!" + match[2] + '-' + (parseFloat(match[3]) + valor); }
    }
    sim.getRange(row, 1).setValue('PRINCIPAL');
    sim.getRange(row, 2).setValue(c.label);
    sim.getRange(row, 3).setValue(cellRef);
    sim.getRange(row, 4).setValue(formula || '(sin formula)');
    sim.getRange(row, 5).setValue(valor);
    sim.getRange(row, 6).setValue(formula ? valor : '(manual - no cambia)');
    sim.getRange(row, 7).setValue(nextCell);
    sim.getRange(row, 8).setValue(nextFormula || '(no aplica)');
    row++;
  });

  // Cotizacion
  var cotizCell = 'G' + filasCotiz[mesCongelar];
  var cotizFormula = dash.getRange(cotizCell).getFormula();
  var cotizValor = dash.getRange(cotizCell).getValue();
  sim.getRange(row, 1).setValue('DASHBOARD');
  sim.getRange(row, 2).setValue('COTIZACION');
  sim.getRange(row, 3).setValue(cotizCell);
  sim.getRange(row, 4).setValue(cotizFormula || '(sin formula)');
  sim.getRange(row, 5).setValue(cotizValor);
  sim.getRange(row, 6).setValue(cotizFormula ? cotizValor : '(manual - no cambia)');
  sim.getRange(row, 7).setValue('G' + filasCotiz[mesSig]);
  sim.getRange(row, 8).setValue(cotizFormula || '(no aplica)');
  row++;

  // Separador
  row++;
  sim.getRange(row, 1).setValue('--- VENDEDORES ---').setFontWeight('bold').setFontSize(12);
  row++;

  // --- VENDEDORES ---
  VENDEDORES.forEach(function(v) {
    try {
      var extSS = SpreadsheetApp.openById(v.id);
      var vPanel = extSS.getSheetByName('PANEL DE CONTROL');
      if (!vPanel) {
        sim.getRange(row, 1).setValue(v.nombre);
        sim.getRange(row, 2).setValue('ERROR: No tiene PANEL DE CONTROL');
        sim.getRange(row, 1, 1, 8).setBackground('#ffcccc');
        row++; return;
      }
      v.cols.forEach(function(col) {
        var cellRef = col + v.filas[mesCongelar];
        var cellSig = col + v.filas[mesSig];
        var formula = vPanel.getRange(cellRef).getFormula();
        var valor = vPanel.getRange(cellRef).getValue();
        var nextFormula = '';
        if (formula) {
          var match = formula.match(/^='?([^'!]+)'?!([A-Z]+\d+)-(.+)$/);
          if (match) { nextFormula = "='" + match[1] + "'!" + match[2] + '-' + (parseFloat(match[3]) + valor); }
        }
        sim.getRange(row, 1).setValue(v.nombre);
        sim.getRange(row, 2).setValue('Col ' + col);
        sim.getRange(row, 3).setValue(cellRef);
        sim.getRange(row, 4).setValue(formula || '(sin formula)');
        sim.getRange(row, 5).setValue(valor);
        sim.getRange(row, 6).setValue(formula ? valor : '(manual - no cambia)');
        sim.getRange(row, 7).setValue(cellSig);
        sim.getRange(row, 8).setValue(nextFormula || '(no aplica)');
        if (!formula && valor) sim.getRange(row, 1, 1, 8).setBackground('#fff3cd');
        row++;
      });
    } catch(e) {
      sim.getRange(row, 1).setValue(v.nombre);
      sim.getRange(row, 2).setValue('ERROR: ' + e.message);
      sim.getRange(row, 1, 1, 8).setBackground('#ffcccc');
      row++;
    }
  });

  // Resumen
  row++;
  sim.getRange(row, 1).setValue('LEYENDA:').setFontWeight('bold');
  row++;
  sim.getRange(row, 1).setValue('Fondo blanco = tiene formula, se va a congelar');
  row++;
  sim.getRange(row, 1).setValue('Fondo amarillo = valor manual, no necesita congelamiento (ya es fijo)');
  sim.getRange(row, 1, 1, 8).setBackground('#fff3cd');
  row++;
  sim.getRange(row, 1).setValue('Fondo rojo = error al leer la planilla');
  sim.getRange(row, 1, 1, 8).setBackground('#ffcccc');

  // Autoajustar
  for (var c = 1; c <= 8; c++) { sim.autoResizeColumn(c); }

  ui.alert(
    'SIMULACION COMPLETA',
    'Se creo la hoja "SIMULACION" en tu planilla principal.\n\n' +
    'Muestra que pasaria al congelar ' + MESES[mesCongelar] + '.\n\n' +
    'NO se modifico ninguna celda original de ninguna planilla.\n\n' +
    'Podes borrar la hoja SIMULACION cuando quieras.',
    ui.ButtonSet.OK
  );
}

// ============================================================
// VERIFICAR MES - Parte 1 (vendedores 1-5)
// ============================================================
function verificarMes() {
  var ui = SpreadsheetApp.getUi();
  var mesCongelar = detectarMesActivo_();
  if (mesCongelar < 0) { ui.alert('No se encontro un mes con formulas.'); return; }
  if (mesCongelar > 10) { ui.alert('Diciembre es el ultimo mes.'); return; }

  var resp = ui.alert(
    'VERIFICAR ' + MESES[mesCongelar] + ' (Parte 1/2)',
    'Escribe verificacion en: LINEA 314, OSITO, MELY, GONZA, TOBIAS\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = procesarVendedores_(0, 5, mesCongelar);
  ui.alert('PARTE 1 COMPLETA', log.join('\n') + '\n\nAhora ejecuta verificarMes2 para el resto.', ui.ButtonSet.OK);
}

// ============================================================
// VERIFICAR MES - Parte 2 (vendedores 6-9 + principal)
// ============================================================
function verificarMes2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = ss.getSheetByName('PANEL DE CONTROL');
  var dash = ss.getSheetByName('DASHBOARD');
  var ui = SpreadsheetApp.getUi();
  var mesCongelar = detectarMesActivo_();
  if (mesCongelar < 0) { ui.alert('No se encontro un mes con formulas.'); return; }

  var filasPesos = [17,22,27,32,37,42,47,52,57,62,67,72];
  var filasUSD = [18,23,28,33,38,43,48,53,58,63,68,73];
  var filasCotiz = [12,13,14,15,16,17,18,19,20,21,22,23];

  var resp = ui.alert(
    'VERIFICAR ' + MESES[mesCongelar] + ' (Parte 2/2)',
    'Escribe verificacion en: TINO, PATO, MODULO 1, MODULO 2 + principal\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = procesarVendedores_(5, VENDEDORES.length, mesCongelar);

  var ep = panel.getRange('E' + filasPesos[mesCongelar]).getValue();
  var eu = panel.getRange('E' + filasUSD[mesCongelar]).getValue();
  var ip = panel.getRange('I' + filasPesos[mesCongelar]).getValue();
  var iu = panel.getRange('I' + filasUSD[mesCongelar]).getValue();
  var cot = dash.getRange('G' + filasCotiz[mesCongelar]).getValue();

  ui.alert(
    'VERIFICACION ' + MESES[mesCongelar] + ' COMPLETA',
    'PLANILLA PRINCIPAL:\n' +
    '  TECNO $: ' + ep + '  |  TECNO USD: ' + eu + '\n' +
    '  CRED $: ' + ip + '  |  CRED USD: ' + iu + '\n' +
    '  COTIZ: ' + cot + '\n\n' +
    'VENDEDORES:\n' + log.join('\n') + '\n\n' +
    'Ahora congela manualmente y compara con columnas V.',
    ui.ButtonSet.OK
  );
}

// ============================================================
// FUNCIONES INTERNAS DE VERIFICACION
// ============================================================
function detectarMesActivo_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = ss.getSheetByName('PANEL DE CONTROL');
  var filasPesos = [17,22,27,32,37,42,47,52,57,62,67,72];
  var filasUSD = [18,23,28,33,38,43,48,53,58,63,68,73];
  for (var m = 0; m < 12; m++) {
    if (panel.getRange('E' + filasPesos[m]).getFormula() ||
        panel.getRange('E' + filasUSD[m]).getFormula() ||
        panel.getRange('I' + filasPesos[m]).getFormula() ||
        panel.getRange('I' + filasUSD[m]).getFormula()) return m;
  }
  for (var vi = 0; vi < VENDEDORES.length; vi++) {
    try {
      var vP = SpreadsheetApp.openById(VENDEDORES[vi].id).getSheetByName('PANEL DE CONTROL');
      if (!vP) continue;
      for (var m = 0; m < 12; m++) {
        if (vP.getRange(VENDEDORES[vi].cols[0] + VENDEDORES[vi].filas[m]).getFormula()) return m;
      }
    } catch(e) {}
  }
  return -1;
}

function procesarVendedores_(desde, hasta, mesCongelar) {
  var VERIF_START = {
    'LINEA 314': 7, 'OSITO S.R.L': 14, 'MELY': 18, 'GONZA': 13,
    'TOBIAS': 13, 'TINO': 13, 'PATO': 12, 'MODULO 1': 18, 'MODULO 2': 18
  };
  var mesNombre = MESES[mesCongelar];
  var log = [];

  for (var i = desde; i < hasta && i < VENDEDORES.length; i++) {
    var v = VENDEDORES[i];
    try {
      var extSS = SpreadsheetApp.openById(v.id);
      var vPanel = extSS.getSheetByName('PANEL DE CONTROL');
      if (!vPanel) { log.push(v.nombre + ': ERROR'); continue; }

      var startCol = VERIF_START[v.nombre] || 20;
      var numCols = v.cols.length;

      // Leer todos los valores en batch
      var colNums = [];
      for (var ci = 0; ci < numCols; ci++) { colNums.push(v.cols[ci].charCodeAt(0) - 64); }
      var minC = Math.min.apply(null, colNums);
      var maxC = Math.max.apply(null, colNums);
      var rowData = vPanel.getRange(v.filas[mesCongelar], minC, 1, maxC - minC + 1).getValues()[0];
      var verifValues = [];
      for (var ci = 0; ci < numCols; ci++) { verifValues.push(rowData[colNums[ci] - minC]); }

      // Escribir titulo en una celda (sin merge para velocidad)
      vPanel.getRange(Math.max(1, v.filas[0] - 2), startCol)
        .setValue('VERIF ' + mesNombre)
        .setFontWeight('bold')
        .setBackground('#c8e6c9');

      // Escribir encabezados en batch
      var headers = [];
      for (var ci = 0; ci < numCols; ci++) { headers.push('V.' + v.cols[ci]); }
      vPanel.getRange(Math.max(2, v.filas[0] - 1), startCol, 1, numCols)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#e8f5e9');

      // Escribir valores en batch
      vPanel.getRange(v.filas[mesCongelar], startCol, 1, numCols)
        .setValues([verifValues])
        .setBackground('#f1f8e9');

      SpreadsheetApp.flush();
      log.push(v.nombre + ': OK');
    } catch(e) {
      log.push(v.nombre + ': ERROR - ' + e.message);
    }
  }
  return log;
}

// ============================================================
// FUNCION AUXILIAR - congela una celda y crea la formula siguiente
// ============================================================
function congelarCelda(sheet, cellActual, cellSiguiente, log, label) {
  var formula = sheet.getRange(cellActual).getFormula();
  var valor = sheet.getRange(cellActual).getValue();

  if (!formula) {
    if (valor !== '' && valor !== 0 && valor !== null) {
      log.push(label + ': ya congelado (' + valor + ')');
    }
    return;
  }

  var match = formula.match(/^='?([^'!]+)'?!([A-Z]+\d+)-(.+)$/);
  if (match) {
    var source = match[1];
    var cell = match[2];
    var offset = parseFloat(match[3]);

    // Congelar
    sheet.getRange(cellActual).setValue(valor);

    // Formula siguiente
    var newOffset = offset + valor;
    sheet.getRange(cellSiguiente).setFormula("='" + source + "'!" + cell + '-' + newOffset);

    log.push(label + ': ' + valor + ' -> prox offset ' + newOffset);
  } else {
    log.push(label + ': formula no estandar');
  }
}
