/**
 * DASHBOARD LOOKER STUDIO - Script Automático
 *
 * INSTRUCCIONES:
 * 1. Abrí tu Google Sheets
 * 2. Menú: Extensiones > Apps Script
 * 3. Borrá todo el código que aparece
 * 4. Pegá este código completo
 * 5. Hacé click en el botón ▶ (Ejecutar)
 * 6. La primera vez te va a pedir permisos - aceptá todo
 * 7. Esperá a que termine (te aparece un mensaje de éxito)
 *
 * El script crea 4 hojas nuevas optimizadas para Looker Studio.
 * Podés ejecutarlo cada vez que actualices datos para refrescar las tablas.
 */

function crearTablasLookerStudio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('DASHBOARD');

  if (!dash) {
    SpreadsheetApp.getUi().alert('ERROR: No se encontró la hoja "DASHBOARD". Verificá que existe.');
    return;
  }

  // =============================================
  // HOJA 1: Resumen_Mensual
  // =============================================
  var sheet1 = getOrCreateSheet(ss, 'LS_Resumen_Mensual');
  sheet1.clear();

  var headers1 = ['Mes', 'Num_Mes', 'Año', 'Ingresos_USD', 'Ingresos_Pesos',
                  'Cotizacion_Dolar', 'Total_Unificado_Pesos', 'Total_Unificado_USD',
                  'Gastos_Fijos_USD', 'Total_Final_Pesos', 'Total_Final_USD'];
  sheet1.getRange(1, 1, 1, headers1.length).setValues([headers1]);
  styleHeader(sheet1, 1, headers1.length, '#2C3E50');

  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  for (var i = 0; i < 12; i++) {
    var row = i + 2;
    var srcRow = 12 + i; // Dashboard rows 12-23
    sheet1.getRange(row, 1).setValue(meses[i]);
    sheet1.getRange(row, 2).setValue(i + 1);
    sheet1.getRange(row, 3).setValue(2026);
    // Formulas that reference DASHBOARD directly
    sheet1.getRange(row, 4).setFormula("=DASHBOARD!E" + srcRow);   // USD
    sheet1.getRange(row, 5).setFormula("=DASHBOARD!F" + srcRow);   // Pesos
    sheet1.getRange(row, 6).setFormula("=DASHBOARD!G" + srcRow);   // Cotización
    sheet1.getRange(row, 7).setFormula("=DASHBOARD!H" + srcRow);   // Total Unificado Pesos
    sheet1.getRange(row, 8).setFormula("=DASHBOARD!I" + srcRow);   // Total Unificado USD
    sheet1.getRange(row, 9).setFormula("=DASHBOARD!J" + srcRow);   // Gastos
    sheet1.getRange(row, 10).setFormula("=DASHBOARD!K" + srcRow);  // Total Final Pesos
    sheet1.getRange(row, 11).setFormula("=DASHBOARD!L" + srcRow);  // Total Final USD
  }

  autoResize(sheet1, headers1.length);

  // =============================================
  // HOJA 2: Ganancias_Vendedor (flat table)
  // =============================================
  var sheet2 = getOrCreateSheet(ss, 'LS_Ganancias_Vendedor');
  sheet2.clear();

  var headers2 = ['Mes', 'Num_Mes', 'Año', 'Vendedor', 'Ganancia_USD', 'Ganancia_Pesos'];
  sheet2.getRange(1, 1, 1, headers2.length).setValues([headers2]);
  styleHeader(sheet2, 1, headers2.length, '#1ABC9C');

  var vendedores = ['Tecno','Inversiones','Creditos','Osito S.r.l.','Patito S.A.',
                    'Gonza','Mely','Linea 314','Tobias','Tino'];
  // USD: rows 35-46, cols E-N (5-14)
  // Pesos: rows 57-68, cols E-N (5-14)

  var row2 = 2;
  for (var m = 0; m < 12; m++) {
    for (var v = 0; v < vendedores.length; v++) {
      var usdRow = 35 + m;
      var pesosRow = 57 + m;
      var col = 5 + v; // E=5, F=6, ... N=14
      var colLetter = columnToLetter(col);

      sheet2.getRange(row2, 1).setValue(meses[m]);
      sheet2.getRange(row2, 2).setValue(m + 1);
      sheet2.getRange(row2, 3).setValue(2026);
      sheet2.getRange(row2, 4).setValue(vendedores[v]);
      sheet2.getRange(row2, 5).setFormula("=DASHBOARD!" + colLetter + usdRow);
      sheet2.getRange(row2, 6).setFormula("=DASHBOARD!" + colLetter + pesosRow);
      row2++;
    }
  }

  autoResize(sheet2, headers2.length);

  // =============================================
  // HOJA 3: Modulos_Negocio (detail per module)
  // =============================================
  var sheet3 = getOrCreateSheet(ss, 'LS_Modulos_Negocio');
  sheet3.clear();

  var headers3 = ['Mes', 'Num_Mes', 'Año', 'Modulo', 'Ganancia_USD', 'Ganancia_Pesos',
                  'Inversion_USD', 'Resta_Cobrar_USD', 'Caja_Actual_USD',
                  'Inversion_Pesos', 'Resta_Cobrar_Pesos'];
  sheet3.getRange(1, 1, 1, headers3.length).setValues([headers3]);
  styleHeader(sheet3, 1, headers3.length, '#E74C3C');

  // Module layout on DASHBOARD:
  // Row 81: Osito(C), Patito(H), Gonza(M) - data rows 83-94, totals 96, invest 98, resta 100, caja 101
  // Row 105: Mely(C), VendedorX(H), Tobias(M) - data rows 107-118, totals 120, invest 122, resta 124, caja 125
  // Row 129: Tecno(C), Inversiones(H), Creditos(M) - data rows 131-142, totals 144, invest 146, resta 148, caja 149
  // Row 153: Tino(C), Linea314(H) - data rows 155-166, totals 168, invest 170, resta 172, caja 173

  var moduleConfig = [
    // [name, usdCol, pesosCol, dataStartRow, investRow, restaRow, cajaRow, investPesosCol, restaPesosCol]
    {name:'Osito S.r.l.', usdCol:'E', pesosCol:'F', startRow:83, investRow:98, investCol:'E', restaRow:100, restaCol:'E', cajaRow:101, cajaCol:'E', investPesosCol:'F', restaPesosCol:'F'},
    {name:'Patito S.A.', usdCol:'J', pesosCol:'K', startRow:83, investRow:98, investCol:'J', restaRow:100, restaCol:'J', cajaRow:101, cajaCol:'J', investPesosCol:'K', restaPesosCol:'K'},
    {name:'Gonza', usdCol:'O', pesosCol:'P', startRow:83, investRow:98, investCol:'O', restaRow:100, restaCol:'O', cajaRow:101, cajaCol:'O', investPesosCol:'P', restaPesosCol:'P'},
    {name:'Mely', usdCol:'E', pesosCol:'F', startRow:107, investRow:122, investCol:'E', restaRow:124, restaCol:'E', cajaRow:125, cajaCol:'E', investPesosCol:'F', restaPesosCol:'F'},
    {name:'Tobias', usdCol:'O', pesosCol:'P', startRow:107, investRow:122, investCol:'O', restaRow:124, restaCol:'O', cajaRow:125, cajaCol:'O', investPesosCol:'P', restaPesosCol:'P'},
    {name:'Tecno', usdCol:'E', pesosCol:'F', startRow:131, investRow:146, investCol:'E', restaRow:148, restaCol:'E', cajaRow:149, cajaCol:'E', investPesosCol:'F', restaPesosCol:'F'},
    {name:'Inversiones', usdCol:'J', pesosCol:'K', startRow:131, investRow:146, investCol:'J', restaRow:148, restaCol:'J', cajaRow:149, cajaCol:'J', investPesosCol:'K', restaPesosCol:'K'},
    {name:'Creditos', usdCol:'O', pesosCol:'P', startRow:131, investRow:146, investCol:'O', restaRow:148, restaCol:'O', cajaRow:149, cajaCol:'O', investPesosCol:'P', restaPesosCol:'P'},
    {name:'Tino', usdCol:'E', pesosCol:'F', startRow:155, investRow:170, investCol:'E', restaRow:172, restaCol:'E', cajaRow:173, cajaCol:'E', investPesosCol:'F', restaPesosCol:'F'},
    {name:'Linea 314', usdCol:'J', pesosCol:'K', startRow:155, investRow:170, investCol:'J', restaRow:172, restaCol:'J', cajaRow:173, cajaCol:'J', investPesosCol:'K', restaPesosCol:'K'},
  ];

  var row3 = 2;
  for (var mc = 0; mc < moduleConfig.length; mc++) {
    var mod = moduleConfig[mc];
    for (var m = 0; m < 12; m++) {
      var dataRow = mod.startRow + m;
      sheet3.getRange(row3, 1).setValue(meses[m]);
      sheet3.getRange(row3, 2).setValue(m + 1);
      sheet3.getRange(row3, 3).setValue(2026);
      sheet3.getRange(row3, 4).setValue(mod.name);
      sheet3.getRange(row3, 5).setFormula("=DASHBOARD!" + mod.usdCol + dataRow);
      sheet3.getRange(row3, 6).setFormula("=DASHBOARD!" + mod.pesosCol + dataRow);
      sheet3.getRange(row3, 7).setFormula("=DASHBOARD!" + mod.investCol + mod.investRow);
      sheet3.getRange(row3, 8).setFormula("=DASHBOARD!" + mod.restaCol + mod.restaRow);
      sheet3.getRange(row3, 9).setFormula("=DASHBOARD!" + mod.cajaCol + mod.cajaRow);
      sheet3.getRange(row3, 10).setFormula("=DASHBOARD!" + mod.investPesosCol + mod.investRow);
      sheet3.getRange(row3, 11).setFormula("=DASHBOARD!" + mod.restaPesosCol + mod.restaRow);
      row3++;
    }
  }

  autoResize(sheet3, headers3.length);

  // =============================================
  // HOJA 4: Panel_Control
  // =============================================
  var sheet4 = getOrCreateSheet(ss, 'LS_Panel_Control');
  sheet4.clear();

  var headers4 = ['Mes', 'Num_Mes', 'Año', 'Totales_USD', 'Totales_Pesos',
                  'Gastos_USD', 'Balance_USD'];
  sheet4.getRange(1, 1, 1, headers4.length).setValues([headers4]);
  styleHeader(sheet4, 1, headers4.length, '#3498DB');

  // PANEL DE CONTROL sheet: Row 6=Totales USD, Row 7=Totales $, Row 8=Gastos, Row 9=Balance
  // Columns: C=Enero, E=Febrero, G=Marzo, I=Abril, K=Mayo, M=Junio, O=Julio, Q=Agosto, S=Sept, U=Oct, W=Nov, Y=Dic
  var panelCols = ['C','E','G','I','K','M','O','Q','S','U','W','Y'];

  var panelSheet = ss.getSheetByName('PANEL DE CONTROL');
  if (panelSheet) {
    for (var m = 0; m < 12; m++) {
      var r4 = m + 2;
      var pc = panelCols[m];
      sheet4.getRange(r4, 1).setValue(meses[m]);
      sheet4.getRange(r4, 2).setValue(m + 1);
      sheet4.getRange(r4, 3).setValue(2026);
      sheet4.getRange(r4, 4).setFormula("='PANEL DE CONTROL'!" + pc + "6");
      sheet4.getRange(r4, 5).setFormula("='PANEL DE CONTROL'!" + pc + "7");
      sheet4.getRange(r4, 6).setFormula("='PANEL DE CONTROL'!" + pc + "8");
      sheet4.getRange(r4, 7).setFormula("='PANEL DE CONTROL'!" + pc + "9");
    }
  }

  autoResize(sheet4, headers4.length);

  // =============================================
  // Success message
  // =============================================
  SpreadsheetApp.getUi().alert(
    '✅ ¡LISTO!\n\n' +
    'Se crearon 4 hojas nuevas:\n' +
    '• LS_Resumen_Mensual\n' +
    '• LS_Ganancias_Vendedor\n' +
    '• LS_Modulos_Negocio\n' +
    '• LS_Panel_Control\n\n' +
    'Estas hojas se actualizan automáticamente cuando modificás datos en DASHBOARD.\n\n' +
    'Ahora podés conectar estas hojas a Looker Studio.'
  );
}

// =============================================
// Helper functions
// =============================================

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    return sheet;
  }
  return ss.insertSheet(name);
}

function styleHeader(sheet, row, numCols, color) {
  var range = sheet.getRange(row, 1, 1, numCols);
  range.setBackground(color);
  range.setFontColor('#FFFFFF');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  range.setBorder(true, true, true, true, true, true);
}

function autoResize(sheet, numCols) {
  for (var i = 1; i <= numCols; i++) {
    sheet.setColumnWidth(i, 180);
  }
  // Freeze header row
  sheet.setFrozenRows(1);
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
