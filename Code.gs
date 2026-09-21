const CATEGORIAS_DEFAULT = [
  'Mercado', 'Transporte', 'Vivienda/Servicios', 'Salud',
  'Entretenimiento', 'Ropa', 'Deudas/Préstamos', 'Otros'
];

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page === 'resumen') ? 'Resumen' : 'Formulario';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const gastos = ss.getSheetByName('Gastos') || ss.insertSheet('Gastos');
  if (gastos.getLastRow() === 0) {
    gastos.appendRow(['Fecha', 'Categoría', 'Descripción', 'Monto']);
  }

  const categorias = ss.getSheetByName('Categorías') || ss.insertSheet('Categorías');
  if (categorias.getLastRow() === 0) {
    categorias.appendRow(['Categoría']);
    CATEGORIAS_DEFAULT.forEach(c => categorias.appendRow([c]));
  }

  const presupuesto = ss.getSheetByName('Presupuesto') || ss.insertSheet('Presupuesto');
  if (presupuesto.getLastRow() === 0) {
    presupuesto.appendRow(['Categoría', 'Límite Mensual']);
    CATEGORIAS_DEFAULT.forEach(c => presupuesto.appendRow([c, 0]));
  }

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name !== 'Gastos' && name !== 'Categorías' && name !== 'Presupuesto') {
      ss.deleteSheet(sheet);
    }
  });
}

function getCategorias() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categorías');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(String);
}

function addExpense(fecha, categoria, descripcion, monto) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gastos');
  sheet.appendRow([new Date(fecha), categoria, descripcion, parseFloat(monto)]);
  return { ok: true };
}

function getPresupuesto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Presupuesto');
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  sheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(row => {
    if (row[0]) map[row[0]] = Number(row[1]) || 0;
  });
  return map;
}

function getResumen(anio, mes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gastos');
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

  const now = new Date();
  const y = anio || now.getFullYear();
  const m = mes || (now.getMonth() + 1);

  const totales = {};
  let total = 0;

  data.forEach(row => {
    const fecha = row[0];
    if (!(fecha instanceof Date)) return;
    if (fecha.getFullYear() === y && fecha.getMonth() + 1 === m) {
      const cat = row[1];
      const monto = Number(row[3]) || 0;
      totales[cat] = (totales[cat] || 0) + monto;
      total += monto;
    }
  });

  const presupuesto = getPresupuesto();
  const categorias = getCategorias();
  const detalle = categorias.map(cat => ({
    categoria: cat,
    gastado: totales[cat] || 0,
    presupuesto: presupuesto[cat] || 0
  }));

  return { anio: y, mes: m, total, detalle };
}
