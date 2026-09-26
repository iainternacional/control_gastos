/**
 * Este script es standalone (no está vinculado a una hoja), así que
 * SpreadsheetApp.getActiveSpreadsheet() devuelve null. Se resuelve la hoja por ID.
 *
 * Para obtener el ID: abre tu Google Sheet y copia el segmento entre
 * /spreadsheets/d/ y /edit en la URL. Ejemplo:
 * https://docs.google.com/spreadsheets/d/1AbC...XyZ/edit  ->  1AbC...XyZ
 */
const SPREADSHEET_ID = '10O8anjtuoIF4IV1ZjVW8Bo8yx5ro4mo1SRW58YP15E4';

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * URL fija del despliegue activo. ScriptApp.getService().getUrl() no es
 * confiable en despliegues restringidos al dominio de Workspace (a veces
 * devuelve la URL sin el prefijo /a/macros/<dominio>/, que falla al abrirse
 * con "No se puede abrir el archivo en estos momentos"). Si el usuario crea
 * un nuevo despliegue con otra URL, este valor debe actualizarse a mano.
 */
const APP_URL = 'https://script.google.com/a/macros/ipuc.org.co/s/AKfycbw3Ha8N1l4dm7HPmv-me279741xCleAIzw2ct0PETqRU57ZZ1W2Xrmn8ae6PbBnilfpfg/exec';

const CATEGORIAS_DEFAULT = [
  'Mercado', 'Transporte', 'Vivienda/Servicios', 'Salud',
  'Entretenimiento', 'Ropa', 'Deudas/Préstamos', 'Otros'
];

const CATEGORIAS_INGRESOS_DEFAULT = [
  'Salario', 'Dinero regalado', 'Ingreso extra', 'Intereses/Rendimientos'
];

function doGet(e) {
  const param = e && e.parameter && e.parameter.page;
  const page = param === 'resumen' ? 'Resumen' : param === 'ingreso' ? 'Ingreso' : 'Formulario';
  const authuser = e && e.parameter && e.parameter.authuser;
  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = APP_URL;
  template.authSuffix = authuser ? ('&authuser=' + authuser) : '';
  return template.evaluate()
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheets() {
  const ss = getSheet();

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

  const ingresos = ss.getSheetByName('Ingresos') || ss.insertSheet('Ingresos');
  if (ingresos.getLastRow() === 0) {
    ingresos.appendRow(['Fecha', 'Categoría', 'Descripción', 'Monto']);
  }

  const categoriasIngresos = ss.getSheetByName('CategoriasIngresos') || ss.insertSheet('CategoriasIngresos');
  if (categoriasIngresos.getLastRow() === 0) {
    categoriasIngresos.appendRow(['Categoría']);
    CATEGORIAS_INGRESOS_DEFAULT.forEach(c => categoriasIngresos.appendRow([c]));
  }

  const permitidas = ['Gastos', 'Categorías', 'Presupuesto', 'Ingresos', 'CategoriasIngresos'];
  ss.getSheets().forEach(sheet => {
    if (permitidas.indexOf(sheet.getName()) === -1) {
      ss.deleteSheet(sheet);
    }
  });
}

function getCategorias() {
  const sheet = getSheet().getSheetByName('Categorías');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(String);
}

function addCategoria(nombre) {
  const limpio = String(nombre).trim();
  if (!limpio) throw new Error('El nombre no puede estar vacío');

  const existentes = getCategorias();
  const yaExiste = existentes.some(c => c.toLowerCase() === limpio.toLowerCase());
  if (yaExiste) throw new Error('Esa categoría ya existe');

  const ss = getSheet();
  ss.getSheetByName('Categorías').appendRow([limpio]);
  ss.getSheetByName('Presupuesto').appendRow([limpio, 0]);

  return { ok: true, categoria: limpio };
}

function addExpense(fecha, categoria, descripcion, monto) {
  const sheet = getSheet().getSheetByName('Gastos');
  sheet.appendRow([new Date(fecha), categoria, descripcion, parseFloat(monto)]);
  return { ok: true };
}

function getCategoriasIngresos() {
  const sheet = getSheet().getSheetByName('CategoriasIngresos');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(row => row[0])
    .filter(String);
}

function addCategoriaIngreso(nombre) {
  const limpio = String(nombre).trim();
  if (!limpio) throw new Error('El nombre no puede estar vacío');

  const existentes = getCategoriasIngresos();
  const yaExiste = existentes.some(c => c.toLowerCase() === limpio.toLowerCase());
  if (yaExiste) throw new Error('Esa categoría ya existe');

  getSheet().getSheetByName('CategoriasIngresos').appendRow([limpio]);
  return { ok: true, categoria: limpio };
}

function addIngreso(fecha, categoria, descripcion, monto) {
  const sheet = getSheet().getSheetByName('Ingresos');
  sheet.appendRow([new Date(fecha), categoria, descripcion, parseFloat(monto)]);
  return { ok: true };
}

function getPresupuesto() {
  const sheet = getSheet().getSheetByName('Presupuesto');
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  sheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(row => {
    if (row[0]) map[row[0]] = Number(row[1]) || 0;
  });
  return map;
}

function sumarPorMes(sheet, y, m) {
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

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

  return { totales, total };
}

function getResumen(anio, mes) {
  const now = new Date();
  const y = anio || now.getFullYear();
  const m = mes || (now.getMonth() + 1);

  const gastos = sumarPorMes(getSheet().getSheetByName('Gastos'), y, m);
  const ingresos = sumarPorMes(getSheet().getSheetByName('Ingresos'), y, m);

  const total = gastos.total;
  const totalIngresos = ingresos.total;
  const disponible = totalIngresos - total;
  const porcentajeGastado = totalIngresos > 0 ? (total / totalIngresos * 100) : null;

  const presupuesto = getPresupuesto();
  const categorias = getCategorias();
  const detalle = categorias.map(cat => {
    const gastado = gastos.totales[cat] || 0;
    return {
      categoria: cat,
      gastado: gastado,
      presupuesto: presupuesto[cat] || 0,
      porcentajeIngreso: totalIngresos > 0 ? (gastado / totalIngresos * 100) : null
    };
  });

  return { anio: y, mes: m, total, totalIngresos, disponible, porcentajeGastado, detalle };
}
