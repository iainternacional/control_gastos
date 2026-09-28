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
  const page = param === 'resumen' ? 'Resumen' : param === 'ingreso' ? 'Ingreso' : param === 'historial' ? 'Historial' : param === 'categorias' ? 'Categorias' : param === 'presupuesto' ? 'Presupuesto' : 'Formulario';
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

function getPresupuestoDetalle() {
  const presupuesto = getPresupuesto();
  return getCategorias().map(cat => ({ categoria: cat, limite: presupuesto[cat] || 0 }));
}

function setPresupuestos(valores) {
  const sheet = getSheet().getSheetByName('Presupuesto');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };

  const rango = sheet.getRange(2, 1, lastRow - 1, 2);
  const datos = rango.getValues();
  let cambios = false;

  datos.forEach(row => {
    const cat = row[0];
    if (Object.prototype.hasOwnProperty.call(valores, cat)) {
      const nuevoLimite = Number(valores[cat]) || 0;
      if (row[1] !== nuevoLimite) {
        row[1] = nuevoLimite;
        cambios = true;
      }
    }
  });

  if (cambios) rango.setValues(datos);
  return { ok: true };
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
  const categoriasHuerfanas = Object.keys(gastos.totales).filter(cat => categorias.indexOf(cat) === -1);
  const detalle = categorias.concat(categoriasHuerfanas).map(cat => {
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

function movimientosDelMes(nombreHoja, anio, mes) {
  const now = new Date();
  const y = anio || now.getFullYear();
  const m = mes || (now.getMonth() + 1);

  const sheet = getSheet().getSheetByName(nombreHoja);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const movimientos = [];

  data.forEach((row, i) => {
    const fecha = row[0];
    if (!(fecha instanceof Date)) return;
    if (fecha.getFullYear() === y && fecha.getMonth() + 1 === m) {
      movimientos.push({
        fila: i + 2,
        fecha: fecha.toISOString(),
        categoria: row[1],
        descripcion: row[2],
        monto: Number(row[3]) || 0
      });
    }
  });

  movimientos.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
  return movimientos;
}

function getGastos(anio, mes) {
  return movimientosDelMes('Gastos', anio, mes);
}

function getIngresos(anio, mes) {
  return movimientosDelMes('Ingresos', anio, mes);
}

function eliminarFila(nombreHoja, fila) {
  const sheet = getSheet().getSheetByName(nombreHoja);
  const f = parseInt(fila, 10);
  if (!f || f < 2 || f > sheet.getLastRow()) throw new Error('Movimiento no encontrado');
  sheet.deleteRow(f);
  return { ok: true };
}

function deleteExpense(fila) {
  return eliminarFila('Gastos', fila);
}

function deleteIngreso(fila) {
  return eliminarFila('Ingresos', fila);
}

function buscarFilaPorValor(sheet, columna, valor) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const valores = sheet.getRange(2, columna, lastRow - 1, 1).getValues();
  const buscado = String(valor).trim().toLowerCase();
  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim().toLowerCase() === buscado) return i + 2;
  }
  return -1;
}

function reemplazarCategoriaEnMovimientos(sheet, actual, nuevo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rango = sheet.getRange(2, 2, lastRow - 1, 1);
  const valores = rango.getValues();
  const buscado = actual.toLowerCase();
  let cambios = false;
  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]).trim().toLowerCase() === buscado) {
      valores[i][0] = nuevo;
      cambios = true;
    }
  }
  if (cambios) rango.setValues(valores);
}

function renameCategoria(actual, nuevo) {
  const actualLimpio = String(actual).trim();
  const nuevoLimpio = String(nuevo).trim();
  if (!nuevoLimpio) throw new Error('El nombre no puede estar vacío');
  if (actualLimpio.toLowerCase() === nuevoLimpio.toLowerCase()) return { ok: true, categoria: nuevoLimpio };

  const ss = getSheet();
  const categorias = ss.getSheetByName('Categorías');
  const fila = buscarFilaPorValor(categorias, 1, actualLimpio);
  if (fila === -1) throw new Error('Esa categoría no existe');

  const yaExiste = getCategorias().some(c => c.toLowerCase() === nuevoLimpio.toLowerCase());
  if (yaExiste) throw new Error('Ya existe una categoría con ese nombre');

  categorias.getRange(fila, 1).setValue(nuevoLimpio);

  const presupuesto = ss.getSheetByName('Presupuesto');
  const filaPresupuesto = buscarFilaPorValor(presupuesto, 1, actualLimpio);
  if (filaPresupuesto !== -1) presupuesto.getRange(filaPresupuesto, 1).setValue(nuevoLimpio);

  reemplazarCategoriaEnMovimientos(ss.getSheetByName('Gastos'), actualLimpio, nuevoLimpio);

  return { ok: true, categoria: nuevoLimpio };
}

function deleteCategoria(nombre) {
  const limpio = String(nombre).trim();
  const ss = getSheet();
  const categorias = ss.getSheetByName('Categorías');
  const fila = buscarFilaPorValor(categorias, 1, limpio);
  if (fila === -1) throw new Error('Esa categoría no existe');
  categorias.deleteRow(fila);

  const presupuesto = ss.getSheetByName('Presupuesto');
  const filaPresupuesto = buscarFilaPorValor(presupuesto, 1, limpio);
  if (filaPresupuesto !== -1) presupuesto.deleteRow(filaPresupuesto);

  return { ok: true };
}

function renameCategoriaIngreso(actual, nuevo) {
  const actualLimpio = String(actual).trim();
  const nuevoLimpio = String(nuevo).trim();
  if (!nuevoLimpio) throw new Error('El nombre no puede estar vacío');
  if (actualLimpio.toLowerCase() === nuevoLimpio.toLowerCase()) return { ok: true, categoria: nuevoLimpio };

  const ss = getSheet();
  const categorias = ss.getSheetByName('CategoriasIngresos');
  const fila = buscarFilaPorValor(categorias, 1, actualLimpio);
  if (fila === -1) throw new Error('Esa categoría no existe');

  const yaExiste = getCategoriasIngresos().some(c => c.toLowerCase() === nuevoLimpio.toLowerCase());
  if (yaExiste) throw new Error('Ya existe una categoría con ese nombre');

  categorias.getRange(fila, 1).setValue(nuevoLimpio);

  reemplazarCategoriaEnMovimientos(ss.getSheetByName('Ingresos'), actualLimpio, nuevoLimpio);

  return { ok: true, categoria: nuevoLimpio };
}

function deleteCategoriaIngreso(nombre) {
  const limpio = String(nombre).trim();
  const ss = getSheet();
  const categorias = ss.getSheetByName('CategoriasIngresos');
  const fila = buscarFilaPorValor(categorias, 1, limpio);
  if (fila === -1) throw new Error('Esa categoría no existe');
  categorias.deleteRow(fila);
  return { ok: true };
}
