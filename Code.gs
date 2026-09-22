const SHEET_ID = '1j8FOnh8OQyhcgXEQjJv57mnCSS10ASHV4mbj6VA_HyY';
const SHEET_NAME = 'Inventario';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Inventario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover');
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Nombre', 'Cantidad']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getProducts() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const products = [];
  data.forEach((r, i) => {
    if (r[0] !== '') products.push({ row: i + 2, nombre: r[0], cantidad: r[1] });
  });
  return products.reverse(); // más reciente primero
}

function findProduct(nombre) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
      return { row: i + 2, nombre: data[i][0], cantidad: data[i][1] };
    }
  }
  return null;
}

function addProduct(nombre, cantidad) {
  getSheet_().appendRow([nombre, Number(cantidad)]);
  return getProducts();
}

function sumToProduct(row, cantidad) {
  const sheet = getSheet_();
  const current = sheet.getRange(row, 2).getValue();
  sheet.getRange(row, 2).setValue(Number(current) + Number(cantidad));
  return getProducts();
}

function updateProduct(row, nombre, cantidad) {
  getSheet_().getRange(row, 1, 1, 2).setValues([[nombre, Number(cantidad)]]);
  return getProducts();
}

function deleteProduct(row) {
  getSheet_().deleteRow(row);
  return getProducts();
}
