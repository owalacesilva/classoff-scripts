var URL_API = 'https://localhost/api';
var CLIENT_ID = null;
var CLIENT_NAME = '';

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Upload Products Variations ' + CLIENT_NAME, [
    { name: 'Upload Peças', functionName: 'exportParts'},
    { name: 'Upload Motos', functionName: 'exportVehicles'},
    { name: 'Upload Equipamentos', functionName: 'exportEquipaments'},
    { name: 'Upload Acessorios', functionName: 'exportAccessories'},
    { name: 'Upload Oleos', functionName: 'exportOils'},
    { name: 'Upload Casual', functionName: 'exportCasual'}
  ]);
}

/**
 * Exporta os produtos da planilha
 */
function exportVehicles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Motos');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Motos', CLIENT_ID, CLIENT_NAME);

  exportProducts(1, sheet, values);
}

function exportEquipaments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Equipamentos');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Equipamentos', CLIENT_ID, CLIENT_NAME);

  exportProducts(2, sheet, values);
}

function exportParts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Peças');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Peças', CLIENT_ID, CLIENT_NAME);

  exportProducts(3, sheet, values);
}

function exportAccessories() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Acessorios');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Acessorios', CLIENT_ID, CLIENT_NAME);

  exportProducts(4, sheet, values);
}

function exportOils() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Oleos');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Oleos', CLIENT_ID, CLIENT_NAME);

  exportProducts(5, sheet, values);
}

function exportCasual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Casual');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  showAlert(values.length, 'Casual', CLIENT_ID, CLIENT_NAME);

  exportProducts(6, sheet, values);
}

/**
 * Realiza o carregamento do resource para API
 * 
 * @param {*} url 
 * @param {*} resource 
 */
function uploadProduct(url, resource) {
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(resource)
  }  
    
  var res = UrlFetchApp.fetch(url, options)
  if (res.getResponseCode() == 201) {
    return JSON.parse(res.getContentText());
  } else {
    SpreadsheetApp.getUi().alert('Something happen wrong here!!!');
    return false;
  }
}

/**
 * Agrupa todos as variações de produtos 
 * usando o codigo de referencia.
 * 
 * @param {*} rows 
 */
function groupProductRows(rows) {
  result = rows.reduce(function(group, row) {
    var rowValue = '' + row[0]; // Código do Produto
    
    if (typeof group[rowValue] == 'undefined') {
      group[rowValue] = [];
    }
    
    group[rowValue].push(row);
    return group; 
  }, {});
  
  return result;
}

function exportProducts(categoryId, sheet, values) {
  var matrix = sheet.getDataRange().getValues();
  var rowsWithoutHeader = matrix.slice(1);
  var productGroup = groupProductRows(rowsWithoutHeader);
  
  for (var productRef in productGroup) {
    var preProduct = prepareProduct(categoryId, productGroup[productRef]);
    preProduct['category_id'] = categoryId;
    
    // SpreadsheetApp.getUi().alert(JSON.stringify(preProduct));    
    uploadProduct(URL_API + '/products/bulk', preProduct);
  }
  
  SpreadsheetApp.getUi().alert('Operation finished');
}

/**
 * 
 * @param {*} row 
 */
function prepareProductVariantResource(row) {
  var stock_quantity = row[12]; // (M) Estoque

  var images = [];
  if ( typeof row[23] == 'string' ) {
    var images = row[23].split(';'); // (X) Conjunto de Images
  }

  var attribute_terms = [];
  if ( typeof row[10] == 'string' || typeof row[10] == 'number' ) {
    attribute_terms.push({
      "type": "product_size",
      "name": row[10], // (K) Tamanho
    });
  }

  if ( typeof row[11] == 'string' ) {
    attribute_terms.push({
      "type": "product_color",
      "name": row[11], // (L) Cor
    });
  }
  
  return {
	  "sku": row[1], // (B) Código Interno
	  "stock_quantity": stock_quantity,
    "stock_status": stock_quantity > 0 ? 'instock' : null,
    "vendor_internal_code": row[1], // (B) Código Interno
    "price": row[13], // (N) Preço
    "weight": row[18], // (S) Peso
    "depth": row[19], // (T) Comprimento
    "width": row[20], // (U) Largura
    "height": row[21], // (V) Altura
    "permalink": row[1], // (B) Código Interno
    "attribute_terms": attribute_terms,
    "images": images,
    "image": images.length > 0 ? images[0] : null
  }
}

/**
 * 
 * @param {*} categoryId 
 * @param {*} row 
 */
function prepareProductResource(categoryId, row) {
  return {      
    "account_id": CLIENT_ID, // ID Conta
    "category_id": categoryId, // ID Categoria
    "vendor_internal_code": row[1], // (B) Código Interno
    "product_brands": row[2] ? row[2].split(',') : [], // (C) Marca da Moto
    "product_models": row[3] ? row[3].split(',') : [], // (D) Modelo da Moto
    "product_years": row[4] ? row[4].split(';') : [], // (E) Ano da Moto
    "product_group": row[5], // (F) Grupo do Produto
    "product_type": row[6], // (G) ID Item
    "title": row[7], // (H) Nome do Produto
    "short_description": row[8], // (I) Breve descrição
    "description": row[9], // (J) Descrição
    "price": row[13], // (N) Preço
    "discount": row[14], // (O) % Desconto (Clube Veloce)
    "product_used": row[15], // (P) Usado 1 /Novo 0
    "accept_swap": row[16], // (Q) Aceita trocas? 0N - 1S
    "has_installments": row[17], // (R) Aceita parcelamentos? 0N - 1S
    "weight": row[18], // (S) Peso
    "depth": row[19], // (T) Comprimento
    "width": row[20], // (U) Largura
    "height": row[21], // (V) Altura
    "product_wear": row[22], // (W) Desgaste
    "images": row[23] ? row[23].split(';') : [] // (X) Conjunto de Images
  }
}

/**
 * 
 * @param {*} categoryId 
 * @param {*} rows 
 */
function prepareProduct(categoryId, rows) {
  var product = prepareProductResource(categoryId, rows[0]);
  product['variants_attributes'] = [];
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    product['variants_attributes'].push(prepareProductVariantResource(row));
  }

  return product;
}
