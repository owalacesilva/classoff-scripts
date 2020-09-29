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
    var rowValue = '' + row[1]; // (B) Código do Produto
    
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
    
    try {
      var product = uploadProduct(URL_API + '/products/bulk', preProduct);

      for(var i = 0; i < rowsWithoutHeader.length; i++) {
        if (rowsWithoutHeader[i][1] === productRef) {
          sheet.getRange('A' + (i + 2)).setValue(product.id);
        }
      }
    } catch (e) {
      SpreadsheetApp.getUi().alert(e);
    }
  }
  
  SpreadsheetApp.getUi().alert('Operation finished');
}

/**
 * 
 * @param {*} row 
 */
function prepareProductVariantResource(row) {
  var stock_quantity = row[13]; // (N) Estoque

  var images = [];
  if ( typeof row[24] == 'string' ) {
    var images = row[24].split(';'); // (Y) Conjunto de Images
  }

  var attribute_terms = [];
  if ( (typeof row[11] == 'string' && row[11].length > 0) || typeof row[11] == 'number' ) {
    attribute_terms.push({
      "type": "product_size",
      "name": row[11], // (L) Tamanho
    });
  }

  if ( typeof row[12] == 'string' && row[12].length > 0 ) {
    attribute_terms.push({
      "type": "product_color",
      "name": row[12], // (M) Cor
    });
  }
  
  return {
	  "sku": row[2], // (C) Código Interno
	  "stock_quantity": stock_quantity,
    "stock_status": stock_quantity > 0 ? 'instock' : null,
    "vendor_internal_code": row[2], // (C) Código Interno
    "price": row[14], // (O) Preço
    "weight": row[15], // (T) Peso
    "depth": row[20], // (U) Comprimento
    "width": row[21], // (V) Largura
    "height": row[22], // (W) Altura
    "permalink": row[2], // (C) Código Interno
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
    "vendor_internal_code": row[2], // (C) Código Interno
    "product_brands": row[3] ? row[3].split(',') : [], // (D) Marca da Moto
    "product_models": row[4] ? row[4].split(',') : [], // (E) Modelo da Moto
    "product_years": row[5] ? row[5].split(';') : [], // (F) Ano da Moto
    "product_group": row[6], // (G) Grupo do Produto
    "product_type": row[7], // (H) ID Item
    "title": row[8], // (I) Nome do Produto
    "short_description": row[9], // (J) Breve descrição
    "description": row[10], // (K) Descrição
    "price": row[14], // (O) Preço
    "discount": row[15], // (P) % Desconto (Clube Veloce)
    "product_used": row[16], // (Q) Usado 1 /Novo 0
    "accept_swap": row[17], // (R) Aceita trocas? 0N - 1S
    "has_installments": row[18], // (S) Aceita parcelamentos? 0N - 1S
    "weight": row[19], // (T) Peso
    "length": row[20], // (U) Comprimento
    "width": row[21], // (V) Largura
    "height": row[22], // (W) Altura
    "product_wear": row[23], // (X) Desgaste
    "images": row[24] ? row[24].split(';') : [] // (Y) Conjunto de Images
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
