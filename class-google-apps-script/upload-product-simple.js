var URL_API = 'https://localhost/api';
var CLIENT_ID = null;
var CLIENT_NAME = '';

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Upload Products ' + CLIENT_NAME, [
    { name: 'Upload Peças', functionName: 'exportParts'},
    { name: 'Upload Motos', functionName: 'exportVehicles'},
    { name: 'Upload Equipamentos', functionName: 'exportEquipaments'},
    { name: 'Upload Acessorios', functionName: 'exportAccessories'},
    { name: 'Upload Oleos', functionName: 'exportOils'},
    { name: 'Upload Casual', functionName: 'exportCasual'}
  ]);
}

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

function prepareProductResource(categoryId, row) {
  return {
	"account_id": CLIENT_ID,
	"category_id": categoryId,
    "vendor_internal_code": row[1],
    "product_brands": row[2] ? row[2].split(',') : [],
    "product_models": row[3] ? row[3].split(',') : [],
    "product_years": row[4] ? row[4].split(';') : [],
    "product_group": row[5],
    "product_type": row[6],
    "title": row[7],
    "short_description": row[8],
    "description": row[9],
    "price": row[10],
    "discount": row[11],
    "product_used": row[12],
    "accept_swap": row[13],
    "has_installments": row[14],
    "stock_quantity": row[15],
    "weight": row[16],
    "length": row[17],
    "width": row[18],
    "height": row[19],
    "product_wear": row[20],    
    "images": row[21] ? row[21].split(',') : []
  }
}

function uploadResource(url, resource) {
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(resource)
  }  
    
  var res = UrlFetchApp.fetch(url, options)
  if (res.getResponseCode() == 201) {
    return JSON.parse(res.getContentText()).id;
  } else {
    SpreadsheetApp.getUi().alert('Something happen wrong here!!!');
    return false;
  }
}

function exportProducts(categoryId, sheet, values) {    
  for(var i = 1; i < values.length; i++) {
    var rowId = values[i][0];

    if (!rowId) {
      try {
        // create it
        var resource = prepareProductResource(categoryId, values[i]);
        var newId = uploadResource(URL_API + '/products/bulk', resource);
        sheet.getRange('A' + (i + 1)).setValue(newId);
      } catch (e) {
        SpreadsheetApp.getUi().alert('Produto na linha ' + (i + 1) + ' não pode ser salvo: ' + e);
      }
    } else if (typeof rowId == 'number') {
      // update it
    } else {
      // cancel
      SpreadsheetApp.getUi().alert("Invalid ID: " + rowId);
    }
  }
  
  SpreadsheetApp.getUi().alert('Operation finished');
}
