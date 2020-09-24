/**
 * Shows an alert dialog.
 */
function showAlert(countProducts, categoryName, clientId, clientName) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
    'Confirme a operação',
    'Você está prestes a subir ' +
    countProducts + ' produtos da categoria ' + 
    categoryName + ' do cliente ' + 
    clientName + ' com id ' +
    clientId + '.\n\n Confirma a ação?',
    ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result === ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}
