

function eliminariga() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('5:5').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};