function elimina() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('503:503').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('E503'));
};

function elimi() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('503:503').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('E503'));
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};

function Macrosenzatitolo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('503:503').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};