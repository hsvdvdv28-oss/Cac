function RDESC() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('U6:AD14').activate();
  spreadsheet.getRange('PORTES!B3:K11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};
