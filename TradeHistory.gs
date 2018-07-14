//This function adds a new line to the Trade History and increments the Trade ID
function AddNewTrade() {
  
  const sheetName = "Transaction History"
  const namedRangeTradeHeader = "TH_HEADER"
  
  //Get Spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  
  //Get Sheet & Find Line to Enter Trade On
  var sheet = spreadsheet.getSheetByName(sheetName);
  var lastTradeCell = spreadsheet.getRangeByName(namedRangeTradeHeader).getCell(1, 1).offset(2, 0); //Offset 2 lines (one header + one hidden 'start' row)
  var lastTradeNum = parseInt(lastTradeCell.getValue());
  
  //Insert New Row For Trade and Update Trade Number and Date
  sheet.insertRowBefore(lastTradeCell.getRowIndex());
  newTradeNum = lastTradeNum + 1;
  lastTradeCell.setValue(lastTradeNum+1);
  lastTradeCell.offset(0,1).setValue(Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss"));
  
  return lastTradeCell.getRowIndex();
}
