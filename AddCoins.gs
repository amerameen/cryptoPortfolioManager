//Adds a new coin to the holdings sheet (must exist on Data Feed)
function AddCoin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName = "Holdings";
  var sheet = ss.getSheetByName(sheetName);
  var lastTickerRow = ss.getRangeByName("H_TABLE_ENDROW").getCell(1, 1).offset(-1, 0).getRowIndex();
  var newRow = lastTickerRow + 1;
  var numCols = 13;
  
  var ui = SpreadsheetApp.getUi();

  
  //Check to make sure existing tickers match the tickers on the closes page
  //If not, this should be fixed before adding any coins
  if(checkTickers("AddCoin") == false) return;
  
  //Get coin CMC ID
  var coin = getAndValidateInput("Coin ID (e.g. bitcoin)", "DF_IDS");
  if(coin == false) return;

  //Confirm it doesn't already exist on Holdings Page
  if(validate(coin,"H_IDS") == true){
    ui.alert("ERROR: " + coin + " already exists on Holdings sheet");
    return;
  }
    
  //Insert New Row
  sheet.insertRowAfter(lastTickerRow);
  sheet.getRange(newRow, 1).setValue(coin);
  
  //Copy Formulas from Above Row
  for(var col = 2; col <= numCols; col++){
    sheet.getRange(lastTickerRow, col).copyTo(sheet.getRange(newRow, col));
  }
  
  //Add to Closes Sheet
  UpdateTickersCloseSheet();
}
