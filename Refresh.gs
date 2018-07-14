//Refresh Prices Feed
function RefreshDataFeed()
{
  RefreshFormulasColumn("DF_FORMULA_1");
  RefreshFormulasColumn("DF_FORMULA_2");
  RefreshFormulasColumn("DF_FORMULA_3");
  RefreshFormulasColumn("DF_FORMULA_4");
  RefreshFormulasColumn("DF_FORMULA_5");
  RefreshFormulasColumn("DF_FORMULA_6");
  RefreshFormulasColumn("DF_FORMULA_7");
  Utilities.sleep(500);
}

//Deletes and Reenters URL to refresh the price feed
function RefreshFormulasColumn(namedRange) {
  
  //Get Spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  
  //Get prices range
  var range = spreadsheet.getRangeByName(namedRange);
  var numRows = range.getNumRows();

  for (var i = 1; i <= numRows; i++) {
    var cell = range.getCell(i, 1);
    var formula = cell.getFormula();
    
    //Clear and Set Cell for it refresh
    cell.clear();
    SpreadsheetApp.flush(); //Commits changes to sheet, otherwise won't refresh data
    cell.setFormula(formula);
    
    //Both refreshes don't come in consistently without the sleep
    
  
  }   
  return "done";
}