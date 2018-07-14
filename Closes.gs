//Regular Closing Process
function StandardClose(){
  Close("Close","");
}

//Regular Closing Process
function AutoClose(){
  Close("Auto","");
}

//This performs the closing process which is used for daily recordkeeping and recordkeeping during deposits/withdrawals
//Records current prices, balances, total value, and equity to the table on the closes sheet
function Close(closeType, dwID)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName = "Closes";
  var sheet = ss.getSheetByName(sheetName);
  var namedRangeLive = "C_LIVE";
  var numCoinsHoldings = parseInt(ss.getRangeByName("MD_COINS_HOLDINGS").getValue());
  var numCoinsCloses = parseInt(ss.getRangeByName("MD_COINS_CLOSES").getValue());  
  var numInvestors = parseInt(ss.getRangeByName("MD_NUM_INVESTORS").getValue());
  var numInvestorsCloses = parseInt(ss.getRangeByName("MD_NUM_INVESTORS_CLOSES").getValue());
  var rangeName;  
  var range;
  var cell;
  var value;
  var newCell;
  var totalCells;
  var confirmClose = false;
  var closeDescription;
  
  //Close Description. Needed to distinguish between type and description when we added weekly p&l. 
  closeDescription = closeType;
  if(closeType == "Weekly P&L")
    closeType = "Auto";
    
    

  //Prompt for confirmation if Standard Close
  if(closeType != "Auto") var ui = SpreadsheetApp.getUi();    
  
  if (closeType == "Close"){
    var response = ui.alert('Are You Sure You Want To Close?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) 
      confirmClose = true;
  }
  else 
    confirmClose = true;

  // Process the user's response.
  if (confirmClose == true) {
    
    //Check if Tickers match across holdings and closes 
    var check = checkTickers(closeType);
    if(check == false) {
      //if(closeType == "Auto") closeErrorEmail();
      return;
    }
    
    //Check if any new Coins have been added, add to Closes if necessary
    if(numCoinsHoldings != numCoinsCloses || numInvestors != numInvestorsCloses) {
      var success = UpdateTickersCloseSheet();
      
      //Check if user cancelled
      if(success == false) {
        ui.alert('Close Aborted'); 
        return; 
      }
    }
    
    //Refresh Prices
    RefreshDataFeed();
    SpreadsheetApp.flush();
    var message = ""
    if(closeType != "Close") var message = " " + closeType + " completed, but not recorded on Close sheet. Please manually close for this " + closeType;
    
    //Check No Errors On Holdings or Closes Sheet
    if (errorExists("H_TABLE") || errorExists("H_TOTAL_USD") || errorExists("H_TOTAL_CAD") || errorExists("H_TOTAL_BTC") || errorExists("H_TOTAL_ETH")) {
      if(closeType != "Auto") 
        ui.alert('Error detected on Holdings sheet. Aborting Close.' + message, ui.ButtonSet.OK);
      else 
        closeErrorEmail("Error detected on Holdings sheet. Aborting Close.");
      
      return;
    }
    if (errorExists("C_PRICES") || errorExists("C_BALANCES") || errorExists("C_TOTALS")) {
      if(closeType != "Auto")
        ui.alert('Error detected on Closes sheet. Aborting Close.' + message, ui.ButtonSet.OK);
      else
        closeErrorEmail("Error detected on Closes sheet. Aborting Close.");
      return;
    }
      
    //Create New Row for the close and add date
    var lastCloseCell = ss.getRangeByName(namedRangeLive).getCell(1, 1).offset(1, 0); 
    sheet.insertRowBefore(lastCloseCell.getRowIndex());
    lastCloseCell.setValue(Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss"));
    lastCloseCell.setform

    lastCloseCell.offset(0, 1).setValue(closeDescription);
    lastCloseCell.offset(0, 2).setValue(dwID);

    
    //Copy Prices, Balances, Total, and Equity to New Row
    for(var x = 1; x <= 4; x++)
    {
      if(x == 1){
        rangeName = "C_PRICES";
        totalCells = numCoinsHoldings;
      }
      else if(x == 2){
        rangeName = "C_BALANCES";
        totalCells = numCoinsHoldings;
      }
      else if(x == 3){
        rangeName = "C_TOTALS";
        totalCells = 4;
      }
      else if(x == 4){
        rangeName = "C_EQUITY";
        totalCells = numInvestors;
      }
      
      range = ss.getRangeByName(rangeName);  
      
      range.copyTo(range.offset(1, 0), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

/*
      for(var i = 1; i <= totalCells; i++){
        cell = range.getCell(1, i);
        value = cell.getValue();
        newCell = cell.offset(1, 0);
        
        newCell.setValue(value);
        
      }
      */
    } 
    
    if(closeType != "Auto")
      ui.alert(closeType + " Completed");

  }
  
}

//Updates the tickers on the close sheet by comparing the tickers on the Holdings sheet to the ones on prices/balances on Close sheet
function UpdateTickersCloseSheet()
{ 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Closes");
  var numCoinsHoldings = parseInt(ss.getRangeByName("MD_COINS_HOLDINGS").getValue());
  var numInvestors = parseInt(ss.getRangeByName("MD_NUM_INVESTORS").getValue());
  var liveRow = ss.getRangeByName("C_LIVE").getRowIndex();
  var desc;
  var sourceRange;
  var numCols;
  var lastCell;
  var ticker;
  var tickersRange;
  var tickersEndCell;
  
  //Go through the Tickers on Holdings Page and Add Column to Prices/Balances on Closes Sheet if missing
  for(var i=1; i <= 3; i++)
  {
    if(i == 1) {
      tickersRange = "C_PRICE_TICKERS";
      tickersEndCell = "C_PRICES_END";
      numCols = numCoinsHoldings;
      sourceRange = ss.getRangeByName("H_TICKERS");
      desc = "coin";
    }
    else if(i==2){
      tickersRange = "C_BALANCE_TICKERS";
      tickersEndCell = "C_BALANCES_END";
      numCols = numCoinsHoldings;
      sourceRange = ss.getRangeByName("H_TICKERS");
      desc = "coin";
    }
    else if(i==3){
      tickersRange = "C_INVESTORS";
      tickersEndCell = "C_EQUITY_END";
      numCols = numInvestors;
      sourceRange = ss.getRangeByName("E_USERS");
      desc = "investor";
    }
    
    //Last Cell for the current set of tickers
    lastCell = ss.getRangeByName(tickersEndCell).getCell(1, 1).offset(0, -1);
    
    //Compare each ticker on holdings to close sheet
    for(var c=1; c <= numCols; c++){
      ticker = sourceRange.getCell(c, 1).getDisplayValue();
      
      //Check if it doesn't exists on closes
      if(!validate(ticker,tickersRange)) {
        
        //Only prompt once per ticker
        if(i != 2) var response = ui.alert("Adding " + desc + " " + ticker + " to close sheet", ui.ButtonSet.OK_CANCEL);
        
        if(response == ui.Button.OK || i == 2)
        {
            sheet.insertColumnAfter(lastCell.getColumn());
            var newCol = ss.getRangeByName(tickersEndCell).getColumn() - 1;
            
            //Set Ticker and copy formula for the live price, from the column the left 
            sheet.getRange(liveRow - 1, newCol).setValue(ticker);       
            sheet.getRange(liveRow, newCol-1).copyTo(sheet.getRange(liveRow, newCol));
            
            //Update Last Cell (in case we have to add more)
            lastCell = ss.getRangeByName(tickersEndCell).getCell(1, 1).offset(0, -1);
        }
        else
          return false;
      }    
    }
  }
}  

function closeErrorEmail(errorMsg){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var emails = ss.getRangeByName("TM_EMAILS").getDisplayValue();  
  var emails = PropertiesService.getScriptProperties().getProperty("emails")
  
  MailApp.sendEmail({
        to: emails,
        subject: "CryptoTrader: Error on AutoClose. Please Close Manually",
        htmlBody: errorMsg + "<br><br> Sent via <a href=\"https://docs.google.com/spreadsheets/d/1lvafocby-uv-N5P3efAGwJXDGqjWoM_8aBRxA6CBYic/edit#gid=0\">CryptoTrader</a>"                                   
      });  
}