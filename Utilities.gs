//Only works for 1-dimensional ranges
//Checks if the text exists within the range
function validate(text, rangeName) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var range = ss.getRangeByName(rangeName);
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  var cell;
  var listText;
  var rangeValues = range.getValues();
  //var ui = SpreadsheetApp.getUi();

  if(cols == 1){
    /*
    for(var i = 1; i <= rows; i++) {
      cell = range.getCell(i, 1);
      listText = cell.getDisplayValue();
    
      if(text == listText) { return true; }
    }
    */
    
    for(var i = 0; i < rows; i++) {
      listText = rangeValues[i][0];
      if(text == listText) return true;                                         
    }
  }
  else if(rows == 1){
    /*
    for(var i = 1; i <= cols; i++) {
      cell = range.getCell(1, i);
      listText = cell.getDisplayValue();
    
      if(text == listText) { return true; }
    }
    */
    for(var i = 0; i < cols; i++) {
      listText = rangeValues[0][i];
      if(text == listText) return true;                                         
    }
  }
  return false;  
}

//Prompts for input and checks if the inputted text/data is in the specified range
function getAndValidateInput(dataName, validationRangeName) {
  
  //var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ui = SpreadsheetApp.getUi();
  
  //Initial Prompt
  var valid = false;
  var prompt = "Enter " + dataName;

  //Keep looping until valid input or cancelled 
  while(valid == false)
  {
    var response = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      var responseText = response.getResponseText();
      
      //Change Coin to All Caps
      if(dataName == "Coin Ticker") responseText = responseText.toUpperCase();
      Logger.log(responseText);
      
      valid = validate(responseText, validationRangeName);
      }
    else
      return false; //User Cancelled
    
    //New prompt specifying invalid input
    prompt = dataName + " '" + responseText + "' not valid. Please enter again.";
  }
  
  //Valid input
  return responseText;
}

//Prompts for input (number) and checks if the inputted text/data is in the specified range
function getAndValidateNumber(dataName) {
  
  //var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ui = SpreadsheetApp.getUi();
  
  //Initial Prompt
  var valid = false;
  var prompt = "Enter " + dataName;

  //Keep looping until valid input or cancelled 
  while(valid == false)
  {
    var response = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.OK) {
      var responseText = response.getResponseText();
      valid = !isNaN(parseFloat(responseText)) && isFinite(responseText);
      }
    else
      return "cancel"; //User Cancelled
    
    //New prompt specifying invalid input
    prompt = dataName + " '" + responseText + "' not valid. Please enter again.";
  }
  
  //Valid input
  responseText = parseFloat(responseText);
  return responseText;
}

//Checks a given range for an errors (i.e. starts with # or has a value of 0.0009999)
function errorExists(rangeName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var range = ss.getRangeByName(rangeName);
  var rows = range.getNumRows();
  var cols = range.getNumColumns();

  for(var c = 1; c <= cols; c++) {
    for(var r = 1; r <= rows; r++) {
      var cell = range.getCell(r, c);
      var cellValue = cell.getValue();
      var cellText = cell.getDisplayValue();

      
      if(cellValue == 0.0009999 || cellText.charAt(0) == "#" || cellText == "0.0009999") 
        return true;
    }   
  }
  
  return false;  

}

//Checks if the tickers match (text and order) on the Holdings and Closes sheets
function checkTickers(closeType)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Closes");
  var tickersHoldingsRange = ss.getRangeByName("H_TICKERS");
  var tickersClosePricesRange = ss.getRangeByName("C_PRICE_TICKERS");
  var tickersCloseBalancesRange = ss.getRangeByName("C_BALANCE_TICKERS");
  var investorsCloseRange = ss.getRangeByName("C_INVESTORS");
  var investorsEquityRange = ss.getRangeByName("E_USERS");
  
  var holdingsRangeSize = tickersHoldingsRange.getNumRows();
  var pricesRangeSize = tickersClosePricesRange.getNumColumns();
  var balancesRangeSize = tickersCloseBalancesRange.getNumColumns();

  var numCoinsHoldings = parseInt(ss.getRangeByName("MD_COINS_HOLDINGS").getValue());
  var numCoinsCloses = parseInt(ss.getRangeByName("MD_COINS_CLOSES").getValue());
  var numNewCoins = numCoinsHoldings - numCoinsCloses;
 
  var numInvestors = parseInt(ss.getRangeByName("MD_NUM_INVESTORS").getValue());
  var numInvestorsCloses = parseInt(ss.getRangeByName("MD_NUM_INVESTORS_CLOSES").getValue());
  
  var tickerHolding;
  var tickerCPrice;
  var tickerCBalance;
  
  var investorCloseSheet;
  var investorEquitySheet;
  
  if(closeType != "Auto") var ui = SpreadsheetApp.getUi();
  
  SpreadsheetApp.flush();
  
  //Check if sizes match (taking into account new coins)
  if(holdingsRangeSize != pricesRangeSize + numNewCoins || holdingsRangeSize != balancesRangeSize + numNewCoins){
    if(closeType != "Auto") 
      ui.alert("ERROR: Size of Range doesn't match on Holdings, Close Prices, Close Balances Ranges. Aborting.");
    else
      closeErrorEmail("ERROR: Size of Range doesn't match on Holdings, Close Prices, Close Balances Ranges. Aborting.");
    return false;
  }
  
  if(numInvestorsCloses > numInvestors){
    if(closeType != "Auto") 
      ui.alert("ERROR: More investors on closes sheet than on the Equity sheet. Aborting.");
    else
      closeErrorEmail("ERROR: More investors on closes sheet than on the Equity sheet. Aborting.");
    return false;
  }
  
  //Compare the tickers from each relevant section, not including new coins
  for(i = 1; i <= numCoinsHoldings - numNewCoins; i++){
    tickerHolding = tickersHoldingsRange.getCell(i, 1).getDisplayValue();
    tickerCPrice = tickersClosePricesRange.getCell(1, i).getDisplayValue();
    tickerCBalance = tickersCloseBalancesRange.getCell(1, i).getDisplayValue();
    
    if(tickerHolding != tickerCPrice || tickerHolding != tickerCBalance){
      if(closeType != "Auto")
        ui.alert("ERROR: Tickers don't match on Holdings, Close Prices, Close Balances. Aborting. [" + tickerHolding + "," + tickerCPrice + "," + tickerCBalance + "]");
      else
        closeErrorEmail("ERROR: Tickers don't match on Holdings, Close Prices, Close Balances. Aborting. [" + tickerHolding + "," + tickerCPrice + "," + tickerCBalance + "]");
      return false;
    }
  }
  
  //Check Users
  for(i = 1; i <= numInvestorsCloses; i ++){
    investorEquitySheet = investorsEquityRange.getCell(i, 1).getDisplayValue();
    investorCloseSheet = investorsCloseRange.getCell(1, i).getDisplayValue();
    
    if(investorEquitySheet != investorCloseSheet){
      if(closeType != "Auto")
        ui.alert("ERROR: Investors don't match on Closes sheet and Equity sheet. Aborting. [" + investorEquitySheet + "," + investorCloseSheet + "]");
      else
        closeErrorEmail("ERROR: Investors don't match on Closes sheet and Equity sheet. Aborting. [" + investorEquitySheet + "," + investorCloseSheet + "]");
      return false;
    }   
  }
  return true;
}

//Gets the last closePrice for a given ticker, from a 'Standard Close' (i.e. not a deposit or withdrawal)
function getClosePriceDate(ticker, rangeName){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var range = ss.getRangeByName(rangeName);
  var numCoinsCloses = parseInt(ss.getRangeByName("MD_COINS_CLOSES").getValue());  
  var offsetToLastClose = 0;
  var cell;
  var liveCellCloseSheet =  ss.getRangeByName("C_LIVE");
  var closeDate
  var offset = 1;
  var closeType;
  
  while(offsetToLastClose == 0){
    closeType = liveCellCloseSheet.getCell(1, 1).offset(offset, 1).getDisplayValue(); //Check the 2nd column (Close Type), starting from the 2nd row (offset is initialized to 2)
    
    if(closeType == "Close") {
      offsetToLastClose = offset + 1; //Since we are counting from the LIVE row, not the Tickers row (which is used to find the cell;
      closeDate = liveCellCloseSheet.getCell(1, 1).offset(offset, 0).getDisplayValue(); //This is the close date
    }
    else
      offset++;
  }
  
  for(var i=1; i <= numCoinsCloses; i++)
  {
    cell = range.getCell(1, i);
    if(cell.getDisplayValue() == ticker){
      var closePrice = parseFloat(cell.offset(offsetToLastClose, 0).getValue());
      return [closePrice, closeDate]; 
    }
  }
}


//Gets the current equity value in USD for a given userID from the Equity sheet
function getEquityValue(userID){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var range = ss.getRangeByName("E_USERS");
  var numInvestors = parseInt(ss.getRangeByName("MD_NUM_INVESTORS").getValue()); 
  var cell;
  
  for(var i=1; i <= numInvestors; i++){
    cell = range.getCell(i, 1);
    
    if(cell.getDisplayValue() == userID){
      return parseFloat(cell.offset(0,3).getValue());
    }
  }
}

function testLivePrice(){
  var ui = SpreadsheetApp.getUi();
  ui.alert("XMR" + getLivePrice("XMR"));
  ui.alert("XLM" + getLivePrice("XLM"));
  ui.alert("BTC" + getLivePrice("BTC"));


  
}

function getLivePrice(ticker){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var range = ss.getRangeByName("H_TICKERS");
  var sheet = ss.getSheetByName("HOLDINGS");
  var numInvestors = parseInt(ss.getRangeByName("MD_COINS_HOLDINGS").getValue());
  var pricesCol = ss.getRangeByName("H_PRICES_COL").getColumn();
  var cell;
  
  ticker = ticker.toUpperCase();
  
  for(var i=1; i <= numInvestors; i++){
    cell = range.getCell(i,1);
    
    if(cell.getDisplayValue() == ticker){
      return parseFloat(sheet.getRange(cell.getRowIndex(), pricesCol).getValue());
    }
  }
}

//Returns the balance of a coin. Returns 0 if the coin is not found
//Uses the LIVE row on the Close sheet for the balance
function checkBalance(coin){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var ui = SpreadsheetApp.getUi();
  
  var coin = coin.toUpperCase();
  var balance;
  var ticker;
  
  var tickers =  ss.getRangeByName("C_BALANCE_TICKERS").getValues();
  var balances =  ss.getRangeByName("C_BALANCES").getValues();
  var numCoinsCloses = ss.getRangeByName("C_BALANCE_TICKERS").getNumColumns();
  
  //Go through the (balances) tickers on close sheet. Get the balance for the requested coin.
  for(var i = 0; i < numCoinsCloses; i++){
    ticker = tickers[0][i];
    
    if(ticker == coin) {
      balance = balances[0][i];
      return balance;
    }
  }
  return 0;
}



//Returns the balance of a coin. Returns 0 if the coin is not found
//Uses the LIVE row on the Close sheet for the balance
function getAllocation(coin){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var coin = coin.toUpperCase();
  var allocation;
  var ticker;
  
  var coinAllocations =  ss.getRangeByName("H_COIN_ALLOCATIONS").getValues();
  var numCoins = ss.getRangeByName("H_COIN_ALLOCATIONS").getNumRows();
  
  //Go through the coins on holdings sheet. Get the balance for the requested coin.
  for(var i = 0; i < numCoins; i++){
    ticker = coinAllocations[i][0];
    
    if(ticker == coin) {
      allocation = coinAllocations[i][1];
      return allocation;
    }
  }
  return 0;
}

/*
//Returns the balance of a coin. Returns 0 if the coin is not found
//Uses the LIVE row on the Close sheet for the balance
function checkBalanceOld(coin){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var ui = SpreadsheetApp.getUi();
  
  var balance;
  var ticker;
  var balancesRange =  ss.getRangeByName("C_BALANCE_TICKERS");
  var numCoinsCloses = parseInt(ss.getRangeByName("MD_COINS_CLOSES").getValue());    
  
  //Go through the (balances) tickers on close sheet. Get the balance for the requested coin.
  for(var i = 1; i <= numCoinsCloses; i++){
    ticker = balancesRange.getCell(1, i).getDisplayValue();
    
    if(ticker == coin) {
      balance = parseFloat(balancesRange.getCell(1, i).offset(1, 0).getValue());
      return balance;
    }
  }
  return 0;
}
*/




