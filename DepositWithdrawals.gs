//Gets and Verifies the info for a deposit/withdrawal
//Adds an appropriate buy OR sell trade to the Trade History 
//Adds a line with the verified info to the deposit/withdrawal table
//Updates Ownership
function DepositWithdrawal() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ui = SpreadsheetApp.getUi();
  var transactionType = "";
 
  //Get User ID, Coin, Quantity. Abort if user presses cancel at any tie
  //Verify against close sheet, which forces user to close before making a deposit or withdrawal
  var userID = getAndValidateInput("User ID", "E_USERS");  
  if(userID == false) return;
  
  var coin = getAndValidateInput("Coin Ticker", "C_BALANCE_TICKERS");
  if(coin == false) return;
  
  var quantity = getAndValidateNumber("Quantity (negative quantity for a withdrawal)");
  if(quantity == "cancel") return;
  
  if(quantity > 0) 
    transactionType = "Deposit"; 
  else 
    transactionType = "Withdrawal";
  
  //Check Coin Balance to ensure we have enough to withdraw
  if(transactionType == "Withdrawal"){
    var balance = checkBalance(coin);
    if(balance < Math.abs(quantity)) {
      ui.alert("Insufficient Balance (" + balance + ") of " + coin + " for requested withdrawal (" + Math.abs(quantity) + "). Transaction Aborted.");
      return;
    }
  }
  
  //Refresh before getting prices/fund value
  RefreshDataFeed();
  
  //Close Value. Uses Value from Close sheet, but allows manual entry if the user selects 'No' at prompt
  var closesCoin = getClosePriceDate(coin, "C_PRICE_TICKERS");
  var closePrice = closesCoin[0];
  var closeDate = closesCoin[1];
  var closePriceLive = parseFloat(getLivePrice(coin)).toFixed(2);
  var closePriceResponse = ui.alert("Would you like to use " + coin + " close price of $" + parseFloat(closePrice).toFixed(2) + " USD" + " captured on " + closeDate + " (Live Price: $" + closePriceLive + " USD)", ui.ButtonSet.YES_NO_CANCEL);

  //Cancel or Manual Entry of Close Price
  if(closePriceResponse ==  ui.Button.CANCEL) return;
  else if(closePriceResponse ==  ui.Button.NO)
  {
    var closePrice = getAndValidateNumber("Close Price (USD) of " + coin);
    if(closePrice == "cancel") return;
  }
  
  //Transaction Value
  var transactionValue = closePrice * quantity;
  
  //Check if equity value of user is enough for size of withdrawal
  var equityUSD = parseFloat(getEquityValue(userID)).toFixed(2);
  if(transactionType == "Withdrawal" && equityUSD < Math.abs(transactionValue)){
      ui.alert("Insufficient Equity ($" + equityUSD + " USD) for requested withdrawal ($" + parseFloat(Math.abs(transactionValue)).toFixed(2) + " USD). Transaction Aborted.");
      return;
  }
  
    
  //Gets the last 'Close' fund Value. Ignores fund Value from Deposits and Withdrawals 
  var fundValue = (getClosePriceDate("USD", "C_TOTALS_TICKERS"))[0];
  var liveFundValue = ss.getRangeByName("H_TOTAL_USD").getValue();
  var warning = "";
  
  if(liveFundValue/fundValue > 1.05 || fundValue/liveFundValue > 1.05 || liveFundValue == 0) {
    warning = "WARNING: Check difference between Live Fund Value and Close Fund Value. \n";
  }
  var fundValueResponse = ui.alert(warning + "Would you like to use Close Fund Value (before transaction) of $" + parseFloat(fundValue).toFixed(2) + " USD" + " captured on " + closeDate + ". (Live Fund Value is $" + parseFloat(liveFundValue).toFixed(2) +")", ui.ButtonSet.YES_NO_CANCEL);
 
  //Cancel or Manual Entry of Fund Value
  if(fundValueResponse ==  ui.Button.CANCEL) return;
  else if(fundValueResponse ==  ui.Button.NO)
  {
      fundValue = getAndValidateNumber("Fund Value (USD)");
      if(fundValue == "cancel") return;
  }
  
  //Check Fund Value after transaction and ensure it's not negative
  var newFundValue = fundValue + transactionValue;
  
  if(newFundValue < 0) {
    ui.alert("ERROR: Fund value (USD) after withdrawal is " + newFundValue + ". Withdrawal not valid.");
    return;
  }
  

 
  //Confirm Transaction Details
  var response = ui.prompt("Please confirm details: \n\n" 
                           + "** " + transactionType + " ** \n\n" 
                           + "User ID: " + userID + "\n"
                           + "Coin: " + coin + "\n"
                           + "Quantity: " + Math.abs(quantity) + "\n" 
                           + "Close Price: $" + Math.abs(closePrice) + " USD\n" 
                           + transactionType + " Value: $" + Math.abs(transactionValue).toFixed(2) + " USD\n"
                           + "Fund Value (Before Transaction): $" + parseFloat(fundValue).toFixed(2) + " USD\n"
                           + "Fund Value (After Transaction): $" + parseFloat(newFundValue).toFixed(2) + " USD\n\n" 


                           + "If correct, type 'confirm' and click 'OK' to complete transaction.", ui.ButtonSet.OK_CANCEL);
  
  if(response.getSelectedButton() == "OK" && response.getResponseText() == "confirm") {   
    //Get GoogleID of user, timestamp and record in info
    var adminEmail = Session.getEffectiveUser().getEmail();
    
    
    //Add to Table
    var dwID = AddDepositWithdrawal(userID,coin,quantity,transactionType,closePrice, transactionValue, fundValue, newFundValue,adminEmail); 
    
    //Add Trade
    AddTradeDW(coin,quantity,dwID);
    
    //Update Equity
    updateEquitySplits(fundValue, userID, transactionValue);
    
    //Go through close process to record the current state of the fundtfa
    Close(transactionType, dwID);
    
    //ui.alert(transactionType + " Complete");
  }
  else ui.alert("Transaction Cancelled");

}

//Adds a 'trade' to the Transaction History that represents the deposit/withdrawal. 
//The 'trade' is only a sell order or buy order, based on whether it's a withdrawal or deposits
function AddTradeDW(coin,quantity, dwID){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = "Transaction History"
  var ui = SpreadsheetApp.getUi();

  //Adds a new line to the Trade History (increments the Trade ID, sets Date)
  var newTradeRow = AddNewTrade();
  
  var quantityOffset;
  var coinOffset;
  var dwIDOffset = 6;
  
  //If Withdrawal, Else Deposit
  if(quantity < 0){
    quantityOffset = 2;
    coinOffset = 3;
  }
  else {
    quantityOffset = 4;
    coinOffset = 5;
  }
  
  //Paste Trade Details
  var firstCellNewTrade = ss.getRangeByName("TH_TABLE").getCell(3, 1); 
  firstCellNewTrade.offset(0, quantityOffset).setValue(Math.abs(quantity));
  firstCellNewTrade.offset(0, coinOffset).setValue(coin);
  firstCellNewTrade.offset(0, dwIDOffset).setValue(dwID);  
}


//Updates the equity splits (percentages) on the Equity sheet based on a given deposit/withdrawal
function updateEquitySplits(currentFundValue, investorID, dwAmount){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ui = SpreadsheetApp.getUi();
  var numUsers = parseInt(ss.getRangeByName("MD_NUM_INVESTORS").getValue());
  var sheet = "Equity"
  var equityRange = ss.getRangeByName("E_TABLE");
  var newFundValue = currentFundValue + dwAmount; //Negative for withdrawal
  var user = [];
  var equitySplit = [];
  var investorIndex = -1; //The location in the array for the investor adding money now
  var userIDCol = ss.getRangeByName("E_USERID_COL").getColumn();;
  var equityCol = ss.getRangeByName("E_EQUITYSPLIT_COL").getColumn();
  
  //Get current users and equity split
  for(var i = 1; i <= numUsers; i++){   
    user[i] = equityRange.getCell(1+i, userIDCol).getValue();
    equitySplit[i] = parseFloat(equityRange.getCell(1+i,equityCol).getValue()); //Decimal Number representing percentage of fund owned
    
    if(user[i] == investorID) investorIndex = i; 
  }
  
  //InvestorID not found
  if(investorIndex == -1) return false;
 
  
  //Update Equity Split
  for(var i = 1; i <= numUsers; i++){ 
    if(i == investorIndex) 
      equitySplit[i] = ((equitySplit[i]*currentFundValue) + dwAmount) / newFundValue;
    else
      equitySplit[i] = equitySplit[i]*currentFundValue / newFundValue;
    
    //Save to Table
    equityRange.getCell(1+i,equityCol).setValue(equitySplit[i]);
    
  }             
  
  return true;
}

//Adds line to Deposit Withdrawal Table with all the details gathered in the DepositWithdrawal function.
function AddDepositWithdrawal(userID, coin, quantity, transactionType, closePrice, transactionValue, oldFundValue, newFundValue, adminEmail){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = "Deposit & Widthdrawal History"
  
  //Insert New Row For Transaction and Update Transaction Number and Date
  var lastTransCell = ss.getRangeByName("DW_TABLE").getCell(1, 1).offset(2, 0); //Offset 2 line down from header (header + start row)
  var lastTransNum = parseInt(lastTransCell.getValue());
  ss.insertRowBefore(lastTransCell.getRowIndex());
  var newTransCell = lastTransCell;
  var dwTransID = lastTransNum+1;
  lastTransCell.setValue(dwTransID);
  lastTransCell.offset(0,1).setValue(Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss"));
  
  //Enter all the details. 
  //** No error checking, since this is ONLY called from getDepositWithdrawalInfo which does error checking
  lastTransCell.offset(0,2).setValue(userID);
  lastTransCell.offset(0,3).setValue(transactionType);
  lastTransCell.offset(0,4).setValue(coin);
  lastTransCell.offset(0,5).setValue(quantity);
  lastTransCell.offset(0,6).setValue(closePrice);
  lastTransCell.offset(0,7).setValue(transactionValue);
  lastTransCell.offset(0,8).setValue(oldFundValue);
  lastTransCell.offset(0,9).setValue(newFundValue);
  lastTransCell.offset(0,10).setValue(adminEmail); 
  
  return dwTransID;
}




