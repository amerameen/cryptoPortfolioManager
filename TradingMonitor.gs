//Todo - Where to Trade, Where to Store for each coin
//Todo - does this work without refresh?

function checkLevels() {
 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var ui = SpreadsheetApp.getUi();
  //var sheet = ss.getSheetByName("Trading Monitor");
  var range = ss.getRangeByName("TM_TABLE");
  var table = ss.getRangeByName("TM_TABLE").getValues();
  var numRows = ss.getRangeByName("TM_TABLE").getNumRows();
  var cLvlCol = 1;
  var bLvlCol = 2;
  var bPctCol = 3;
  var bWarnCol = 4
  var sLvlCol = 5;
  var sPctCol = 6;
  var sWarnCol = 7;
  var baseCoin = "BTC";
  var baseCoinTotal = ss.getRangeByName("H_TOTAL_BTC").getValue();
  var coin, level, buyLevel, buyPct, sellLevel, sellPct;
  var warningThreshold = .02

  
  //Go through table and check each row to see if a notification needs to be sent
  //Starts from 1 to exclude header row and excludes the last (blank) row
  for(var r=1; r < numRows-1; r++)
  {
    //Check Current Row
    coin = table[r][0].toUpperCase();
    level = parseFloat(table[r][cLvlCol]);
    buyLevel = parseFloat(table[r][bLvlCol]);
    buyPct = parseFloat(table[r][bPctCol]);
    sellLevel = parseFloat(table[r][sLvlCol]);
    sellPct = parseFloat(table[r][sPctCol]);
    
    //Check Buy Level
    if(buyLevel != "" && buyPct != "")
    {
      if(level <= buyLevel) 
      {
        var baseCoinAmt = parseFloat(baseCoinTotal * buyPct).toFixed(3);
        var comment = "";
        var warning = "";
      
        if(getAllocation(coin) == 0)
          comment = "Position Opened";
        
        //Check if we have enough BTC
        var baseCoinBalance = checkBalance(baseCoin);
        if(baseCoinAmt > baseCoinBalance) warning = "Insufficient " + baseCoin + ". " + baseCoin  + " Balance = " + parseFloat(baseCoinBalance).toFixed(3);
        
        //Add To Notification History
        addNotToHist(coin,"BUY", baseCoinAmt, level, buyLevel, buyPct, comment, warning);
        
        //Clear Notification Level from Table
        range.offset(r, bLvlCol, 1, 3).clear();
     
        //Send Email
        sendEmail(coin,"BUY", baseCoinAmt, level, buyLevel, buyPct, comment, warning);
      }
      
      //Buy Warnings
      else if(1-buyLevel/level < warningThreshold && table[r][bWarnCol] == "" )
      { 
        var baseCoinAmt = parseFloat(baseCoinTotal * buyPct).toFixed(3);
        var comment = "";
        var warning = "";

        //Check if we have enough BTC
        var baseCoinBalance = checkBalance(baseCoin);
        if(baseCoinAmt > baseCoinBalance) 
           warning = "Insufficient " + baseCoin + ". " + baseCoin  + " Balance = " + parseFloat(baseCoinBalance).toFixed(3);
        
        //Add To Notification History
        addNotToHist(coin,"BUY SOON", baseCoinAmt, level, buyLevel, buyPct, comment, warning);
        
        //Mark Warning as Sent
        range.getCell(r+1, bWarnCol+1).setValue("Y");
        
        //Send Email
        sendEmail(coin,"BUY SOON", baseCoinAmt, level, buyLevel, buyPct, comment, warning);
      }
    }
    
    //Check Sell Level
    if(sellLevel != "" && sellPct != "")
    {
      if(level >= sellLevel)
      {
        var currentAllocation = getAllocation(coin);
        var comment = "";
        var warning = ""

        //If Sell Pct is more than we have, sell ALL, otherwise sell appropriate % of total
        if(currentAllocation < sellPct){
          sellPct = currentAllocation;
          comment = "Position Closed";
        }
          
          
        var tradeCoinAmt = parseFloat(checkBalance(coin) * sellPct/currentAllocation).toFixed(3);
        
        //Add To Notification History
        addNotToHist(coin,"SELL", tradeCoinAmt, level, sellLevel, sellPct, comment, warning);
        
        //Clear Notification Level from Table
        range.offset(r, sLvlCol, 1, 3).clear();
     
        //Send Email
        sendEmail(coin,"SELL", tradeCoinAmt, level, sellLevel, sellPct, comment, warning);
      }
      
      //Sell Warnings
      else if(1-level/sellLevel < warningThreshold && table[r][sWarnCol] == "" )
      {
        var currentAllocation = getAllocation(coin);
        var warning = "";

        //If Sell Pct is more than we have, sell ALL, otherwise sell appropriate % of total
        if(currentAllocation < sellPct)
          sellPct = currentAllocation;
          
        var tradeCoinAmt = parseFloat(checkBalance(coin) * sellPct/currentAllocation).toFixed(3);

        //Add To Notification History
        addNotToHist(coin,"SELL SOON", tradeCoinAmt, level, sellLevel, sellPct, "", warning);
        
        //Mark Warning as Sent
        range.getCell(r+1, sWarnCol+1).setValue("Y");
        
        //Send Email
        sendEmail(coin,"SELL SOON", tradeCoinAmt, level, sellLevel, sellPct, "", warning);

      }
    }
  }
}


//Informs us if we hold a coin, but don't have a sell level for it
//Compares Holdings to the tickers that have sell levels set
function checkForMissingSellLevels()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var tickerUC;
  var tickerLC;
  
  var tickerCol = 0;
  var sellLevelCol = 5;
  var tickers =  ss.getRangeByName("H_TICKERS").getValues();
  var numCoins = ss.getRangeByName("H_TICKERS").getNumRows();
  var monitorTable =  ss.getRangeByName("TM_TABLE").getValues();
  var monitorTableNumRows =  ss.getRangeByName("TM_TABLE").getNumRows();
  var monitorTickers = [];
  var missingSellLevels = []
  
  //Sell Level Tickers, Ignore Last (Empty) Row 
  for(var i = 1; i < monitorTableNumRows-1; i++)
  {
    if(monitorTable[i][sellLevelCol] != ""){
      var currentTicker = monitorTable[i][tickerCol]; 
      
      if(monitorTickers.indexOf(currentTicker) == -1) 
        monitorTickers.push(currentTicker);
     
    }
  }
  
  //Go through the coins on holdings sheet and compare to tickers with sell levels
  //Ignore first two (CAD/USD/BTC) and last (blank row)
  for(var i = 3; i < numCoins-1; i++){
    tickerUC = tickers[i][0];
    tickerLC = tickerUC.toLowerCase();
    
    if(monitorTickers.indexOf(tickerUC) == -1 && monitorTickers.indexOf(tickerLC) == -1){
      if(checkBalance(tickerUC) != 0)
        missingSellLevels.push(tickerUC);
    }
  }
  
  //If necessary, print to sheet and send email
  var numMissing = missingSellLevels.length;
  if(numMissing != 0){
    var missingString = "";
    for(var i = 0; i < numMissing; i++){
      missingString += missingSellLevels[i] + " ";
    }
    ss.getRangeByName("TM_MISSING_SELL_LEVELS").setValue(missingString);
  }
  else
    ss.getRangeByName("TM_MISSING_SELL_LEVELS").setValue("");
  
  return;
}

function sendEmail(coin, action, amount, level, notLevel, pct, comment, warning)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var emails = PropertiesService.getScriptProperties().getProperty("TradingMonitorEmails")
  var baseCoin = "BTC";
  var amtCoin;
  var commentLine = "";
  var warningLine = "";
  var subjectLine = "Alert: " + action + " " + coin;

  
  //Coin to Buy/Sell
  if(action == "BUY" || action == "BUY SOON") amtCoin = baseCoin;
  else amtCoin = coin;
  
  //Lines only displayed if there is content
  if(comment != "")
    commentLine = "<br><b> Comment: </b>" + comment;
  if(warning != ""){
    warningLine = "<br><b> Warning: </b>" + warning;
    subjectLine = "Alert + Warning: " + action + " " + coin;
  } 
  
  MailApp.sendEmail({
        to: emails,
        subject: subjectLine,
        htmlBody:     "<b> Action: </b>" + action + " " + coin
                + "<br><b> Amount: </b>" + amount + " " + amtCoin + " ("  + pct*100 + "%)"
                + "<br><b> Current Level: </b>" + level
                + "<br><b> Notification Level: </b>" + notLevel
                + commentLine
                + warningLine
                + "<br><br> Sent via <a href=\"https://docs.google.com/spreadsheets/d/1lvafocby-uv-N5P3efAGwJXDGqjWoM_8aBRxA6CBYic/edit#gid=0\">CryptoTrader</a>"                                   
      });  
}

function addNotToHist(coin, action, amount, level, notLevel, pct, comment, warning)
{
  const sheetName = "Trading Monitor Notification History"
  const namedRangeHeader = "TMNH_HEADER"
  var baseCoin = "BTC";
  var amtCoin;
  
  //Get Spreadsheet
  var ss = SpreadsheetApp.getActive();
  
  //Get Sheet & Find Line to Enter Trade On
  var sheet = ss.getSheetByName(sheetName);
  var lastNotCell = ss.getRangeByName(namedRangeHeader).getCell(1, 1).offset(1, 0); //Offset 2 lines (one header + one hidden 'start' row)
  
  //Coin to Buy/Sell
  if(action == "BUY" || action == "BUY SOON") amtCoin = baseCoin;
  else amtCoin = coin;
  
  //Insert New Row For Trade and Update Trade Number and Date
  sheet.insertRowBefore(lastNotCell.getRowIndex());
  lastNotCell.offset(0,0).setValue(Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss"));
  lastNotCell.offset(0,1).setValue(coin);
  lastNotCell.offset(0,2).setValue(action);
  lastNotCell.offset(0,3).setValue(amount + " " + amtCoin);
  lastNotCell.offset(0,4).setValue(level);
  lastNotCell.offset(0,5).setValue(notLevel);
  lastNotCell.offset(0,6).setValue(pct);
  lastNotCell.offset(0,7).setValue(comment);
  lastNotCell.offset(0,8).setValue(warning);
  
  return;
}

//This function adds a new line to the Trade History and increments the Trade ID
function AddNewMonitorRow() {
  
  const sheetName = "Trading Monitor"
  const namedRangeTradeHeader = "TM_TABLE"
  
  //Get Spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  
  //Get Sheet & Find Line to Enter Trade On
  var sheet = spreadsheet.getSheetByName(sheetName);
  var firstRowCell = spreadsheet.getRangeByName(namedRangeTradeHeader).getCell(1, 1).offset(1, 0); 
 
  
  //Insert New Row For Trade and Update Trade Number and Date
  sheet.insertRowBefore(firstRowCell.getRowIndex());
  firstRowCell.offset(1, 1).copyTo(firstRowCell.offset(0, 1).getCell(1, 1));
  var row = firstRowCell.getColumnIndex();
  return ;
}
