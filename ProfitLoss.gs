//This performs the closing process which is used for daily recordkeeping and recordkeeping during deposits/withdrawals
//Records current prices, balances, total value, and equity to the table on the closes sheet
function WeeklyPL()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Weekly P&L");
  var weekNum = Utilities.formatDate(new Date(), "GMT", "w"); 

  //Close 
  Close("Weekly P&L", weekNum);
  
  //Create New Row for the close and add date
  var lastPLCell = ss.getRangeByName("WPL_LIVE_CELL").getCell(1, 1).offset(1, 0); 
  sheet.insertRowBefore(lastPLCell.getRowIndex());
  
  //Copy Old Values to Empty Row
  range = ss.getRangeByName("WPL_LIVE_VALUES").offset(2, 0)
  range.copyTo(range.offset(-1, 0), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  SpreadsheetApp.flush();
  
  //Copy P&L
  range = ss.getRangeByName("WPL_LIVE_VALUES");  
  range.copyTo(range.offset(1, 0), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  //Week Number
  lastPLCell.setValue(weekNum);
  
  //Send Email
  weeklyPLEmail();
}

function weeklyPLEmail(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var emails = PropertiesService.getScriptProperties().getProperty("emails")
  var emailFooter = PropertiesService.getScriptProperties().getProperty("emailFooter")
  var weekNum = Utilities.formatDate(new Date(), "GMT", "w"); 

  
  MailApp.sendEmail({
        to: emails,
        subject: "CryptoTrader: Weekly P&L Complete (Week " + weekNum + ")",
        htmlBody: "Please check P&L values and rerun if necessary" + emailFooter                                   
      });  
}