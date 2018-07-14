function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('CryptoTrader')
      .addItem('Refresh Prices', 'RefreshDataFeed')
      .addSeparator()
      .addItem('Add Trade', 'AddNewTrade')
      .addSeparator()
      .addItem('Add Coin To Holdings', 'AddCoin')
      .addSeparator()
      .addItem('Deposit or Withdrawal', 'DepositWithdrawal')
      .addToUi();
}
