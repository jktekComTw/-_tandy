function logGridIds() {
  var spreadsheetId = SpreadsheetApp.getActive().getSheetByName("no_acumulation").getIndex();
  // var sheets = Sheets.Spreadsheets.get(spreadsheetId).sheets;
  console.log(spreadsheetId);
  // for (var i = 0; i < sheets.length; i++) {
  //   var sheet = sheets[i];
  //   Logger.log('Sheet name: ' + sheet.properties.title + ', Grid ID: ' + sheet.properties.sheetId);
  // }
}