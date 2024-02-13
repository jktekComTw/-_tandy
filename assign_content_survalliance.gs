//to gen the tempSheet copy from the sheet template assigned
function copySheet4tempPasteUse(sourceSheet,tempName){
  var newSheet=sourceSheet.copyTo(gspreadsheet);
  newSheet.setName(tempName);
  return newSheet;
}

function gen_surv_Tempreport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var wildcard = "temp施工*"; // Replace "YourWildcardHere" with your wildcard pattern
  var matchingSheets = [];
  
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().match(wildcard)) {
      matchingSheets.push(sheets[i]);
    }
  }

  var i=0;
  if (matchingSheets.length > 0) {
    matchingSheets.forEach(function(sheet_diary){
      console.log(sheet_diary.getRange('C2').getValue());
      var tempSheet=copySheet4tempPasteUse(gspreadsheet.getSheetByName('公共工程監造報表'),'temp'+'監造'+(i+1));
      tempSheet.getRange('C2').setValue(sheet_diary.getRange('C2').getValue());
      tempSheet.getRange('D3').setValue(sheet_diary.getRange('D3').getValue());
      tempSheet.getRange('F3').setValue(sheet_diary.getRange('G3').getValue());

      tempSheet.getRange('J3').setValue(sheet_diary.getRange('P3').getValue());
      tempSheet.getRange('K3').setValue(sheet_diary.getRange('S3:T3').getValue());

      tempSheet.getRange('B7').setValue(sheet_diary.getRange('H6').getValue());
      tempSheet.getRange('I7').setValue(sheet_diary.getRange('M6').getValue());
      tempSheet.getRange('K7').setValue(sheet_diary.getRange('R6').getValue());

      tempSheet.getRange('B11:E11').setValue(sheet_diary.getRange('A12').getValue());
      tempSheet.getRange('B12:E12').setValue(sheet_diary.getRange('A13').getValue());
      tempSheet.getRange('B13:E13').setValue(sheet_diary.getRange('A15').getValue());
      tempSheet.getRange('B14:E14').setValue(sheet_diary.getRange('A17').getValue());
      tempSheet.getRange('B15:E15').setValue(sheet_diary.getRange('A19').getValue());
      tempSheet.getRange('B16:E16').setValue(sheet_diary.getRange('A23').getValue());
      tempSheet.getRange('B17:E17').setValue(sheet_diary.getRange('A26').getValue());
      tempSheet.getRange('B18:E18').setValue(sheet_diary.getRange('A30').getValue());
      tempSheet.getRange('B19:E19').setValue(sheet_diary.getRange('A34').getValue());
      tempSheet.getRange('B20:E20').setValue(sheet_diary.getRange('K19').getValue());
      tempSheet.getRange('B21:E21').setValue(sheet_diary.getRange('K21').getValue());
      tempSheet.getRange('B22:E22').setValue(sheet_diary.getRange('K23').getValue());
      tempSheet.getRange('B23:E23').setValue(sheet_diary.getRange('K25').getValue());
      tempSheet.getRange('B24:E24').setValue(sheet_diary.getRange('K27').getValue());
      tempSheet.getRange('B25:E25').setValue(sheet_diary.getRange('K29').getValue());

      tempSheet.getRange('J11').setValue(sheet_diary.getRange('G12').getValue());
      tempSheet.getRange('J12').setValue(sheet_diary.getRange('G13').getValue());
      tempSheet.getRange('J13').setValue(sheet_diary.getRange('G15').getValue());
      tempSheet.getRange('J14').setValue(sheet_diary.getRange('G17').getValue());
      tempSheet.getRange('J15').setValue(sheet_diary.getRange('G19').getValue());
      tempSheet.getRange('J16').setValue(sheet_diary.getRange('G23').getValue());
      tempSheet.getRange('J17').setValue(sheet_diary.getRange('G26').getValue());
      tempSheet.getRange('J18').setValue(sheet_diary.getRange('G30').getValue());
      tempSheet.getRange('J19').setValue(sheet_diary.getRange('G34').getValue());
      tempSheet.getRange('J20').setValue(sheet_diary.getRange('Q19').getValue());
      tempSheet.getRange('J21').setValue(sheet_diary.getRange('Q21').getValue());
      tempSheet.getRange('J22').setValue(sheet_diary.getRange('Q23').getValue());
      tempSheet.getRange('J23').setValue(sheet_diary.getRange('Q25').getValue());
      tempSheet.getRange('J24').setValue(sheet_diary.getRange('Q27').getValue());
      tempSheet.getRange('J25').setValue(sheet_diary.getRange('Q29').getValue());

      tempSheet.getRange('K11').setValue(sheet_diary.getRange('D12').getValue());
      tempSheet.getRange('k12').setValue(sheet_diary.getRange('D13').getValue());
      tempSheet.getRange('K13').setValue(sheet_diary.getRange('D15').getValue());
      tempSheet.getRange('K14').setValue(sheet_diary.getRange('D17').getValue());
      tempSheet.getRange('K15').setValue(sheet_diary.getRange('D19').getValue());
      tempSheet.getRange('K16').setValue(sheet_diary.getRange('D23').getValue());
      tempSheet.getRange('K17').setValue(sheet_diary.getRange('D26').getValue());
      tempSheet.getRange('K18').setValue(sheet_diary.getRange('D30').getValue());
      tempSheet.getRange('K19').setValue(sheet_diary.getRange('D34').getValue());
      tempSheet.getRange('K20').setValue(sheet_diary.getRange('N19').getValue());
      tempSheet.getRange('K21').setValue(sheet_diary.getRange('N21').getValue());
      tempSheet.getRange('K22').setValue(sheet_diary.getRange('N23').getValue());
      tempSheet.getRange('K23').setValue(sheet_diary.getRange('N25').getValue());
      tempSheet.getRange('K24').setValue(sheet_diary.getRange('N27').getValue());
      tempSheet.getRange('K25').setValue(sheet_diary.getRange('N29').getValue());

      tempSheet.getRange('A44:B44').setValue(sheet_diary.getRange('A56').getValue());

      i++;
    });
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error", "沒有temp施工日誌的資料表可供參考.", ui.ButtonSet.OK);
  }
}

