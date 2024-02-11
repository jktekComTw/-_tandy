function gen_detail_TempReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var wildcard = "temp施工*"; // Replace "YourWildcardHere" with your wildcard pattern
  var matchingSheets = [];
  
  for (var i = 0; i < sheets.length; i++) {
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
      tempSheet.getRange('G3').setValue(sheet_diary.getRange('P3:R3').getValue());
      tempSheet.getRange('H3').setValue(sheet_diary.getRange('S3:T3').getValue());

      //items name
      //3 lines
      tempSheet.getRange('A8:D8').setValue(sheet_diary.getRange('A12').getValue());
      tempSheet.getRange('A9:D9').setValue(sheet_diary.getRange('A13').getValue());
      tempSheet.getRange('A10:D10').setValue(sheet_diary.getRange('A15').getValue());
      //11 lines
      tempSheet.getRange('A13:D13').setValue(sheet_diary.getRange('A17').getValue());
      tempSheet.getRange('A14:D14').setValue(sheet_diary.getRange('A19').getValue());
      tempSheet.getRange('A15:D15').setValue(sheet_diary.getRange('A21').getValue());
      tempSheet.getRange('A16:D16').setValue(sheet_diary.getRange('A23').getValue());
      tempSheet.getRange('A17:D17').setValue(sheet_diary.getRange('A24').getValue());
      tempSheet.getRange('A18:D18').setValue(sheet_diary.getRange('A26').getValue());
      tempSheet.getRange('A19:D19').setValue(sheet_diary.getRange('A28').getValue());
      tempSheet.getRange('A20:D20').setValue(sheet_diary.getRange('A30').getValue());
      tempSheet.getRange('A21:D21').setValue(sheet_diary.getRange('A32').getValue());
      tempSheet.getRange('A22:D22').setValue(sheet_diary.getRange('A33').getValue());
      tempSheet.getRange('A23:D23').setValue(sheet_diary.getRange('A34').getValue());
      //14 lines
      tempSheet.getRange('A26:D26').setValue(sheet_diary.getRange('K12').getValue());
      tempSheet.getRange('A27:D27').setValue(sheet_diary.getRange('K13').getValue());
      tempSheet.getRange('A28:D28').setValue(sheet_diary.getRange('K14').getValue());
      tempSheet.getRange('A29:D29').setValue(sheet_diary.getRange('K15').getValue());
      tempSheet.getRange('A30:D30').setValue(sheet_diary.getRange('K16').getValue());
      tempSheet.getRange('A31:D31').setValue(sheet_diary.getRange('K17').getValue());
      tempSheet.getRange('A32:D32').setValue(sheet_diary.getRange('K18').getValue());
      tempSheet.getRange('A33:D33').setValue(sheet_diary.getRange('K19').getValue());
      tempSheet.getRange('A34:D34').setValue(sheet_diary.getRange('K21').getValue());
      tempSheet.getRange('A35:D35').setValue(sheet_diary.getRange('K23').getValue());
      tempSheet.getRange('A36:D36').setValue(sheet_diary.getRange('K25').getValue());
      tempSheet.getRange('A37:D37').setValue(sheet_diary.getRange('K27').getValue());
      tempSheet.getRange('A38:D38').setValue(sheet_diary.getRange('K29').getValue());
      tempSheet.getRange('A39:D39').setValue(sheet_diary.getRange('K30').getValue());

      //2 lines
      tempSheet.getRange('A42:D42').setValue(sheet_diary.getRange('K32').getValue());
      tempSheet.getRange('A45:D45').setValue(sheet_diary.getRange('K34').getValue());
      //unit
      //3 lines
      tempSheet.getRange('E8').setValue(sheet_diary.getRange('D12').getValue());
      tempSheet.getRange('E9').setValue(sheet_diary.getRange('D13').getValue());
      tempSheet.getRange('E10').setValue(sheet_diary.getRange('D15').getValue());
      //11 lines
      tempSheet.getRange('E13').setValue(sheet_diary.getRange('D17').getValue());
      tempSheet.getRange('E14').setValue(sheet_diary.getRange('D19').getValue());
      tempSheet.getRange('E15').setValue(sheet_diary.getRange('D21').getValue());
      tempSheet.getRange('E16').setValue(sheet_diary.getRange('D23').getValue());
      tempSheet.getRange('E17').setValue(sheet_diary.getRange('D24').getValue());
      tempSheet.getRange('E18').setValue(sheet_diary.getRange('D26').getValue());
      tempSheet.getRange('E19').setValue(sheet_diary.getRange('D28').getValue());
      tempSheet.getRange('E20').setValue(sheet_diary.getRange('D30').getValue());
      tempSheet.getRange('E21').setValue(sheet_diary.getRange('D32').getValue());
      tempSheet.getRange('E22').setValue(sheet_diary.getRange('D33').getValue());
      tempSheet.getRange('E23').setValue(sheet_diary.getRange('D34').getValue());
      //14 lines
      tempSheet.getRange('E26').setValue(sheet_diary.getRange('N12').getValue());
      tempSheet.getRange('E27').setValue(sheet_diary.getRange('N13').getValue());
      tempSheet.getRange('E28').setValue(sheet_diary.getRange('N14').getValue());
      tempSheet.getRange('E29').setValue(sheet_diary.getRange('N15').getValue());
      tempSheet.getRange('E30').setValue(sheet_diary.getRange('N16').getValue());
      tempSheet.getRange('E31').setValue(sheet_diary.getRange('N17').getValue());
      tempSheet.getRange('E32').setValue(sheet_diary.getRange('N18').getValue());
      tempSheet.getRange('E33').setValue(sheet_diary.getRange('N19').getValue());
      tempSheet.getRange('E34').setValue(sheet_diary.getRange('N21').getValue());
      tempSheet.getRange('E35').setValue(sheet_diary.getRange('N23').getValue());
      tempSheet.getRange('E36').setValue(sheet_diary.getRange('N25').getValue());
      tempSheet.getRange('E37').setValue(sheet_diary.getRange('N27').getValue());
      tempSheet.getRange('E38').setValue(sheet_diary.getRange('N29').getValue());
      tempSheet.getRange('E39').setValue(sheet_diary.getRange('N30').getValue());

      //2 lines
      tempSheet.getRange('E42').setValue(sheet_diary.getRange('N32').getValue());
      tempSheet.getRange('E45').setValue(sheet_diary.getRange('N34').getValue());

      //contract amount
      //3 lines
      tempSheet.getRange('F8').setValue(sheet_diary.getRange('E12').getValue());
      tempSheet.getRange('F9').setValue(sheet_diary.getRange('E13').getValue());
      tempSheet.getRange('F10').setValue(sheet_diary.getRange('E15').getValue());
      //11 lines
      tempSheet.getRange('F13').setValue(sheet_diary.getRange('E17').getValue());
      tempSheet.getRange('F14').setValue(sheet_diary.getRange('E19').getValue());
      tempSheet.getRange('F15').setValue(sheet_diary.getRange('E21').getValue());
      tempSheet.getRange('F16').setValue(sheet_diary.getRange('E23').getValue());
      tempSheet.getRange('F17').setValue(sheet_diary.getRange('E24').getValue());
      tempSheet.getRange('F18').setValue(sheet_diary.getRange('E26').getValue());
      tempSheet.getRange('F19').setValue(sheet_diary.getRange('E28').getValue());
      tempSheet.getRange('F20').setValue(sheet_diary.getRange('E30').getValue());
      tempSheet.getRange('F21').setValue(sheet_diary.getRange('E32').getValue());
      tempSheet.getRange('F22').setValue(sheet_diary.getRange('E33').getValue());
      tempSheet.getRange('F23').setValue(sheet_diary.getRange('E34').getValue());
      //14 lines
      tempSheet.getRange('F26').setValue(sheet_diary.getRange('O12').getValue());
      tempSheet.getRange('F27').setValue(sheet_diary.getRange('O13').getValue());
      tempSheet.getRange('F28').setValue(sheet_diary.getRange('O14').getValue());
      tempSheet.getRange('F29').setValue(sheet_diary.getRange('O15').getValue());
      tempSheet.getRange('F30').setValue(sheet_diary.getRange('O16').getValue());
      tempSheet.getRange('F31').setValue(sheet_diary.getRange('O17').getValue());
      tempSheet.getRange('F32').setValue(sheet_diary.getRange('O18').getValue());
      tempSheet.getRange('F33').setValue(sheet_diary.getRange('O19').getValue());
      tempSheet.getRange('F34').setValue(sheet_diary.getRange('O21').getValue());
      tempSheet.getRange('F35').setValue(sheet_diary.getRange('O23').getValue());
      tempSheet.getRange('F36').setValue(sheet_diary.getRange('O25').getValue());
      tempSheet.getRange('F37').setValue(sheet_diary.getRange('O27').getValue());
      tempSheet.getRange('F38').setValue(sheet_diary.getRange('O29').getValue());
      tempSheet.getRange('F39').setValue(sheet_diary.getRange('O30').getValue());

      //today amount
      //3 lines
      tempSheet.getRange('G8').setValue(sheet_diary.getRange('G12').getValue());
      tempSheet.getRange('G9').setValue(sheet_diary.getRange('G13').getValue());
      tempSheet.getRange('G10').setValue(sheet_diary.getRange('G15').getValue());
      //11 lines
      tempSheet.getRange('G13').setValue(sheet_diary.getRange('G17').getValue());
      tempSheet.getRange('G14').setValue(sheet_diary.getRange('G19').getValue());
      tempSheet.getRange('G15').setValue(sheet_diary.getRange('G21').getValue());
      tempSheet.getRange('G16').setValue(sheet_diary.getRange('G23').getValue());
      tempSheet.getRange('G17').setValue(sheet_diary.getRange('G24').getValue());
      tempSheet.getRange('G18').setValue(sheet_diary.getRange('G26').getValue());
      tempSheet.getRange('G19').setValue(sheet_diary.getRange('G28').getValue());
      tempSheet.getRange('G20').setValue(sheet_diary.getRange('G30').getValue());
      tempSheet.getRange('G21').setValue(sheet_diary.getRange('G32').getValue());
      tempSheet.getRange('G22').setValue(sheet_diary.getRange('G33').getValue());
      tempSheet.getRange('G23').setValue(sheet_diary.getRange('G34').getValue());
      //14 lines
      tempSheet.getRange('G26').setValue(sheet_diary.getRange('Q12').getValue());
      tempSheet.getRange('G27').setValue(sheet_diary.getRange('Q13').getValue());
      tempSheet.getRange('G28').setValue(sheet_diary.getRange('Q14').getValue());
      tempSheet.getRange('G29').setValue(sheet_diary.getRange('Q15').getValue());
      tempSheet.getRange('G30').setValue(sheet_diary.getRange('Q16').getValue());
      tempSheet.getRange('G31').setValue(sheet_diary.getRange('Q17').getValue());
      tempSheet.getRange('G32').setValue(sheet_diary.getRange('Q18').getValue());
      tempSheet.getRange('G33').setValue(sheet_diary.getRange('Q19').getValue());
      tempSheet.getRange('G34').setValue(sheet_diary.getRange('Q21').getValue());
      tempSheet.getRange('G35').setValue(sheet_diary.getRange('Q23').getValue());
      tempSheet.getRange('G36').setValue(sheet_diary.getRange('Q25').getValue());
      tempSheet.getRange('G37').setValue(sheet_diary.getRange('Q27').getValue());
      tempSheet.getRange('G38').setValue(sheet_diary.getRange('Q29').getValue());
      tempSheet.getRange('G39').setValue(sheet_diary.getRange('Q30').getValue());

      //accumulation amount
      //3 lines
      tempSheet.getRange('H8').setValue(sheet_diary.getRange('I12').getValue());
      tempSheet.getRange('H9').setValue(sheet_diary.getRange('I13').getValue());
      tempSheet.getRange('H10').setValue(sheet_diary.getRange('I15').getValue());
      //11 lines
      tempSheet.getRange('H13').setValue(sheet_diary.getRange('I17').getValue());
      tempSheet.getRange('H14').setValue(sheet_diary.getRange('I19').getValue());
      tempSheet.getRange('H15').setValue(sheet_diary.getRange('I21').getValue());
      tempSheet.getRange('H16').setValue(sheet_diary.getRange('I23').getValue());
      tempSheet.getRange('H17').setValue(sheet_diary.getRange('I24').getValue());
      tempSheet.getRange('H18').setValue(sheet_diary.getRange('I26').getValue());
      tempSheet.getRange('H19').setValue(sheet_diary.getRange('I28').getValue());
      tempSheet.getRange('H20').setValue(sheet_diary.getRange('I30').getValue());
      tempSheet.getRange('H21').setValue(sheet_diary.getRange('I32').getValue());
      tempSheet.getRange('H22').setValue(sheet_diary.getRange('I33').getValue());
      tempSheet.getRange('H23').setValue(sheet_diary.getRange('I34').getValue());
      //14 lines
      tempSheet.getRange('H26').setValue(sheet_diary.getRange('S12').getValue());
      tempSheet.getRange('H27').setValue(sheet_diary.getRange('S13').getValue());
      tempSheet.getRange('H28').setValue(sheet_diary.getRange('S14').getValue());
      tempSheet.getRange('H29').setValue(sheet_diary.getRange('S15').getValue());
      tempSheet.getRange('H30').setValue(sheet_diary.getRange('S16').getValue());
      tempSheet.getRange('H31').setValue(sheet_diary.getRange('S17').getValue());
      tempSheet.getRange('H32').setValue(sheet_diary.getRange('S18').getValue());
      tempSheet.getRange('H33').setValue(sheet_diary.getRange('S19').getValue());
      tempSheet.getRange('H34').setValue(sheet_diary.getRange('S21').getValue());
      tempSheet.getRange('H35').setValue(sheet_diary.getRange('S23').getValue());
      tempSheet.getRange('H36').setValue(sheet_diary.getRange('S25').getValue());
      tempSheet.getRange('H37').setValue(sheet_diary.getRange('S27').getValue());
      tempSheet.getRange('H38').setValue(sheet_diary.getRange('S29').getValue());
      tempSheet.getRange('H39').setValue(sheet_diary.getRange('S30').getValue());



      i++;
    });
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Error", "沒有temp施工日誌的資料表可供參考.", ui.ButtonSet.OK);
  }
}

