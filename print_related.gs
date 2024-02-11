var gspreadsheet=SpreadsheetApp.getActiveSpreadsheet();
var sheet2Hide=['no_acumulation'];


//hide the sheet not 4 報表
function hideSpecificSheet(){
  sheet2Hide.forEach(function(sheetName){
    var sheet=gspreadsheet.getSheetByName(sheetName);
    if(sheet){
      sheet.hideSheet();
    }
  });
}

var sourceSheet=gspreadsheet.getSheetByName('with_acmulation');
var targetSheet=gspreadsheet.getSheetByName('施工日誌');

//parameter should be targetSheet.getName()
function genTempSheet4Print(targetSheetName){
  if(targetSheetName=='施工日誌'){
    // sheet2Hide.push('公共工程監造報表');
    // sheet2Hide.push('施工明細表');
    // hideSpecificSheet();
    

  }else if(targetSheetName=='公共工程監造報表'){
    console.log("not implement yet");
  }else if(targetSheetName=='施工明細表'){
    console.log("not implement yet");
  }else{
    console.log("error:no such sheet name!");
  }
}


//for 施工報表
function print_0(NAMEsheet) {
  targetSheet=gspreadsheet.getSheetByName(NAMEsheet);
  genTempSheet4Print(targetSheet.getName());

}





