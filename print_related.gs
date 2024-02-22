function printWorkTemp2pdf(){
  hideSheet();
  print2pdf();
  
}

gArray=['施工日誌','公共工程監造報表','施工明細表','no_acumulation','with_acumulation','上期累計'];

function hideSheet(){
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for(var i=0;i<gArray.length;i++){
    let sheet = spreadsheet.getSheetByName(gArray[i]);
    sheet.hideSheet();
  }
  
}

function showSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for(var i=0;i<gArray.length;i++){
    let sheet = spreadsheet.getSheetByName(gArray[i]);
    sheet.showSheet();
  }
}


function print2pdf() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var folders = DriveApp.getFoldersByName("paintingHouse_data");
  while (folders.hasNext()) {
    var folder = folders.next();
  }
  exportSpreadsheetAsPdf(spreadsheetId,folder.getId());
}


function exportSpreadsheetAsPdf(spreadsheetId, folderId) {
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf&format=pdf" +
            "&size=A4" + // Or your desired size
            "&portrait=true" + // Or false for landscape
            "&scale=4" +
            "&sheetnames=false&printtitle=false&pagenumbers=false" + // Adjust as needed
            "&gridlines=false" + // Display gridlines
            "&fzr=false"; // Repeat frozen rows

  var options = {
    headers: {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };

  var now=new Date();
  var response = UrlFetchApp.fetch(url, options);
  var blob = response.getBlob().setName("MyRangesAsPDF_"+now.toLocaleTimeString()+".pdf");

  // Save the PDF to Google Drive, optionally in a specific folder
  var file;
  if (folderId) {
    var folder = DriveApp.getFolderById(folderId);
    file = folder.createFile(blob);
  } else {
    file = DriveApp.createFile(blob);
  }

  var sheets = gspreadsheet.getSheets();
  var wildcard = "temp*"; // Replace "YourWildcardHere" with your wildcard pattern
  var matchingSheets = [];
  
  //delete all printed sheet
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().match(wildcard)) {
      matchingSheets.push(sheets[i]);
    }
  }
  showSheet();
  matchingSheets.forEach((sheet)=>{
    gspreadsheet.deleteSheet(sheet);
  });


  return ContentService.createTextOutput(file.getBlob().getBytes())
    .setMimeType(ContentService.MimeType.PDF) // Or use the appropriate MIME type
    .downloadAsFile(file.getName());
}

