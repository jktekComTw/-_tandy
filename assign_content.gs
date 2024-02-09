function ToPrintWorkDiary(){
  mapContent4no_acc();
  mapContent4with_acc();
}

function showDayOfWeek(dateString) {
  var today = new Date(dateString); // 獲取當前日期
  var dayOfWeek = today.getDay(); // 獲取星期幾的數字表示
  console.log(dayOfWeek);
  var daysInChinese = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
  
  var chineseDayOfWeek = daysInChinese[dayOfWeek];
  
  return chineseDayOfWeek;
}




function mapContent4no_acc() {
  var src=gspreadsheet.getSheetByName('no_acumulation');
  var range_src = src.getDataRange();
  var values_src = range_src.getValues();
  var columns_src = [];
  var numRows_src = values_src.length;
  var numCols_src = values_src[0].length;
  for (var col = 0; col < numCols_src; col++) {
    columns_src.push([]);
  }

  for (var row = 0; row < numRows_src; row++) {
    for (var col = D-1; col < numCols_src; col++) {
      columns_src[col].push(values_src[row][col]);
    }
  }
  
  for(var i=D-1;i<numCols_src;i++){
    
    var tempSheet=copySheet4tempPasteUse(gspreadsheet.getSheetByName('施工日誌'),'temp'+'施工'+(i-(D-1)));
    tempSheet.getRange("S3:T3").setValue(showDayOfWeek(columns_src[i][0]));  //show day of week
    
    let dtt=columns_src[i][0].toString();
    let dt=new Date(dtt);
    let day=dt.getDate();
    let mon=dt.getMonth();
    let year=dt.getFullYear()-1911;
    day=parseInt(day,10);
    mon=mon+1;
    let taiwan_format=('民國'+year.toString()+'年'+mon.toString()+'月'+day.toString()+'日');
    tempSheet.getRange("P3:R3").setValue(taiwan_format);  //show date in taiwan format
    tempSheet.getRange("D3").setValue(columns_src[i][1]);
    tempSheet.getRange("G3").setValue(columns_src[i][2]);
    
    tempSheet.getRange("G12").setValue(columns_src[i][5-1]);
    tempSheet.getRange("G13:H14").setValue(columns_src[i][7-1]);
    tempSheet.getRange("G15:H15").setValue(columns_src[i][8-1]);
    tempSheet.getRange("G17:H18").setValue(columns_src[i][10-1]);
    tempSheet.getRange("G19:H20").setValue(columns_src[i][11-1]);
    tempSheet.getRange("G21:H22").setValue(columns_src[i][12-1]);
    tempSheet.getRange("G23:H23").setValue(columns_src[i][13-1]);
    tempSheet.getRange("G24:H25").setValue(columns_src[i][14-1]);
    tempSheet.getRange("G26:H27").setValue(columns_src[i][15-1]);
    tempSheet.getRange("G28:H29").setValue(columns_src[i][16-1]);
    tempSheet.getRange("G30:H31").setValue(columns_src[i][17-1]);
    tempSheet.getRange("G32:H32").setValue(columns_src[i][18-1]);
    tempSheet.getRange("G34:H34").setValue(columns_src[i][19-1]);
    tempSheet.getRange("Q12:R12").setValue(columns_src[i][21-1]);
    tempSheet.getRange("Q13:R13").setValue(columns_src[i][22-1]);
    tempSheet.getRange("Q14:R14").setValue(columns_src[i][23-1]);
    tempSheet.getRange("Q15:R15").setValue(columns_src[i][24-1]);
    tempSheet.getRange("Q16:R16").setValue(columns_src[i][25-1]);
    tempSheet.getRange("Q17:R17").setValue(columns_src[i][26-1]);
    tempSheet.getRange("Q18:R18").setValue(columns_src[i][27-1]);
    tempSheet.getRange("Q19:R20").setValue(columns_src[i][28-1]);
    tempSheet.getRange("Q21:R22").setValue(columns_src[i][29-1]);
    tempSheet.getRange("Q23:R24").setValue(columns_src[i][30-1]);
    tempSheet.getRange("Q25:R26").setValue(columns_src[i][31-1]);
    tempSheet.getRange("Q27:R28").setValue(columns_src[i][32-1]);
    tempSheet.getRange("Q29:R29").setValue(columns_src[i][33-1]);
    tempSheet.getRange("Q30:R30").setValue(columns_src[i][34-1]);
    tempSheet.getRange("Q32:R32").setValue(columns_src[i][36-1]);
    tempSheet.getRange("Q34:R34").setValue(columns_src[i][38-1]);
  }
}

function mapContent4with_acc() {
  var src=gspreadsheet.getSheetByName('with_acumulation');
  var range_src = src.getDataRange();
  var values_src = range_src.getValues();
  var columns_src = [];
  var numRows_src = values_src.length;
  var numCols_src = values_src[0].length;
  // console.log(numCols_src);
  for (var col = 0; col < numCols_src; col++) {
    columns_src.push([]);
  }

  for (var row = 0; row < numRows_src; row++) {
    for (var col = D-1; col < numCols_src; col++) {
      columns_src[col].push(values_src[row][col]);
    }
  }
  for(var i=D-1;i<numCols_src;i++){
    // console.log(columns_src[i]);
    // continue;
    
    var tempSheet=gspreadsheet.getSheetByName('temp'+'施工'+(i-(D-1)));
    
    tempSheet.getRange("I12:J12").setValue(columns_src[i][6-1]);
    tempSheet.getRange("I13:J14").setValue(columns_src[i][7-1]);
    tempSheet.getRange("I15:J15").setValue(columns_src[i][8-1]);
    tempSheet.getRange("I17:J18").setValue(columns_src[i][10-1]);
    tempSheet.getRange("I19:J20").setValue(columns_src[i][11-1]);
    tempSheet.getRange("I21:J22").setValue(columns_src[i][12-1]);
    tempSheet.getRange("I23:J23").setValue(columns_src[i][13-1]);
    tempSheet.getRange("I24:J25").setValue(columns_src[i][14-1]);
    tempSheet.getRange("I26:J27").setValue(columns_src[i][15-1]);
    tempSheet.getRange("I28:J29").setValue(columns_src[i][16-1]);
    tempSheet.getRange("I30:J31").setValue(columns_src[i][17-1]);
    tempSheet.getRange("I32:J32").setValue(columns_src[i][18-1]);
    tempSheet.getRange("I34:J34").setValue(columns_src[i][19-1]);
    tempSheet.getRange("S12:T12").setValue(columns_src[i][21-1]);
    tempSheet.getRange("S13:T13").setValue(columns_src[i][22-1]);
    tempSheet.getRange("S14:T14").setValue(columns_src[i][23-1]);
    tempSheet.getRange("S15:T15").setValue(columns_src[i][24-1]);
    tempSheet.getRange("S16:T16").setValue(columns_src[i][25-1]);
    tempSheet.getRange("S17:T17").setValue(columns_src[i][26-1]);
    tempSheet.getRange("S18:T18").setValue(columns_src[i][27-1]);
    tempSheet.getRange("S19:T20").setValue(columns_src[i][28-1]);
    tempSheet.getRange("S21:T22").setValue(columns_src[i][29-1]);
    tempSheet.getRange("S23:T24").setValue(columns_src[i][30-1]);
    tempSheet.getRange("S25:T26").setValue(columns_src[i][31-1]);
    tempSheet.getRange("S27:T28").setValue(columns_src[i][32-1]);
    tempSheet.getRange("S29:T29").setValue(columns_src[i][33-1]);
    tempSheet.getRange("S30:T30").setValue(columns_src[i][34-1]);
    tempSheet.getRange("S32:T32").setValue(columns_src[i][36-1]);
    tempSheet.getRange("S34:T34").setValue(columns_src[i][38-1]);
  }
}




//to gen the tempSheet copy from the sheet template assigned
function copySheet4tempPasteUse(sourceSheet,tempName){
  var newSheet=sourceSheet.copyTo(gspreadsheet);
  newSheet.setName(tempName);
  return newSheet;
}






