const D=4;

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('漆屋功能')
      .addItem('產生日期', 'showSidebar')
      .addItem('複製並累加', 'wrapCopyAndAcc4Grids')
      .addToUi();
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('漆屋列印報表')
      .addItem('列印施工日誌', 'ToPrintWorkDiary')
      // .addItem('列印監照報表', function(){
      //   console.log("not implemented yet");
      // })
      // .addItem('列印施工明細', function(){
      //   console.log("not implemented yet");
      // })
      .addToUi();
}




function showSidebar() {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("no_acumulation");
  sheet.activate();
  let range=sheet.getRange(1,D,1,100);
  range.setValue("");
  
  var html = HtmlService.createHtmlOutputFromFile('options')
      .setTitle('請選擇日期範圍');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}



function processForm(formObject){

  //close the sidebar
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showSidebar(html);

  //get the data of from from sidebar html file
  var sd = formObject.sd;
  var ed = formObject.ed;
  console.log('起始與結束日期:'+sd+"~"+ed);
  Logger.log('起始與結束日期:'+sd+"~"+ed);
  testListDatesBetween(sd,ed,D);  //4 means D
}



//gen each element and let they do their demand routine
function testListDatesBetween(sd,ed,D) {
  var startDate = sd;
  var endDate = ed;
  var dates = listDatesBetween(startDate, endDate);
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  dates.forEach(function(dt) {
    let day=dt.getDate();
    let mon=dt.getMonth();
    let year=dt.getFullYear();
    day=parseInt(day,10);
    mon=mon+1;
    // console.log(day);
    let range=sheet.getRange(1,(D-1)+day);
    let workdate=(year.toString()+'/'+mon.toString()+'/'+day.toString());
    range.setValue(workdate.toString());
    range.setNumberFormat('@');

  });
}




