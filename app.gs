const D=4;
const gspreadsheet = SpreadsheetApp.getActive();

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('漆屋報表初始化相關功能')
      .addItem('初始資料填寫', 'showSidebar')
      .addItem('複製並累加', 'wrapCopyAndAcc4Grids')
      .addToUi();
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('漆屋報表產生')
      .addItem('產生施工日誌', 'ToPrintWorkDiary')
      .addItem('產生監照報表', 'gen_surv_Tempreport')
      .addItem('產生施工明細', 'gen_detail_TempReport')
      .addToUi();
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('列印漆屋報表')
      .addItem('列印所有報表', 'printWorkTemp2pdf')
      .addToUi();
}





function showSidebar() {
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("with_acumulation");
  
  let range=sheet.getRange("D1:AH38");
  range.setValue("");

  var sheet2 = SpreadsheetApp.getActive().getSheetByName("no_acumulation");
  range=sheet2.getRange("D1:AH38");
  range.setValue("");
  sheet2.activate();
  
  var html = HtmlService.createHtmlOutputFromFile('options')
      .setTitle('請選擇日期範圍');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}



function processForm(formObject){
  var gstartDay=formObject.startDay;
  //close the sidebar
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showSidebar(html);

  //get the data of from from sidebar html file
  var sd = formObject.sd;
  var ed = formObject.ed;
  console.log('起始與結束日期:'+sd+"~"+ed);
  Logger.log('起始與結束日期:'+sd+"~"+ed);

  
  //console.log(startDay);
  let sheet=gspreadsheet.getSheetByName('施工日誌');
  sheet.getRange("H5:I5").setValue(gstartDay);
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




