// function logGridIds() {
//   var spreadsheetId = SpreadsheetApp.getActive().getSheetByName("no_acumulation").getIndex();
//   // var sheets = Sheets.Spreadsheets.get(spreadsheetId).sheets;
//   console.log(spreadsheetId);
//   // for (var i = 0; i < sheets.length; i++) {
//   //   var sheet = sheets[i];
//   //   Logger.log('Sheet name: ' + sheet.properties.title + ', Grid ID: ' + sheet.properties.sheetId);
//   // }
// }
// const D=4;
const EndCol=32;
const EndRow=38;
const StartRow=1;

//to add comma every 3 digits
function numFormat(num){
  return num.toLocaleString('en-US');
}

var sheet_no_acc = SpreadsheetApp.getActive().getSheetByName("no_acumulation");
var sheet_acc = SpreadsheetApp.getActive().getSheetByName("with_acumulation");

// sheet.activate();

function CopyAndAcc4Grids(sheet_org,sheet_tar){
  var preValue;

  //scan cols first, after finished, change row
  for(var j=StartRow;j<=EndRow;j++){  
    for(var i=D;i<EndCol;i++){  
      var grid_orig=sheet_org.getRange(j,i);
      var grid_target=sheet_tar.getRange(j,i);
      if(j!=StartRow){
        if(grid_orig.getValues()!=""&&i==D){ 
          if(!isNaN(parseInt(grid_orig.getValues(),10))&&!grid_orig.getValues().includes('/')){  
            grid_target.setValue(numFormat(grid_orig.getValues()));
            preValue=grid_orig.getValues();
            preValue=parseInt(preValue,10);
          }else{
            grid_target.setValue(grid_orig.getValues());
          }
        }else if(grid_orig.getValues()!=""&&i!=D){//column except D and has content
          if(!isNaN(parseInt(grid_orig.getValues(),10))&&!grid_orig.getValues().includes('/')){  //if is number, the accumulation will be executed
            grid_target.setValue(numFormat(parseInt(grid_orig.getValues(),10)+parseInt(preValue,10)));
            preValue=parseInt(grid_orig.getValues(),10)+parseInt(preValue,10);
          }else{
            grid_target.setValue(grid_orig.getValues());
          }        
          
        }
        if(i==D){
          preValue=parseInt(grid_orig.getValues(),10);
        }
      }else{  //for date
        grid_target.setNumberFormat('@');
        grid_target.setValue(grid_orig.getValues());
      }
    }
  }
}

function wrapCopyAndAcc4Grids(){
  CopyAndAcc4Grids(sheet_no_acc,sheet_acc,D);
} 





