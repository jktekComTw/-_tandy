const EndCol=40;
const EndRow=38;
const StartRow=1;

//to add comma every 3 digits
function numFormat(num){
  // return num;
  return num.toLocaleString('en-US');
}

var sheet_no_acc = SpreadsheetApp.getActive().getSheetByName("no_acumulation");
var sheet_acc = SpreadsheetApp.getActive().getSheetByName("with_acumulation");
var sheetlastMonth = SpreadsheetApp.getActive().getSheetByName("上期累計");



function CopyAndAcc4Grids(sheet_org,sheet_tar){
  var preValue;
  sheet_acc.getRange("D1:AI38").setValue("");
  //scan cols first, after finished, change row
  for(var j=StartRow;j<=EndRow;j++){  
    for(var i=D;i<EndCol;i++){  
      var grid_orig=sheet_org.getRange(j,i);
      var grid_target=sheet_tar.getRange(j,i);
      if(j!=StartRow){
        if(grid_orig.getValues()!=""&&i==D){//for D column 
          if(!isNaN(parseInt(grid_orig.getValues(),10))&&!grid_orig.getValues().includes('/')){ 
            grid_target.setValue(numFormat(parseFloat(grid_orig.getValues(),10)+
            (sheetlastMonth.getRange(j-D,1).getValue()==""?"":parseFloat(sheetlastMonth.getRange(j-D,1).getValue(),10))
            ));
            preValue=parseFloat(grid_target.getValues(),10);
          }else{
            grid_target.setValue(grid_orig.getValues());//for range NaN
          }
        }else if(grid_orig.getValues()!=""&&i!=D){//column except D and has content
          if(!isNaN(parseInt(grid_orig.getValues(),10))&&!grid_orig.getValues().includes('/')){  //if is number, the accumulation will be executed
            grid_target.setValue(numFormat(parseFloat(grid_orig.getValues(),10)+parseFloat(preValue,10)));
            
            preValue=parseFloat(grid_orig.getValues(),10)+parseFloat(preValue,10);
            console.log(preValue);
          }else{
            grid_target.setValue(grid_orig.getValues());
          }        
          
        }
        // if(i==D){
        //   preValue=parseFloat(grid_orig.getValues(),10);//+
        //   //(sheetlastMonth.getRange(j-D,1).getValue()==""?"":parseFloat(sheetlastMonth.getRange(j-D,1).getValue(),10));
        // }
      }else{  //for date
        grid_target.setNumberFormat('@');
        grid_target.setValue(grid_orig.getValues());
      }
    }
  }
}

function wrapCopyAndAcc4Grids(){
  CopyAndAcc4Grids(sheet_no_acc,sheet_acc);
} 





