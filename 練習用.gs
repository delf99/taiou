
function test(){
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("練習用");
  sheet.getRange("AB4:AB").clearContent();
}  


function test2(){
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("練習用");
   var data1 = sheet.getRange("B4:B").getValues() ;
   Logger.log(data1[0]);
   data1[0][0]="鈴木";
   sheet.getRange("AB4:AB").setValues(data1);
 
}  


function test3(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("練習用");
  var data2 = sheet.getRange("B4:B").getValues();
  Logger.log(data2);
  //data2[0][0]="";
  //sheet.getRange("A4:AB").setValues(data2);
  
}

function getBackgroundmyFanction(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("練習用");
  var color = sheet.getRange("S4:S").getBackgrounds();
  Logger.log(color);

}   

function test10(){
for(i=0;i<30;i=i+2){
 Logger.log(i);
}
}


/* 
やりたいこと→対応日をみて、赤色の人を読み取り、その人には本日対応できるように対応と表示させる。

// 現在アクティブなスプレッドシートを取得
var ss = SpreadsheetApp.getActiveSpreadsheet();
// そのスプレッドシートにある最初のシートを取得
var sheet = ss.getSheets()[0];

// そのシートにある B5 セルを取得
var cell = sheet.getRange("B5");
// そのセルに設定されている背景色を取得しログに出力
Logger.log(cell.getBackground());
*/