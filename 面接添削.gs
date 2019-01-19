/*************************************************************************/
// メール送信(汎用）
/*************************************************************************/
function Manual_SendMail_hanyou(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 各種項目の設定列・行を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
  aryCol[1] = mySheet.getRange("e4").getValue();//開始項目
  aryCol[2] = mySheet.getRange("e5").getValue();//終了項目

  var chkcol = mySheet.getRange("e6").getValue();    //※の入力列
  var timecol = mySheet.getRange("e7").getValue();    //送信タイムスタンプ列    
  var chkRow = mySheet.getRange("C3").getValue();   // 手動送信チェック行
  var outRow = mySheet.getRange("C4").getValue();; 　  // 送信対象開始行
  var strTitle = mySheet.getRange("c5").getValue(); //送信タイトル    
　//---  
  
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);

}

//メール送信
//引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
//
function sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle){
  var c = 0;        // カウント用           
  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得

  var aryCheck = new Array(),aryField = new Array();
  aryCheck=mySheet.getRange(chkRow, aryCol[1],1,aryCol[2]-aryCol[1]+1).getValues();        // 出力確認
  aryField=mySheet.getRange(chkRow+1, aryCol[1],1,aryCol[2]-aryCol[1]+1).getValues();    // 項目名
            
  var chkSend=mySheet.getRange(1,chkcol,rowSheet).getValues();         // 送信チェック（※）    

  for(var i = outRow; i <= rowSheet; i++){

      if(chkSend[i-1][0] == "※") {        
        
        // 本文データを取得
        var strTo=mySheet.getRange(i,aryCol[0]).getValue();      // メールアドレス
        var aryBody = new Array();                               // 総評　等
        var c = 0;

        aryBody=mySheet.getRange(i,aryCol[1],1,aryCol[2]-aryCol[1]+1).getValues();

        // 本文作成
        var attachmentFiles = new Array();
        var strBody= "" + "\n";   //本文の最初（""内）を好きに変えてOK（生徒番号と名前を表示したい時は、"生徒番号" + strNo + "　" + strName + " 様\n\n"を入力）
        for(var k = 0; k <= aryCol[2] - aryCol[1]; k++){    
          if(aryCheck[0][k] == "✔") {
            if(aryField[0][k] == "文頭") {
              strBody = strBody + aryBody[0][k] + "\n\n"
            }
          }
        }
        
        for(var k = 0; k <= aryCol[2] - aryCol[1]; k++){    
          if(aryCheck[0][k] == "✔") {
            if(aryField[0][k] == "文頭") {
              // 何もしない
            } else {
              strBody = strBody + "【" + aryField[0][k] + "】\n"  + aryBody[0][k] + "\n\n"
            }
          }
        }
      
        //送信前のスリープ 1sec
        Utilities.sleep(1000);
        
        // メール送信
        MailApp.sendEmail(strTo, strTitle, strBody);

        mySheet.getRange(i,timecol).setValue(new Date())
        mySheet.getRange(i,chkcol).setValue("")
               
      }
    }      
  return;
}