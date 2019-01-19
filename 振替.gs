/*************************************************************************/
/* 自動メール送信（自動送信用）2016/08/27 更新 */ 
/*************************************************************************/
function Auto_SendMail3(){
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSheet();                  // シートを取得
  var sheetName=mySheet.getName();　　　　　　　　　　　　　　　// シート名を取得

  var reg1 = /申し訳ございません。定員に達しているため、別日にて再度申請ください。/; //add 
  var _ = Underscore.load();                                        //add (日付フォーマット用）
  var aryweekday = ["(日)","(月)","(火)","(水)","(木)","(金)","(土)"]; //add (日付フォーマット用）
                            
  var rowSheet=mySheet.getRange("登録行数").getValue();         // 登録行数を取得
  
  var arychk1 = mySheet.getRange(11,5,rowSheet-10,2).getValues();　//add 重複行チェック No-name
  var arychk2 = mySheet.getRange(11,8,rowSheet-10,2).getValues();　//add 重複行チェック 欠席-出席
  
  if(sheetName=="振替"){
  
    // 各種データをセット
    var strFrom = mySheet.getRange("送信元3").getValue(); 
    var strSubject = mySheet.getRange("自動タイトル3").getValue();
    var strBunto = mySheet.getRange("自動返信文頭3").getValue();
    
    // 各種項目の設定列数を取得
    var aryCol = new Array();      
    aryCol[0] = mySheet.getRange("生徒番号3").getValue();
    aryCol[1] = mySheet.getRange("氏名3").getValue();
    aryCol[2] = mySheet.getRange("メールアドレス3").getValue();
    aryCol[3] = mySheet.getRange("開始項目3").getValue();
    aryCol[4] = mySheet.getRange("終了項目3").getValue();
    
    // 出力項目の確認
    var aryCheck = new Array();
    var aryField = new Array();
    var chkRow = 8;   // 自動送信チェック行
    var c = 0;        // カウント用
    for(var i = aryCol[3]; i <= aryCol[4]; i++){    
      aryCheck[c]=mySheet.getRange(chkRow, i).getValue();        // 出力確認
      aryField[c]=mySheet.getRange(chkRow + 2, i).getValue();    // 項目名
      c = c + 1;
    }
    
    // メール送信
    var outRow = 11;   // データ開始行
    var chkSend=mySheet.getRange(rowSheet,2).getValue();              // 送信チェック      
    var strNo=mySheet.getRange(rowSheet,aryCol[0]).getValue();        // 生徒番号
    var strName=mySheet.getRange(rowSheet,aryCol[1]).getValue();      // 氏名
    var strTo=mySheet.getRange(rowSheet,aryCol[2]).getValue();        // メールアドレス
    var aryBody = new Array();                                 // 総評　等
    var c = 0,wkdate='';
    for(var j = aryCol[3]; j <= aryCol[4]; j++){    
      // add_181101 start（セルの値が日付のときフォーマット）
        if(_.isDate(mySheet.getRange(rowSheet,j).getValue())){
          wkdate=mySheet.getRange(rowSheet,j).getValue();
          aryBody[c]= Utilities.formatDate(wkdate,"JST","yyyy/MM/dd")+aryweekday[wkdate.getDay()];  
          
        }else{   
      // add_181101 end
          
          aryBody[c]=mySheet.getRange(rowSheet,j).getValue();
        }      
      
      c = c + 1;
    }
    
    // 本文作成
    var strBody= "" + "\n\n";   //本文の最初（""内）を好きに変えてOK（生徒番号と名前を表示したい時は、"生徒番号" + strNo + "　" + strName + " 様\n\n"を入力）
    
    // 文頭の文章をセット
    strBody = strBunto + "\n\n"
    
    // 他の本文の文章をセット
    for(var k = 0; k <= aryCol[4] - aryCol[3]; k++){    
      if(aryCheck[k] == "※") {
        if(aryField[k] != "文頭（自動返信用）") {
          strBody = strBody + "【" + aryField[k] + "】\n" + " " + aryBody[k] + "\n\n"
        }
      }
    }
    //var lastCol=mySheet.getLastColumn()    // 送信完了日の列数（C列）
    
    
    //送信前のスリープ 1sec
    Utilities.sleep(1000);
    
    // メールを送信（添付ファイルがある場合とない場合で処理分け）
    MailApp.sendEmail(strTo, strSubject, strBody);
    
    //add　本当の振替不可をメモしておく
    if(reg1.test(strBody)){
      //mySheet.getRange(rowSheet,4,1,8).setBackground("lightgrey");
      mySheet.getRange(rowSheet,4).setValue('【振替不可】'+mySheet.getRange(rowSheet,4).getValue());
    }
    
    //重複行ﾁｪｯｸ
    for(var i=0;i<arychk1.length;i++){
      for(var j=i+1;j<arychk1.length;j++){
        if(arychk1[i][0]+arychk1[i][1]+arychk2[i][0]+arychk2[i][1]==arychk1[j][0]+arychk1[j][1]+arychk2[j][0]+arychk2[j][1] && arychk1[j][0] != '重複'){
        //重複する行の日付は1日前にしておく
          arychk1[j][0] ='重複';           
          arychk2[j][0] = '';
          arychk2[j][1] = '';
        }  
      }  
    }
    mySheet.getRange(11,5,rowSheet-10,2).setValues(arychk1);
    mySheet.getRange(11,8,rowSheet-10,2).setValues(arychk2);    
    
    //add-end
    
    //ドキュメントの内容をログに表示
    //Logger.log(strBody);
  }
}

/*************************************************************************/
/* 不要行削除　2016/09/10 更新 */ 
/*************************************************************************/
function DeleteRows() {
  var bk = SpreadsheetApp.getActiveSpreadsheet();
  var mySheet = bk.getSheetByName("振替");                    // シートを取得
  var myRow     = mySheet.getRange("登録行数").getValue();      // 登録行数を取得
  var sheetName = mySheet.getName();　　　　　　　　　　　　　　// シート名を取得
  
  for(var i = myRow; i > 10; --i){ 
    var cdate = new Date();                           // 今日
    var kdate = mySheet.getRange(i, 8).getValue();    // 欠席日
    var sdate = mySheet.getRange(i, 9).getValue();    // 出席日
    
    // 欠席日と出席日に今日より前の日付が設定された行を削除
    if (mySheet.getRange(i, 4).getValue() != '') {
      if (kdate < cdate && sdate < cdate) {
        mySheet.deleteRow(i);
      }
    }
  }
}
