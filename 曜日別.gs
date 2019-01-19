/* メニュー */
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [{name: '列削除', functionName: 'clearColumn'},
               {name: '対応予定セット', functionName: 'taiouset'},
               {name: '今日終わりセット',functionName:'today'},
               {name: 'ABセット', functionName: 'ABset'},
               {name: '現役セット', functionName: 'genekiset'},
               {name: '現役模試不要セット',functionName:'genekimoshi'},
               {name: '月一不要セット',functionName:'month'}
              ];

  ss.addMenu('メニュー', menus);
}

//------------------------------------------
//以下未使用（01課題作成に移動）2018/8/7
//------------------------------------------
/* 列削除 */
function clearColumn(){
    var sheet = SpreadsheetApp.getActiveSheet();
    var column = sheet.getActiveCell().getColumn();
    var lastrow = sheet.getLastRow();
  
  var content = sheet.getRange(1,column).getValue();
  var check = Browser.msgBox(content +"を削除しますか?",Browser.Buttons.OK_CANCEL);
    if(check == "ok"){
      sheet.getRange(3, column,lastrow-2,1).clearContent(); 
      Browser.msgBox("削除しました");
  }else{
    Browser.msgBox("キャンセルしました");
  }  
  }

//退塾メンテ
function maintenance(){
    var sheet = SpreadsheetApp.getActiveSheet();
    var row = sheet.getActiveCell().getRow();
    var lastColumn = sheet.getLastColumn();
  
  var name = sheet.getRange(row,2).getValue();
  var check = Browser.msgBox(name +"を削除しますか?",Browser.Buttons.OK_CANCEL);
    if(check == "ok"){
      sheet.getRange(row,2,1,lastColumn-1).clearContent(); 
      Browser.msgBox("削除しました");
  }else{
    Browser.msgBox("キャンセルしました");
  }  
  }


//カスタム関数　例　　=color("S4")
function color(range){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var color = sheet.getRange(range).getBackground();
  return color;
}




//対応予定セット
function taiouset(){
  koushiset(0); 
}

//現役セット
function genekiset(){
  koushiset(1); 
}

//ABセット
function ABset(){
  koushiset(2); 
}

//D黒セット
function Dkuroset(){
  koushiset(3); 
}

//今日終わりセット
function today(){
 koushiset(4) ;
}

function koushiset(taiou){
  var　sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var now = new Date().getDay();
  var column = now+6;
  var _ = Underscore.load();
  
  var data1 = sheet.getRange("G3:L").getValues(); //曜日を配列で取得
  var data2 = _.zip.apply(_,data1); 
  var data3  = sheet.getRange("Q3:Q").getValues(); //対応予定列を取得
  var teachers = sheet.getRange("Q2:T2").getValues(); //講師の生徒評価別リストを配列で取得
  
  var color = sheet.getRange("S3:S").getBackgrounds() ;//対応日の色を取ってくる
  var data4 = sheet.getRange("N3:N").getValues();//振替列を取得
  var data5 = sheet.getRange("O3:O").getValues();//現役かどうかを取得
  var data6 = sheet.getRange("Z3:Z").getValues();//対応分類
  var data7 = sheet.getRange("M3:M").getValues();//通塾最終曜日
  var now = new Date().getDay();
  var dayNames = '日月火水木金土';
  for (i=0;i<data1.length;i++){
    if(taiou == 1 && ((data2[now-1][i]=='1'&& color[i]=='#f4c7c3'&&data4[i] == '')||(data4[i]=='出席'&&color[i]=='#f4c7c3'))&&data3[i]==''&&data5[i]=='1'){  
      data3[i]=["対応"]; //現役セット
  }else if(taiou == 0 && ((data2[now-1][i]=='1'&& color[i]=='#f4c7c3'&&data4[i] == '')||(data4[i]=='出席'&&color[i]=='#f4c7c3'))&&data3[i]==''){
        data3[i]=["対応"]; //対応セット
  }else if(taiou == 2 && ((data2[now-1][i]=='1'&&data4[i]=='')||data4[i]=="出席")&&(data6[i]=="A"||String(data6[i]).indexOf("B")!=-1)&&data3[i]==''){
        data3[i]=["対応"]; //ABセット
  }else if(taiou == 3 && ((data2[now-1][i]=='1'&&data4[i]=='')||data4[i]=="出席")&&(String(data6[i]).indexOf("D")!=-1||String(data6[i]).indexOf("黒")!=-1)&&data3[i]==''&&color[i]=='#f4c7c3'){
      data3[i]=["対応"]; //D黒セット
      
  }else if(taiou == 4 && ((data2[now-1][i]=='1'&& data7[i]==(dayNames[now])&& color[i]=='#f4c7c3'&&data4[i] == ''))){
        data3[i]=["対応"]; //今日終わりセット
  }            
  sheet.getRange("Q3:Q").setValues(data3);    
}
}


//現役模試不要セット
function genekimoshi(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var column = sheet.getActiveCell().getColumn();
  var data1 = sheet.getRange("O3:O").getValues();//現役かどうか
  var lastrow = sheet.getLastRow();
  var data2 = sheet.getRange(3,column,lastrow-2,1).getValues();
  var content = sheet.getRange(1,column).getValue();
  var check = Browser.msgBox(content +"にセットしますか?",Browser.Buttons.OK_CANCEL);
  if (check =="ok"){
    for(i=0;i<data1.length;i++){
      if (data1[i][0]=="1"){data2[i][0]="不要"};
      }sheet.getRange(3,column,lastrow-2,1).setValues(data2);
      Browser.msgBox("変更しました");
   }else {
      Browser.msgBox("キャンセルしました");
}
} 
 
//曜日をとる
function day(){
  var　sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var now = new Date().getDay();
  var dayNames = '日月火水木金土';
  Logger.log(dayNames[now]);
}

// 色を取得（カスタム関数）
function color(range){
  var sheet = SpreadsheetApp.getActiveSheet();
  var color = sheet.getRange(range).getBackgrounds();
  return color;
}


//-------------------------------------------------------------------------
//  入力した生徒行の対応者(ログインユーザ)、対応日をセットし、対応予定セルをクリアする
//--------------------------------------------------------------------------
function mendanSet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("曜日別");
  var _ = Underscore.load();
 
  var arr1 = sheet.getRange("A1:A").getValues();
  var arr2 = _.zip.apply(_,arr1);
  
  //01. 【講師シフト表】＞対応講師一覧　※こちらの参照権限もつけないとエラーになる
  var sheet2 = SpreadsheetApp.openById("1lD9gWxYZxzxCboEaIg5h2gEA3lgMNlOfqbvqVYmtKyQ").getSheetByName("対応講師一覧");
  var arrTeacher =sheet2.getRange("B3:D").getValues();
  var arrTeacher2 =_.zip.apply(_,arrTeacher); 
 
  var date = new Date();
  var usr  =Session.getActiveUser().getEmail();
  
  if(arrTeacher2[1].indexOf(usr) > -1){
    var idx = arrTeacher2[1].indexOf(usr);
    var teacherNm = arrTeacher2[0][idx];
  }else{
    Browser.msgBox("01.【講師シフト表】＞対応講師一覧にないログインアカウントの為処理を中断します"+usr);
    return;    
  }

  var stuno = Browser.inputBox('面談する生徒番号を入力して下さい', Browser.Buttons.OK_CANCEL);
  if (stuno != 'cancel' && arr2[0].indexOf(parseInt(stuno))!=-1){
    var row = parseInt(arr2[0].indexOf(parseInt(stuno))+1);
    
    sheet.getRange(row,sheet.getRange('Q1').getColumn()).clearContent();
    sheet.getRange(row,sheet.getRange('R1').getColumn()).setValue(teacherNm);
    sheet.getRange(row,sheet.getRange('S1').getColumn()).setValue(date);
  }
  
}


//月一不要セット
function month(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var column = sheet.getActiveCell().getColumn();
  var data1 = sheet.getRange("O3:O").getValues();//浪人かどうか
  var lastrow = sheet.getLastRow();
  var data2 = sheet.getRange(3,column,lastrow-2,1).getValues();
  var content = sheet.getRange(1,column).getValue();
  var check = Browser.msgBox(content +"にセットしますか?",Browser.Buttons.OK_CANCEL);
  if (check =="ok"){
    for(i=0;i<data1.length;i++){
      if (data1[i][0]=="○"){data2[i][0]="不要"};
      }sheet.getRange(3,column,lastrow-2,1).setValues(data2);
      Browser.msgBox("変更しました");
   }else {
      Browser.msgBox("キャンセルしました");
}
} 


function user(){
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var row = sheet.getActiveCell().getRow();
 var today = new Date();
 var date = today.getMonth()+1+"/"+today.getDate();
 var teacher={'aceacademy@delf.co.jp':'高梨',
              'shinesace.kk@gmail.com':'test',
              'yuikodera@gmail.com':'小寺',
              'mana1112chopin0522@gmail.com':'橋本',
              'ryokaboc79@gmail.com':'藤原',
              'yurihonma011@gmail.com':'本間',
              'yygd26chocolate@gmail.com':'鈴木',
              'tomonomisa827@gmail.com':'伴野',
              'yuta1226g@gmail.com':'三松',
              'tomochan1301@gmail.com':'篠原',
              'mkanno.9825@gmail.com':'菅野',
              'm.nibu0213@gmail.com':'仁部'}; 

 var user = Session.getActiveUser();
 var mail = user.getEmail();

 sheet.getRange("Q"+row).clearContent();
  sheet.getRange("R"+row).setValue(teacher[mail]);
 sheet.getRange("S"+row).setValue(date);   
}
