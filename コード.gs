/*
function doPost(e){
  var params = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(params.challenge);
}
*/

// https://script.google.com/macros/s/AKfycby1lxGnLphhMjy-WrLsKglK5ZgEwBUFHVA_VTLUwD_QBIYrlOU/exec
function doPost(e) {
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");
  let sheet2 = spreadSheet1.getSheetByName("勤怠管理");
    
  try{
    // トークンからスラックへのリンクを取得
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
    const slackApp = SlackApp.create(token);
    
    // ポストデータからパラメータを取得
    const params = JSON.parse(e.postData.getDataAsString());
    //sheet2.getRange(1, 1).setValue(params);
    
    if(params.type == "SheetWrite"){
      // シート書き込み要求
      writeLog2(params.name, params.id, params.AftSvr, params.AftChn, params.BefSvr, params.BefChn);
    }
    else{
      // slackからの情報
      writeLog(params);
    }    

    return ContentService.createTextOutput(params.challenge);
    
  }catch(err){
    writeLog(err);
  }
}



// スプレッドシートの行、列情報
const ROW_DATE = 3 - 1;
const ROW_DATA = 4 - 1;
const N_ROW_DATA = 3;
const ROW_DATA_HIS = 1 - 1;
const ROW_DATA_STA = 2 - 1;
const ROW_DATA_END = 3 - 1;
const COL_NAME = 2 - 1;
const COL_ID = 1 - 1;
const COL_DATA = 5 - 1;
const COL_SVR = 4 - 1;
const COL_STT = COL_SVR + 1;
const COL_PLACE = COL_STT + 1;
const COL_CPT = COL_PLACE + 1;
const COL_DATA2 = COL_CPT + 1;

function Test()
{
  //writeLog2("田中良平", "730250456168792000", "", "", "テストサーバー", "テストチャンネル");
  writeLog2("田中良平", "730250456168792000", "テストサーバー", "テストチャンネル2", "", "");
}

function Test2()
{
  writeLog2("田中良平", "730250456168792000", "", "", "テストサーバー", "テストチャンネル");
  //writeLog2("田中良平", "730250456168792000", "テストサーバー", "テストチャンネル2", "", "");
}

function writeLog2(txName1, txId1, txAftSvr1, txAftChn1, txBefSvr1, txBefChn1){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("勤怠管理");
   
  let flNew1 = true;
  
  // 今日の日付を取得
  var date = Moment.moment();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "'" + date.format("MM/DD");
  let txTime1 = date.format("HH:mm");
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1 += N_ROW_DATA){
    // ID列からIDを取得
    let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    if(txId2 == txId1){
      // IDがすでに存在
      flNew1 = false;
      break;
    }
  }
  
  if(flNew1){
    // IDを記載
    sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
    
    // 見出しを記載
    sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + COL_CPT).setValue("勤怠履歴");
    sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_CPT).setValue("出勤");
    sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + COL_CPT).setValue("退勤");
    
    // 行のグループ化を行います。
    sheet1.getRange(1 + idRow1 + 1, 1, N_ROW_DATA - 1).shiftRowGroupDepth(1);
  }
  // 名前を更新
  sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
    
  let txChn1 = "C01AG9H3GBF";
  if(txAftChn1 == "出社"){
    // 場所を更新
    if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
      PostMessage(txName1 + "がテレワークを終了し、出社しました。", txChn1);
      sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
    }
  }
  else if(txAftChn1 == "テレワーク開始"){
    // 場所を更新
    if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() == ""){
      PostMessage(txName1 + "がテレワークを開始しました。", txChn1);
      sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("テレワーク中");
    }
  }
  else if(txAftChn1 == "退勤"){
    if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
      PostMessage(txName1 + "がテレワークを終了しました", txChn1);
      sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
    }
  }
  
  if(txAftChn1 == "出社" || txAftChn1 == "退勤"){
    // 場所を更新
    sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
  }
  else if(txAftChn1 == "テレワーク開始"){
    // 場所を更新
    sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("テレワーク中");
  }
    
  // 今日の日付列を取得
  let idCol1 = COL_DATA2;
  let flNew2 = true;
  for(idCol1 = sheet1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
    // 日付行から日付を取得
    let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      flNew2 = false;
      break;
    }
  }
  
  if(flNew2){
    idCol1 = sheet1.getLastColumn();
    // 日付を更新
    sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  }
  
  let flTime1 = true; 
  if(txAftChn1 == ""){
    // 退室時間を記録
    sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_STT).setValue(date.format("YYYY-MM-DD HH:mm"));
  }
  else{
    let txDate2 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_STT).getValue();
    if(txDate2 != ""){
      let date2 = Moment.moment(txDate2);
      if(date.diff(date2, 'hours') >= 2){
        // 退室から2時間以上
        UpdHis(sheet1, idRow1, idCol1, "", "退勤", date2.format("HH:mm"), false);
      }
      else{
        flTime1 = false;
      }

      // 退室時間を削除
      sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_STT).setValue("");
    }
  }
    
  UpdHis(sheet1, idRow1, idCol1, txAftSvr1, txAftChn1, txTime1, flTime1);

  // 状態を更新
  sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
  // サーバを更新
  sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
}

// ============================================================================
// 勤怠履歴更新
// ============================================================================
function UpdHis(sheet1, idRow1, idCol1, txAftSvr1, txAftChn1, txTime1, flTime1){
  let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).getValue();
  let txEnd1 = sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).getValue();
  if(txAftChn1 == "出社" || txAftChn1 == "テレワーク開始"){
    if(txSta1 == ""){
      // 出社ずみでなければ記録
      sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).setValue(txTime1);
    }
  }
  else{
    if(txSta1 == ""){
      // 出社済みでなければ日を跨いでいるので、昨日の扱いとする
      idCol1 = idCol1 - 1;
    }
  }

  let txHis1 = sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).getValue();
  
  if(idCol1 >= COL_DATA2){
    if(txAftChn1 == "退勤"){
      // 退勤済みでも記録
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).setValue(txTime1);
    }
  
    if(txHis1 == ""){
      txHis1 = txHis1 + "["+ txTime1 + "]";      
    }
    else{
      if(txAftChn1 == "" || flTime1){
        // 出るときは出た時間を記録
        txHis1 = txHis1 + "⇒";
        txHis1 = txHis1 + "["+ txTime1 + "]";              
      }
      else{
        // 入ってくるときは最後に出た時間が入ってるはずなので、何も記載しない
      }
    }
    
    if(txAftSvr1 == "" || txAftSvr1 == "KEY_勤怠管理"){
      
    }
    else{
      if(sheet1.getRange(1 + idRow1, 1 + COL_SVR).getValue() != txAftSvr1){
        txHis1 = txHis1 + "(" + txAftSvr1 + ")";
      }
    }
    txHis1 = txHis1 + txAftChn1;
    
    // 履歴を更新
    sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).setValue(txHis1);
  }
}

// 体温を更新
function writeLog(params){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");
  const channel = params.event.channel;

  Logger.log(params);
    
  let txName1 = getUserName(params.event.user);
  let txId1 = params.event.user;

  if(txName1 == "KintaiKanri"){
    return;
  }
  
  // テキストから体温のみ抽出
  let txVal1 = params.event.text;
  let idChrSta1 = txVal1.indexOf("\n");
  txVal1 = txVal1.slice(idChrSta1 + 1);
  
  if(txVal1.match(/^\d\d\d$/g)) {
    // (365)
    txVal1 = txVal1.substr(0,2) + "." + txVal1.substr(2,1);
  }   
  else if(txVal1.match(/^\d\d(\.\d)?$/g)) {
    // (36.0、36)
  } 
  else{
    // 体温ではない(36、36.5、365)
    PostMessage("「"+ txVal1 + "」"+ "\n入力が無効です。\n例:36.2、36、362", "@" + txId1);
    return;
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // ID列からIDを取得
    let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    if(txId2 == txId1){
      break;
    }
  }
  
  // 今日の日付を取得
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // 今日の日付列を取得
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // 日付を更新
  sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // 名前を更新
  sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // 念のためIDを更新
  sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
  // 今日の体温を更新
  sheet1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // デバッグ用
  //sheet1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(params);
}




// ユーザ名を取得
function getUserName(userId){
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+token+"&user="+userId).getContentText();

  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");

  const userInfo = JSON.parse(userData).user;  
  const userProf =userInfo.profile;
  const userName1 = userProf.display_name;
  const userName2 = userInfo.real_name;

  return userName1 ? userName1 : (userName2 ? userName2 : userId); 
}

// メッセージを送信
function SendMessage(){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");

  // 今日の日付を取得
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // 今日の日付列を取得
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // 名称列から名称を取得
    
    let txOndo2 = sheet1.getRange(1 + idRow1, 1 + idCol1).getValue();
    if(txOndo2 == ""){
      // 温度が記入されていない
      let txName2 = sheet1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
      let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      PostMessage("体温を入力してください", "@"+txId2);
    }
  }
}

function PostMessage(txMsg1, txChn1){
  var url = "https://slack.com/api/chat.postMessage";  
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');

  var payload = {
      "token" : token,
      "channel" : txChn1,
      "text" : txMsg1
    };
    
    var params = {
      "method" : "post",
      "payload" : payload
    };
    
    // Slackに投稿する
    let res1 = UrlFetchApp.fetch(url, params);
}
