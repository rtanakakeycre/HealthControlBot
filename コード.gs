/*
function doPost(e){
  var params = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(params.challenge);
}
*/

function doPost(e) {
  
  
  try{
    // トークンからスラックへのリンクを取得
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
    const slackApp = SlackApp.create(token);
    
    // ポストデータからパラメータを取得
    const params = JSON.parse(e.postData.getDataAsString());
    writeLog(params);

    return ContentService.createTextOutput(params.challenge);
    
  }catch(err){
    writeLog(err);
  }
}

// スプレッドシートの行、列情報
const ROW_DATE = 0;
const ROW_DATA = 1;
const COL_NAME = 0;
const COL_ID = 1;
const COL_DATA = 2;

function writeLog(params){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");
  const channel = params.event.channel;

  Logger.log(params);
  
  let txName1 = getUserName(params.event.user);

  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // 名称列から名称を取得
    let txName2 = sheet1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
    if(txName2 == txName1){
      break;
    }
  }
  
  // 今日の日付を取得
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "_" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // 今日の日付列を取得
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // テキストから体温のみ抽出
  let txVal1 = params.event.text;
  let idChrSta1 = txVal1.indexOf("\n");
  txVal1 = txVal1.slice(idChrSta1 + 1);
  
  // 日付を更新
  sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // 名前を更新
  sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // 念のためIDを更新
  sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(params.event.user);
  // 今日の体温を更新
  sheet1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // デバッグ用
  //sheet1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(params);
}

// ユーザ名を取得
function getUserName(userId){
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+token+"&user="+userId).getContentText();

  const userName = JSON.parse(userData).user.real_name;
  Logger.log(userId,userName);

  return userName ? userName : userId; 
}

// メッセージを送信
function SendMessage(){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("体温管理");

  var url = "https://slack.com/api/chat.postMessage";
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
  var channel = "C019TVCKLEB";
  
  // 今日の日付を取得
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "_" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // 今日の日付列を取得
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
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
      
      // 変更するのは、この部分だけ!
      var payload = {
        "token" : token,
        "channel" : channel,
        "text" : txName2 + "\n体温を入力してください。"
      };
      
      var params = {
        "method" : "post",
        "payload" : payload
      };
      
      // Slackに投稿する
      let res1 = UrlFetchApp.fetch(url, params);
    }
  }

}