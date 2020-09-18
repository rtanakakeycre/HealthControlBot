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
const COL_OUT = COL_PLACE + 1;
const COL_CPT = COL_OUT + 1;
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

// 勤怠情報解析
function writeLog2(txName1, txId1, txAftSvr1, txAftChn1, txBefSvr1, txBefChn1){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("勤怠管理");
   
  let flNew1 = true;
  // IDから対応メンバーの行を取得
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
    
  // 今日の日付を取得
  var dateNow1 = Moment.moment();

  if(txAftChn1 == ""){
    // 退室の場合、退室時間を記録
    sheet1.getRange(1 + idRow1, 1 + COL_OUT).setValue(dateNow1.format("YYYY-MM-DD HH:mm"));
    // 状態を消去
    sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    // サーバを消去
    sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
  }
  else{
    // 入室の場合、前回の退室からの時間を取得
    let txDate2 = sheet1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
    if(txDate2 == ""){
      // ただの部屋移動
      ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
    }
    else{
      // 前回が退室の場合は「サーバ移動」、「退室⇒入室」、「退勤⇒出勤」のいずれか      
      let dateOut1 = Moment.moment(txDate2);
      // 退室時間(分)
//      let ctOutTime1 = dateNow1.diff(dateOut1, 'minutes');
      //sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_SVR).setValue(ctOutTime1);
      ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      
//      if(ctOutTime1 >= 2){
//        // 退室から2分以上は退室扱い
//        ChnSftExe(sheet1, dateOut1, idRow1, "", "退室", txName1);
//        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
//      }
//      else{
//        // 退室から2分以内はサーバ移動
//        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
//      }
//      // 退室時間を削除
//      sheet1.getRange(1 + idRow1, 1 + COL_OUT).setValue("");
    }
  }
}

// ============================================================================
// チャンネル移動実行
// ============================================================================
function ChnSftExe(sheet1, date1, idRow1, txAftSvr1, txAftChn1, txName1){
  try{
    // 打刻日付は翌日の6:00を区切りとする
    let txDate1 = date1.subtract(6, "h").format("MM/DD");
    let txTime1 = date1.format("HH:mm");
    
    // 打刻日付を取得
    let idCol1 = 0;
    // 日付行を右から検索
    for(idCol1 = sheet1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
      // 日付行から日付を取得
      let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
      if(txDate2 == txDate1){
        // 発見
        break;
      }
    }
    
    if(idCol1 < COL_DATA2){
      // 日付がなければ右端に追加
      idCol1 = sheet1.getLastColumn();
      sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue("'" + txDate1);
    }
    
    let idCol2 = idCol1;
    let ctDayMnt1 = date1.startOf('day').diff(date1, 'minutes');
    if(ctDayMnt1 >= 10 * 60){
      // 10時以降なら前日の退勤がなかろうが本日の打刻として扱う
    }
    else{
      // 10時以前なら念のため、前日の退勤までチェック
      let idDay1;
      let ctDay1 = 2;
      for(idCol2 = idCol1, idDay1 = 0; idCol2 >= COL_DATA2, idDay1 < 2; idCol2--, idDay1++){
        // 今日の日付から直前の出社打刻ありの日付を検索
        let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
        if(txSta1 != ""){
          // 出社打刻あり      
          let txEnd1 = sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).getValue();
          if(txEnd1 != ""){
            // 出社打刻も退勤打刻もありの場合は現在の日付を打刻日付とする
            idCol2 = idCol1;
          }
          else{
            // 出社打刻ありで退勤打刻なしの場合はその日付を打刻日付とする
          }
          break;
        }    
      }
      
      if(idDay1 < COL_DATA2 || idDay1 >= 2){
        // 出社打刻が見つからなければ現在の日付を打刻日付とする
        idCol2 = idCol1;
      }
    }    
    
    if(txAftChn1 == "退勤"){
      // 退勤
      // 退勤時は気にせず退勤打刻を行う
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).setValue(txTime1);
    }
    else{
    //else if(txAftChn1 == "出社" || txAftChn1 == "テレワーク開始"){
      // 退勤以外は出勤扱い
      // 出勤
      let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
      if(txSta1 == ""){
        // 出勤打刻なしなら出勤打刻を行う
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).setValue(txTime1);
      }
    }
    
    //let txChn1 = "C01AG9H3GBF";
    let txChn1 = "C01805HS02F";
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
    
    let txHis1 = sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).getValue();
    
    if(txHis1 != ""){
      txHis1 = txHis1 + "⇒";
    }
    
    // 時間を記録
    txHis1 = txHis1 + "["+ txTime1 + "]";      
    let txNowSvr1 = sheet1.getRange(1 + idRow1, 1 + COL_SVR).getValue(); 
    if(txAftSvr1 == "" || txAftSvr1 == "KEY_勤怠管理"){
      // 特殊サーバは無視
    }
    else{
      if(txAftSvr1 != txNowSvr1){
        // サーバ移動が発生していれば追記
        txHis1 = txHis1 + "(" + txAftSvr1 + ")";
      }
    }
    
    txHis1 = txHis1 + txAftChn1;
    
    // 履歴を更新
    sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).setValue(txHis1);
    
    if(txAftSvr1 != ""){
      // 状態を更新
      sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
      // サーバを更新
      sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
    }
  }
  catch(e){
    let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
    let sheetErr1 = spreadSheet1.getSheetByName("エラーログ");
    sheetErr1.getRange(1 + sheetErr1.getLastColumn(), 1).setValue(e);
  }
}

// ============================================================================
// 体温を更新
// ============================================================================
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
