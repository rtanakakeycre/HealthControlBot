/*
function doPost(e){
  var jsPrm1 = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(jsPrm1.challenge);
}
*/


// エントリポイント
// https://script.google.com/macros/s/AKfycby1lxGnLphhMjy-WrLsKglK5ZgEwBUFHVA_VTLUwD_QBIYrlOU/exec

// ============================================================================
// POSTのハンドラ
// ============================================================================
function doPost(e) {
   
  try{    
    // ポストデータからパラメータを取得
    const jsPrm1 = JSON.parse(e.postData.getDataAsString());
    
    if(jsPrm1.type == "SheetWrite"){
      // シート書き込み要求
      UpdKintai(jsPrm1.name, jsPrm1.id, jsPrm1.AftSvr, jsPrm1.AftChn, jsPrm1.BefSvr, jsPrm1.BefChn);
    }
    else{
      // slackからの情報
      UpdTaion(jsPrm1);
    }    

    return ContentService.createTextOutput(jsPrm1.challenge);
    
  }catch(err){
    UpdTaion(err);
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
const COL_SLK_ID = 3 - 1;
const COL_DATA = 5 - 1;
const COL_SVR = 4 - 1;
const COL_STT = COL_SVR + 1;
const COL_PLACE = COL_STT + 1;
const COL_OUT = COL_PLACE + 1;
const COL_CPT = COL_OUT + 1;
const COL_DATA2 = COL_CPT + 1;

function Test()
{
  //UpdKintai("田中良平", "730250456168792000", "", "", "テストサーバー", "テストチャンネル");
  UpdKintai("田中良平", "730250456168792000", "テストサーバー", "テストチャンネル2", "", "");
}

function Test2()
{
  UpdKintai("田中良平", "730250456168792000", "", "", "テストサーバー", "テストチャンネル");
  //UpdKintai("田中良平", "730250456168792000", "テストサーバー", "テストチャンネル2", "", "");
}

// ============================================================================
// 勤怠情報の解析と更新
// ============================================================================
function UpdKintai(txName1, txDscId1, txAftSvr1, txAftChn1, txBefSvr1, txBefChn1){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("勤怠管理");

  // DiscordIDは「_」付きで管理しているため「_」を付加
  let txId1 = "_" + txDscId1;
   
  let flNew1 = true;
  // IDから対応メンバーの行を取得
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1 += N_ROW_DATA){
    // ID列からIDを取得
    let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    if(txId2 == txId1){
      // IDがすでに存在
      flNew1 = false;
      break;
    }
  }
  
  if(flNew1){
    // IDを記載
    sht1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
    
    // 見出しを記載
    sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + COL_CPT).setValue("勤怠履歴");
    sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_CPT).setValue("出勤");
    sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + COL_CPT).setValue("退勤");
    
    // 行のグループ化を行います。
    sht1.getRange(1 + idRow1 + 1, 1, N_ROW_DATA - 1).shiftRowGroupDepth(1);
  }
  // 名前を更新
  sht1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
    
  // 今日の日付を取得
  var dateNow1 = Moment.moment();

  if(txAftChn1 == ""){
    let txNowChn1 = sht1.getRange(1 + idRow1, 1 + COL_STT).getValue();
    if(txNowChn1 == "退勤"){
      // 退勤からの退室は無視
    }
    else{
      // 退室の場合、退室時間を記録
      sht1.getRange(1 + idRow1, 1 + COL_OUT).setValue(dateNow1.format("YYYY-MM-DD HH:mm"));
      // 状態を消去
      sht1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    }
  }
  else{
    // 入室の場合、前回のチャンネルを確認
    let txNowChn1 = sht1.getRange(1 + idRow1, 1 + COL_STT).getValue();    
    if(txNowChn1 != ""){
      // 前回のチャンネルがある場合はただの部屋移動
      ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
    }
    else{
      // 前回が退室の場合は「サーバ移動」、「退室⇒入室」、「退勤⇒出勤」のいずれか      
      // 前回の退室からの時間を取得
      let txDate2 = sht1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
      let dateOut1 = Moment.moment(txDate2);
      // 退室時間(分)
      let ctOutTime1 = dateNow1.clone().diff(dateOut1, 'minutes');
      
      if(ctOutTime1 >= 240 && GetDayMnt(dateOut1) < 6 * 60 && 6 * 60 < GetDayMnt(dateNow1)){
        // 退室から4時間以上かつ6:00をまたいでいた場合は退勤扱い
        ChnSftExe(sht1, dateOut1, idRow1, "", "退室", txName1);
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else if(ctOutTime1 >= 2){
        // 退室から2分以上は退室扱い
        ChnSftExe(sht1, dateOut1, idRow1, "", "退室", txName1);
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else{
        // 退室から2分以内はサーバ移動
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
    }
  }
}

// ============================================================================
// 毎朝8時に退勤のチェックを行う
// ============================================================================
function Test5()
{
  let date1 = Moment.moment();
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("エラーログ");

  let txId1 = sht1.getRange(1, 3).getValue();
  let idMemRow1 = GetMemRow("Discord ID", txId1);
  AkashiDakoku(idMemRow1, "退勤", date1);
}

// ============================================================================
// AKASHIに打刻を行う
// ============================================================================
function AkashiDakoku(idMemRow1, txType1, date1)
{
  try{ 
    if(idMemRow1 < 0){
      return;
    }  
    
    var txTkn1 = GetMemVal(idMemRow1, "AKASHI Token");
    
    if(txTkn1 == ""){
      // AKASHIトークンがなければ打刻は行わない
      return;
    }
    
    var txKigyoId1 = "keycre7127";
    var url = "https://atnd.ak4.jp/api/cooperation/" + txKigyoId1 + "/stamps";
    let txTime1 = date1.format("YYYY/MM/DD HH:mm:ss");

  //  11 : 出勤
  //  12 : 退勤
  //  21 : 直行
  //  22 : 直帰
  //  31 : 休憩入
  //  32 : 休憩戻
    let tyType1 = "";
    
    if(txType1 == "退勤"){
      // 退勤
      tyType1 = "12";
    }
    else{
      // 出勤
      tyType1 = "11";
    }
    
    var jsPyl1 = {
      "token" : txTkn1,
      "type" : tyType1,
      "stampedAt" : txTime1
    };
    
    var jsPrm1 = {
      "method" : "post",
      "payload" : jsPyl1
    };
  
    // Slackに投稿する
    let res1 = UrlFetchApp.fetch(url, jsPrm1);
    
    const resInfo1 = JSON.parse(res1);
    if(resInfo1.success){
      let txSlkId1 = GetMemVal(idMemRow1, "Slack ID");
      if(txSlkId1 != ""){                               
        SendSlkMsg("AKASHIに" + txType1 + "打刻がされました。\n" + txTime1, "@" + txSlkId1);
      }
    }
  }
  catch(e){
    AddLog(e);
  }
}

// ============================================================================
// 毎朝8時に退勤のチェックを行う
// ============================================================================
function ChkKintai()
{
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("勤怠管理");
  
  // 現在の日付を取得
  var dateNow1 = Moment.moment();
  
  // 昨日の列を取得
  let idCol1 = GetDateCol(sht1, dateNow1.clone().subtract(1, "d"));
  if(idCol1 >= 0){
    
    let idRow1;
    for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1 += N_ROW_DATA){
      let txHis1 = sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).getValue();
      if(txHis1 != ""){
        let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).getValue();
        let txEnd1 = sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).getValue();
        if(txSta1 != "" && txEnd1 != ""){
          // 打刻がそろっている
        }
        else{
          // 打刻がそろっていない
          let txSlackId1 = sht1.getRange(1 + idRow1, 1 + COL_SLK_ID).getValue();
          if(txSlackId1 != ""){
            // SLACKIDがあればDMを送信
            let txMsg1 = "############こちらテストで送信していますので無視してください。##############\n"
            if(txSta1 == "" && txEnd1 == ""){
              txMsg1 = txMsg1 + "出勤と退勤がされていません。\n";
            }
            else if(txSta1 == ""){
              txMsg1 = txMsg1 + "出勤がされていません。\n";
            }
            else{
              txMsg1 = txMsg1 + "退勤がされていません。\n";
            }
            
            let txTimeSta1 = txHis1.slice(1, 6);
            let txDate2 = sht1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
            let dateOut1 = Moment.moment(txDate2);
            
            let txTimeEnd1 = dateOut1.format("HH:mm");
            
            txMsg1 = txMsg1 + "出勤:" + txTimeSta1 + "\n"
            txMsg1 = txMsg1 + "退勤:" + txTimeEnd1 + "\n"
            
            SendSlkMsg(txMsg1, "@"+txSlackId1);
          }
        }
      }
    }
  }
}

// ============================================================================
// 指定した時間の0時からの経過分数を取得
// ============================================================================
function GetDayMnt(date1)
{
  let ctDayMnt1 = date1.clone().startOf('day').diff(date1, 'minutes');
  return ctDayMnt1;
}

// ============================================================================
// 指定の日付の列番号を取得
// ============================================================================
function GetDateCol(sht1, date1){
  let txDate1 = date1.clone().add(-6, "h").format("MM/DD");
  // 打刻日付を取得
  let idCol1 = 0;
  // 日付行を右から検索
  for(idCol1 = sht1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
    // 日付行から日付を取得
    let txDate2 = sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      // 発見
      break;
    }
  }
  
  if(idCol1 < COL_DATA2){
    idCol1 = -1;
  }
  return(idCol1);
  
}

const ID_ROW_CPT = 1 - 1;
const ID_ROW_DATA = ID_ROW_CPT + 1;
const ID_COL_DATA = 3 - 1;

// ============================================================================
// ID管理シート行番号を取得
// ============================================================================
function GetMemRow(txCol1, txVal1)
{
  AddLog(txVal1);
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("ID管理");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    let txCol2 = sht1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){
      break;
    }
  }
  
  //AddLog(txVal1);
  if(idCol1 >= sht1.getLastColumn()){
    // 指定の列がありません。
    AddLog("列");
    return(-1);
  }
    
  let idRow1;
  for(idRow1 = ID_ROW_DATA; idRow1 < sht1.getLastRow(); idRow1++){
    let txVal2 = "'" + sht1.getRange(1 + idRow1, 1 + idCol1).getValue();
    
    //AddLog(txVal2 +" == "+txVal1);
    
    if(txVal2 == txVal1){
      break;
    }
  }
  
  if(idRow1 >= sht1.getLastRow()){
    // 指定のIDがありません。
    AddLog("行");
    return(-1);
  }
  
  return(idRow1);
}
    
// ============================================================================
// ID管理シートパラメータを取得
// ============================================================================
function GetMemVal(idMemRow1, txCol1)
{
  if(idMemRow1 < 0){
    // 指定の行がありません。
    return("");
  }

  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("ID管理");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    let txCol2 = sht1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){

      break;
    }
  }
  
  if(idCol1 >= sht1.getLastColumn()){
    // 指定の列がありません。
    return("");
  }
  
  let txVal1 = sht1.getRange(1 + idMemRow1, 1 + idCol1).getValue()
  
  return(txVal1);
}

// ============================================================================
// チャンネル移動実行
// ============================================================================
function ChnSftExe(sht1, date1, idRow1, txAftSvr1, txAftChn1, txName1){
  try{
    // 打刻日付は翌日の6:00を区切りとする
    let txDate1 = date1.clone().subtract(6, "h").format("MM/DD");
    // subtractするとなにやらdate1自体が変化するみたいなので、戻してやる
//    date1.add(6, "h")
    //let txDate1 = date1.format("MM/DD");
    let txTime1 = date1.format("HH:mm");
    
    // 打刻日付を取得
    let idCol1 = 0;
    // 日付行を右から検索
    for(idCol1 = sht1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
      // 日付行から日付を取得
      let txDate2 = sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
      if(txDate2 == txDate1){
        // 発見
        break;
      }
    }
    
    if(idCol1 < COL_DATA2){
      // 日付がなければ右端に追加
      idCol1 = sht1.getLastColumn();
      sht1.getRange(1 + ROW_DATE, 1 + idCol1).setValue("'" + txDate1);
    }
    
    let idCol2 = idCol1;
    if(GetDayMnt(date1) >= 10 * 60){
      // 10時以降なら前日の退勤がなかろうが本日の打刻として扱う
    }
    else{
      // 10時以前なら念のため、前日の退勤までチェック
      let idDay1;
      let ctDay1 = 2;
      for(idCol2 = idCol1, idDay1 = 0; idCol2 >= COL_DATA2, idDay1 < 2; idCol2--, idDay1++){
        // 今日の日付から直前の出社打刻ありの日付を検索
        let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
        if(txSta1 != ""){
          // 出社打刻あり      
          let txEnd1 = sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).getValue();
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
    
    let txId3 = "'" + sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    let idMemRow1 = GetMemRow("Discord ID", txId3);
    if(txAftChn1 == "退勤"){
      // 退勤
      // 退勤時は気にせず退勤打刻を行う
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).setValue(txTime1);
      //let txId3 = "'" + sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      //AddLog(txId3);
      //let idMemRow1 = GetMemRow("Discord ID", txId3);
      AkashiDakoku(idMemRow1, "退勤", date1);
    }
    else if(txAftChn1 == "退室"){
      // 退室
      // 退室は打刻情報は更新せず
    }
//    else{
    else if(txAftChn1 == "出社" || txAftChn1 == "テレワーク開始"){
      // 退勤以外は出勤扱い
      // 出勤
      let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
      if(txSta1 == ""){
        // 出勤打刻なしなら出勤打刻を行う
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).setValue(txTime1);
        //let txId3 = "'" + sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
        //let idMemRow1 = GetMemRow("Discord ID", txId3);
        AkashiDakoku(idMemRow1, "出勤", date1);
      }
    }
    
    //let txChn1 = "C01AG9H3GBF";
    let txChn1 = "C01805HS02F";
    
    if(txAftChn1 == "出社"){
      // 場所を更新
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        SendSlkMsg(txName1 + "がテレワークを終了し、出社しました。", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
      // Irucaステータスを「在席」に変更
      Iruca_WorkStartOffice(idMemRow1);
    }
    else if(txAftChn1 == "テレワーク開始"){
      // 場所を更新
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() == ""){
        SendSlkMsg(txName1 + "がテレワークを開始しました。", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("テレワーク中");
      }
      // Irucaステータスを「在席」テレワークに変更
      Iruca_WorkStartHome(idMemRow1);
    }
    else if(txAftChn1 == "退勤"){
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        SendSlkMsg(txName1 + "がテレワークを終了しました", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
      // Irucaステータスを「休暇」に変更
      Iruca_WorkEnd(idMemRow1);
    }
    
    let txHis1 = sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).getValue();
    
    if(txHis1 != ""){
      txHis1 = txHis1 + "⇒";
    }
    
    // 時間を記録
    txHis1 = txHis1 + "["+ txTime1 + "]";      
    let txNowSvr1 = sht1.getRange(1 + idRow1, 1 + COL_SVR).getValue(); 
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
    sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).setValue(txHis1);
    
    if(txAftSvr1 != ""){
      // 状態を更新
      sht1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
      // サーバを更新
      sht1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
      
      // Irucaメッセージを更新
      Iruca_SetMessage(idMemRow1, txAftChn1);
    }
  }
  catch(e){
    AddLog(e);
  }
}

function AddLog(log)
{
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheetErr1 = book1.getSheetByName("エラーログ");
  sheetErr1.getRange(1 + sheetErr1.getLastRow(), 1).setValue(log);
}

// ============================================================================
// 体温を更新
// ============================================================================
function UpdTaion(jsPrm1){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("体温管理");

  let txName1 = getUserName(jsPrm1.event.user);
  let txId1 = jsPrm1.event.user;

  if(txName1 == "KintaiKanri"){
    return;
  }
  
  // テキストから体温のみ抽出
  let txVal1 = jsPrm1.event.text;
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
    SendSlkMsg("「"+ txVal1 + "」"+ "\n入力が無効です。\n例:36.2、36、362", "@" + txId1);
    return;
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1++){
    // ID列からIDを取得
    let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
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
  for(idCol1 = COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = "'" + sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // 日付を更新
  sht1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // 名前を更新
  sht1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // 念のためIDを更新
  sht1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
  // 今日の体温を更新
  sht1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // デバッグ用
  //sht1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(jsPrm1);
}


// ============================================================================
// SlackのIDからSlackの表示名を取得
// ============================================================================
function getUserName(txSlkId1){
  const txTkn1 = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+txTkn1+"&user="+txSlkId1).getContentText();

  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("体温管理");

  const userInfo = JSON.parse(userData).user;  
  const userProf =userInfo.profile;
  const userName1 = userProf.display_name;
  const userName2 = userInfo.real_name;

  return userName1 ? userName1 : (userName2 ? userName2 : txSlkId1); 
}

// メッセージを送信
function SendMessage(){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("体温管理");

  // 今日の日付を取得
  var date = new Date();
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // 今日の日付列を取得
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    // 日付行から日付を取得
    let txDate2 = "'" + sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1++){
    // 名称列から名称を取得
    
    let txOndo2 = sht1.getRange(1 + idRow1, 1 + idCol1).getValue();
    if(txOndo2 == ""){
      // 温度が記入されていない
      let txName2 = sht1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
      let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      SendSlkMsg("体温を入力してください", "@"+txId2);
    }
  }
}

// ============================================================================
// Slackにメッセージを送信
// txChn1:@ユーザIDを指定するとDMを送信できます。
// ============================================================================
function SendSlkMsg(txMsg1, txChn1){
  var url = "https://slack.com/api/chat.postMessage";  
  var txTkn1 = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');

  var jsPyl1 = {
      "token" : txTkn1,
      "channel" : txChn1,
      "text" : txMsg1
    };
    
    var jsPrm1 = {
      "method" : "post",
      "payload" : jsPyl1
    };
    
    // Slackに投稿する
    let res1 = UrlFetchApp.fetch(url, jsPrm1);
}




// ============================================================================
// イルカ操作関数
// ============================================================================

// 指定ルームのメンバー情報取得
function Iruca_getMenbers(roomid){
  
  // メンバーリスト取得API
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members';
  
  // APIにリクエストしJSONデータを受け取る
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() >= 400) {
    // エラー
    Logger.log('Error: status = ' + response.getResponseCode());
    return null;
  }
  else{
    //Logger.log(response);
    return JSON.parse(response.getContentText());
  }
}

// 個人単位のメンバー情報取得
function Iruca_getMenber(roomid, memberid){
  
  if( roomid == "" ) return;
  if( memberid == "" ) return;
  
  // メンバー取得API
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members/' + memberid;
  
  // APIにリクエストしJSONデータを受け取る
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() >= 400) {
    // エラー
    Logger.log('Error: status = ' + response.getResponseCode());
    return null;
  }
  else{
    //Logger.log(response);
    return JSON.parse(response.getContentText());
  }  
}

// メンバー状態を変更する
function Iruca_setMemberStatus( roomid, id, status, msg ){
  
  if( roomid == "" ) return;
  if( id == "" ) return;
  
  // メンバー情報更新API
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members/' + id;
  
  var payload = {
    "status":status,
    "message": msg
  };
  var params = {
    "method": "put",
    "contentType" : "application/json", //データの形式を指定
    "payload" : JSON.stringify(payload),
     muteHttpExceptions : true
  };
  
  var response = UrlFetchApp.fetch(url,params);
  if (response.getResponseCode() >= 400) {
    // エラー
    Logger.log('Error: SetMemverStatus ErrSts = ' + response.getResponseCode());
  }
}

// 出社
function Iruca_WorkStartOffice( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",出社" );
    Iruca_setMemberStatus( room_id, member_id, "在席", "");
  }
}

// テレワーク
function Iruca_WorkStartHome( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",てれわーく");
    Iruca_setMemberStatus( room_id, member_id, "在席", "[テレワーク]" );
  }
}

// 退勤
function Iruca_WorkEnd( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",退勤");
    Iruca_setMemberStatus( room_id, member_id, "休暇", "" );
  }
}

// メッセージ（一言）設定
function Iruca_SetMessage( idRow, msg ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    
    if(( room_id != "" ) && (member_id != "") ){
      var member_inf = Iruca_getMenber( room_id, member_id );
      //WrtErrLog( idRow + "," + room_id+ "," + member_id + "," + member_inf.message);
      if( member_inf != null ){
        if( member_inf.message.includes("[テレワーク]") ){
          Iruca_setMemberStatus( room_id, member_id, member_inf.status , "[テレワーク]"+ msg);
        }
        else{
          Iruca_setMemberStatus( room_id, member_id, member_inf.status , msg );
        }
      }
    }
  }
}


// メンバーの状態をデバッグ表示
function Iruca_writeMenberList(members){
  /*
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("デバッグシート");

  if( members != null ){
    // メンバーの名前,状況を取得
    for( i=0; i<members.length; i++ ){
      if( members[i] != null ){
        sheet1.appendRow([members[i].id ,members[i].name , members[i].status, members[i].message ]);
      }
    }
  }
  */
}
