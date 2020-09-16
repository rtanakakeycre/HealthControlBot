/*
function doPost(e){
  var params = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(params.challenge);
}
*/

// https://script.google.com/macros/s/AKfycby1lxGnLphhMjy-WrLsKglK5ZgEwBUFHVA_VTLUwD_QBIYrlOU/exec
function doPost(e) {
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");
  let sheet2 = spreadSheet1.getSheetByName("�ΑӊǗ�");
    
  try{
    // �g�[�N������X���b�N�ւ̃����N���擾
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
    const slackApp = SlackApp.create(token);
    
    // �|�X�g�f�[�^����p�����[�^���擾
    const params = JSON.parse(e.postData.getDataAsString());
    //sheet2.getRange(1, 1).setValue(params);
    
    if(params.type == "SheetWrite"){
      // �V�[�g�������ݗv��
      writeLog2(params);
    }
    else{
      // slack����̏��
      writeLog(params);
    }    

    return ContentService.createTextOutput(params.challenge);
    
  }catch(err){
    writeLog(err);
  }
}

// �X�v���b�h�V�[�g�̍s�A����
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

function writeLog2(params){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�ΑӊǗ�");
   
  let flNew1 = true;
  let txName1 = params.name;
  let txId1 = params.id;
  let txAftSvr1 = params.AftSvr;
  let txAftChn1 = params.AftChn;
  let txBefSvr1 = params.BefSvr;
  let txBefChn1 = params.BefChn;
  
  if(txAftChn1 != ""){

    let idRow1;
    for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1 += N_ROW_DATA){
      // ID�񂩂�ID���擾
      let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      if(txId2 == txId1){
        // ID�����łɑ���
        flNew1 = false;
        break;
      }
    }
    
    if(flNew1){
      // ID���L��
      sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
      
      // ���o�����L��
      sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + COL_CPT).setValue("�Αӗ���");
      sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_CPT).setValue("�o��");
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + COL_CPT).setValue("�ދ�");
      
      // �s�̃O���[�v�����s���܂��B
      sheet1.getRange(1 + idRow1 + 1, 1, N_ROW_DATA - 1).shiftRowGroupDepth(1);
    }
    // ���O���X�V
    sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
    
    // �T�[�o���X�V
    sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
    // ��Ԃ��X�V
    sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    
    let txChn1 = "C01AG9H3GBF";
    if(txAftChn1 == "�o��"){
      // �ꏊ���X�V
      if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        PostMessage(txName1 + "���e�����[�N���I�����A�o�Ђ��܂����B", txChn1);
        sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
    }
    else if(txAftChn1 == "�e�����[�N�J�n"){
      // �ꏊ���X�V
      if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() == ""){
        PostMessage(txName1 + "���e�����[�N���J�n���܂����B", txChn1);
        sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("�e�����[�N��");
      }
    }
    else if(txAftChn1 == "�ދ�"){
      if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        PostMessage(txName1 + "���e�����[�N���I�����܂���", txChn1);
        sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
    }
    
    if(txAftChn1 == "�o��" || txAftChn1 == "�ދ�"){
      // �ꏊ���X�V
      sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
    }
    else if(txAftChn1 == "�e�����[�N�J�n"){
      // �ꏊ���X�V
      sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("�e�����[�N��");
    }
    
    // �����̓��t���擾
    var date = new Date();
    //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
    let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
    let txTime1 = Utilities.formatDate( date, 'Asia/Tokyo', 'HH:mm');
    
    // �����̓��t����擾
    let idCol1 = COL_DATA2;
    for(idCol1 = COL_DATA2; idCol1 < sheet1.getLastColumn(); idCol1++){
      // ���t�s������t���擾
      let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
      if(txDate2 == txDate1){
        break;
      }
    }
    
    // ���t���X�V
    sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
    
    let txHis1 = sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).getValue();
    let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).getValue();
    let txEnd1 = sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).getValue();
    if(txAftChn1 == "�o��" || txAftChn1 == "�e�����[�N�J�n"){
      if(txSta1 == ""){
        // �o�Ђ��݂łȂ���΋L�^
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).setValue(txTime1);
      }
    }
    else if(txAftChn1 == "�ދ�"){
      if(txSta1 == ""){
        // �o�Ѝς݂łȂ���Γ����ׂ��ł���̂ŁA����̑ދ΂Ƃ���B
        if(idCol1 - 1 >= COL_DATA2){
          sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1 - 1).setValue(txTime1);
        }
      }
      else{
        // �ދ΍ς݂ł��L�^
        sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).setValue(txTime1);
      }
    }

    if(txHis1 != ""){
      txHis1 = txHis1 + "��";
    }
    txHis1 = txHis1 + "["+ txTime1 + "]" + txAftChn1;
    sheet1.getRange(1 + idRow1, 1 + idCol1).setValue(txHis1);
  }
  
}


function writeLog(params){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");
  const channel = params.event.channel;

  Logger.log(params);
    
  let txName1 = getUserName(params.event.user);
  let txId1 = params.event.user;

  if(txName1 == "KintaiKanri"){
    return;
  }
  
  // �e�L�X�g����̉��̂ݒ��o
  let txVal1 = params.event.text;
  let idChrSta1 = txVal1.indexOf("\n");
  txVal1 = txVal1.slice(idChrSta1 + 1);
  
  if(txVal1.match(/^\d\d\d$/g)) {
    // (365)
    txVal1 = txVal1.substr(0,2) + "." + txVal1.substr(2,1);
  }   
  else if(txVal1.match(/^\d\d(\.\d)?$/g)) {
    // (36.0�A36)
  } 
  else{
    // �̉��ł͂Ȃ�(36�A36.5�A365)
    PostMessage("�u"+ txVal1 + "�v"+ "\n���͂������ł��B\n��:36.2�A36�A362", "@" + txId1);
    return;
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // ID�񂩂�ID���擾
    let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    if(txId2 == txId1){
      break;
    }
  }
  
  // �����̓��t���擾
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // �����̓��t����擾
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // ���t���X�V
  sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // ���O���X�V
  sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // �O�̂���ID���X�V
  sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
  // �����̑̉����X�V
  sheet1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // �f�o�b�O�p
  //sheet1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(params);
}




// ���[�U�����擾
function getUserName(userId){
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+token+"&user="+userId).getContentText();

  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");

  const userInfo = JSON.parse(userData).user;  
  const userProf =userInfo.profile;
  const userName1 = userProf.display_name;
  const userName2 = userInfo.real_name;

  return userName1 ? userName1 : (userName2 ? userName2 : userId); 
}

// ���b�Z�[�W�𑗐M
function SendMessage(){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");

  // �����̓��t���擾
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // �����̓��t����擾
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = "'" + sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // ���̗񂩂疼�̂��擾
    
    let txOndo2 = sheet1.getRange(1 + idRow1, 1 + idCol1).getValue();
    if(txOndo2 == ""){
      // ���x���L������Ă��Ȃ�
      let txName2 = sheet1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
      let txId2 = sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      PostMessage("�̉�����͂��Ă�������", "@"+txId2);
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
    
    // Slack�ɓ��e����
    let res1 = UrlFetchApp.fetch(url, params);
}
