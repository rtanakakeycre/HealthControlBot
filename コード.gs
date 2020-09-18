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
      writeLog2(params.name, params.id, params.AftSvr, params.AftChn, params.BefSvr, params.BefChn);
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
const COL_OUT = COL_PLACE + 1;
const COL_CPT = COL_OUT + 1;
const COL_DATA2 = COL_CPT + 1;

function Test()
{
  //writeLog2("�c���Ǖ�", "730250456168792000", "", "", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��");
  writeLog2("�c���Ǖ�", "730250456168792000", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��2", "", "");
}

function Test2()
{
  writeLog2("�c���Ǖ�", "730250456168792000", "", "", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��");
  //writeLog2("�c���Ǖ�", "730250456168792000", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��2", "", "");
}

// �Αӏ����
function writeLog2(txName1, txId1, txAftSvr1, txAftChn1, txBefSvr1, txBefChn1){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�ΑӊǗ�");
   
  let flNew1 = true;
  // ID����Ή������o�[�̍s���擾
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
    
  // �����̓��t���擾
  var dateNow1 = Moment.moment();

  if(txAftChn1 == ""){
    // �ގ��̏ꍇ�A�ގ����Ԃ��L�^
    sheet1.getRange(1 + idRow1, 1 + COL_OUT).setValue(dateNow1.format("YYYY-MM-DD HH:mm"));
    // ��Ԃ�����
    sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    // �T�[�o������
    sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
  }
  else{
    // �����̏ꍇ�A�O��̑ގ�����̎��Ԃ��擾
    let txDate2 = sheet1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
    if(txDate2 == ""){
      // �����̕����ړ�
      ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
    }
    else{
      // �O�񂪑ގ��̏ꍇ�́u�T�[�o�ړ��v�A�u�ގ��˓����v�A�u�ދ΁ˏo�΁v�̂����ꂩ      
      let dateOut1 = Moment.moment(txDate2);
      // �ގ�����(��)
//      let ctOutTime1 = dateNow1.diff(dateOut1, 'minutes');
      //sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_SVR).setValue(ctOutTime1);
      ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      
//      if(ctOutTime1 >= 2){
//        // �ގ�����2���ȏ�͑ގ�����
//        ChnSftExe(sheet1, dateOut1, idRow1, "", "�ގ�", txName1);
//        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
//      }
//      else{
//        // �ގ�����2���ȓ��̓T�[�o�ړ�
//        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
//      }
//      // �ގ����Ԃ��폜
//      sheet1.getRange(1 + idRow1, 1 + COL_OUT).setValue("");
    }
  }
}

// ============================================================================
// �`�����l���ړ����s
// ============================================================================
function ChnSftExe(sheet1, date1, idRow1, txAftSvr1, txAftChn1, txName1){
  try{
    // �ō����t�͗�����6:00����؂�Ƃ���
    let txDate1 = date1.subtract(6, "h").format("MM/DD");
    let txTime1 = date1.format("HH:mm");
    
    // �ō����t���擾
    let idCol1 = 0;
    // ���t�s���E���猟��
    for(idCol1 = sheet1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
      // ���t�s������t���擾
      let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
      if(txDate2 == txDate1){
        // ����
        break;
      }
    }
    
    if(idCol1 < COL_DATA2){
      // ���t���Ȃ���ΉE�[�ɒǉ�
      idCol1 = sheet1.getLastColumn();
      sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue("'" + txDate1);
    }
    
    let idCol2 = idCol1;
    let ctDayMnt1 = date1.startOf('day').diff(date1, 'minutes');
    if(ctDayMnt1 >= 10 * 60){
      // 10���ȍ~�Ȃ�O���̑ދ΂��Ȃ��낤���{���̑ō��Ƃ��Ĉ���
    }
    else{
      // 10���ȑO�Ȃ�O�̂��߁A�O���̑ދ΂܂Ń`�F�b�N
      let idDay1;
      let ctDay1 = 2;
      for(idCol2 = idCol1, idDay1 = 0; idCol2 >= COL_DATA2, idDay1 < 2; idCol2--, idDay1++){
        // �����̓��t���璼�O�̏o�Бō�����̓��t������
        let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
        if(txSta1 != ""){
          // �o�Бō�����      
          let txEnd1 = sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).getValue();
          if(txEnd1 != ""){
            // �o�Бō����ދΑō�������̏ꍇ�͌��݂̓��t��ō����t�Ƃ���
            idCol2 = idCol1;
          }
          else{
            // �o�Бō�����őދΑō��Ȃ��̏ꍇ�͂��̓��t��ō����t�Ƃ���
          }
          break;
        }    
      }
      
      if(idDay1 < COL_DATA2 || idDay1 >= 2){
        // �o�Бō���������Ȃ���Ό��݂̓��t��ō����t�Ƃ���
        idCol2 = idCol1;
      }
    }    
    
    if(txAftChn1 == "�ދ�"){
      // �ދ�
      // �ދΎ��͋C�ɂ����ދΑō����s��
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).setValue(txTime1);
    }
    else{
    //else if(txAftChn1 == "�o��" || txAftChn1 == "�e�����[�N�J�n"){
      // �ދΈȊO�͏o�Έ���
      // �o��
      let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
      if(txSta1 == ""){
        // �o�Αō��Ȃ��Ȃ�o�Αō����s��
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).setValue(txTime1);
      }
    }
    
    //let txChn1 = "C01AG9H3GBF";
    let txChn1 = "C01805HS02F";
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
    
    let txHis1 = sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).getValue();
    
    if(txHis1 != ""){
      txHis1 = txHis1 + "��";
    }
    
    // ���Ԃ��L�^
    txHis1 = txHis1 + "["+ txTime1 + "]";      
    let txNowSvr1 = sheet1.getRange(1 + idRow1, 1 + COL_SVR).getValue(); 
    if(txAftSvr1 == "" || txAftSvr1 == "KEY_�ΑӊǗ�"){
      // ����T�[�o�͖���
    }
    else{
      if(txAftSvr1 != txNowSvr1){
        // �T�[�o�ړ����������Ă���ΒǋL
        txHis1 = txHis1 + "(" + txAftSvr1 + ")";
      }
    }
    
    txHis1 = txHis1 + txAftChn1;
    
    // �������X�V
    sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).setValue(txHis1);
    
    if(txAftSvr1 != ""){
      // ��Ԃ��X�V
      sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
      // �T�[�o���X�V
      sheet1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
    }
  }
  catch(e){
    let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
    let sheetErr1 = spreadSheet1.getSheetByName("�G���[���O");
    sheetErr1.getRange(1 + sheetErr1.getLastColumn(), 1).setValue(e);
  }
}

// ============================================================================
// �̉����X�V
// ============================================================================
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
