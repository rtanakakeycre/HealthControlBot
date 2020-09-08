/*
function doPost(e){
  var params = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(params.challenge);
}
*/

function doPost(e) {
  
  
  try{
    // �g�[�N������X���b�N�ւ̃����N���擾
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
    const slackApp = SlackApp.create(token);
    
    // �|�X�g�f�[�^����p�����[�^���擾
    const params = JSON.parse(e.postData.getDataAsString());
    writeLog(params);

    return ContentService.createTextOutput(params.challenge);
    
  }catch(err){
    writeLog(err);
  }
}

// �X�v���b�h�V�[�g�̍s�A����
const ROW_DATE = 0;
const ROW_DATA = 1;
const COL_NAME = 0;
const COL_ID = 1;
const COL_DATA = 2;

function writeLog(params){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");
  const channel = params.event.channel;

  Logger.log(params);
  
  let txName1 = getUserName(params.event.user);

  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    // ���̗񂩂疼�̂��擾
    let txName2 = sheet1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
    if(txName2 == txName1){
      break;
    }
  }
  
  // �����̓��t���擾
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "_" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // �����̓��t����擾
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // �e�L�X�g����̉��̂ݒ��o
  let txVal1 = params.event.text;
  let idChrSta1 = txVal1.indexOf("\n");
  txVal1 = txVal1.slice(idChrSta1 + 1);
  
  // ���t���X�V
  sheet1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // ���O���X�V
  sheet1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // �O�̂���ID���X�V
  sheet1.getRange(1 + idRow1, 1 + COL_ID).setValue(params.event.user);
  // �����̑̉����X�V
  sheet1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // �f�o�b�O�p
  //sheet1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(params);
}

// ���[�U�����擾
function getUserName(userId){
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+token+"&user="+userId).getContentText();

  const userName = JSON.parse(userData).user.real_name;
  Logger.log(userId,userName);

  return userName ? userName : userId; 
}

// ���b�Z�[�W�𑗐M
function SendMessage(){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�̉��Ǘ�");

  var url = "https://slack.com/api/chat.postMessage";
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
  var channel = "C019TVCKLEB";
  
  // �����̓��t���擾
  var date = new Date();
  //let txData1 = Utilities.formatDate( date, 'Asia/Tokyo', 'MMdd');  
  let txDate1 = "_" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // �����̓��t����擾
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = sheet1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
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
      
      // �ύX����̂́A���̕�������!
      var payload = {
        "token" : token,
        "channel" : channel,
        "text" : txName2 + "\n�̉�����͂��Ă��������B"
      };
      
      var params = {
        "method" : "post",
        "payload" : payload
      };
      
      // Slack�ɓ��e����
      let res1 = UrlFetchApp.fetch(url, params);
    }
  }

}