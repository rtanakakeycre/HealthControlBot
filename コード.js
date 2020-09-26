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
  txId1 = "_" + txId1;
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
    let txNowChn1 = sheet1.getRange(1 + idRow1, 1 + COL_STT).getValue();
    if(txNowChn1 == "�ދ�"){
      // �ދ΂���̑ގ��͖���
    }
    else{
      // �ގ��̏ꍇ�A�ގ����Ԃ��L�^
      sheet1.getRange(1 + idRow1, 1 + COL_OUT).setValue(dateNow1.format("YYYY-MM-DD HH:mm"));
      // ��Ԃ�����
      sheet1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    }
  }
  else{
    // �����̏ꍇ�A�O��̃`�����l�����m�F
    let txNowChn1 = sheet1.getRange(1 + idRow1, 1 + COL_STT).getValue();    
    if(txNowChn1 != ""){
      // �O��̃`�����l��������ꍇ�͂����̕����ړ�
      ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
    }
    else{
      // �O�񂪑ގ��̏ꍇ�́u�T�[�o�ړ��v�A�u�ގ��˓����v�A�u�ދ΁ˏo�΁v�̂����ꂩ      
      // �O��̑ގ�����̎��Ԃ��擾
      let txDate2 = sheet1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
      let dateOut1 = Moment.moment(txDate2);
      // �ގ�����(��)
      let ctOutTime1 = dateNow1.clone().diff(dateOut1, 'minutes');
      
      if(ctOutTime1 >= 240 && GetDayMnt(dateOut1) < 6 * 60 && 6 * 60 < GetDayMnt(dateNow1)){
        // �ގ�����4���Ԉȏォ��6:00���܂����ł����ꍇ�͑ދΈ���
        ChnSftExe(sheet1, dateOut1, idRow1, "", "�ގ�", txName1);
        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else if(ctOutTime1 >= 2){
        // �ގ�����2���ȏ�͑ގ�����
        ChnSftExe(sheet1, dateOut1, idRow1, "", "�ގ�", txName1);
        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else{
        // �ގ�����2���ȓ��̓T�[�o�ړ�
        ChnSftExe(sheet1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
    }
  }
}

// ============================================================================
// ����8���ɑދ΂̃`�F�b�N���s��
// ============================================================================
function Test5()
{
  let date1 = Moment.moment();
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�G���[���O");

  let txId1 = sheet1.getRange(1, 3).getValue();
  let idMemRow1 = GetMemRow("Discord ID", txId1);
  AkashiDakoku(idMemRow1, "�ދ�", date1);
}

// ============================================================================
// ����8���ɑދ΂̃`�F�b�N���s��
// ============================================================================
function AkashiDakoku(idMemRow1, txType1, date1)
{
  WrtErrLog(idMemRow1);

  if(idMemRow1 < 0){
    return;
  }  
  
  var token = PropertiesService.getScriptProperties().getProperty('AKASHI_ACCESS_TOKEN');
  var txToken1 = GetMemVal(idMemRow1, "AKASHI Token");
  
  WrtErrLog(txToken1);
  
  if(txToken1 == ""){
    // AKASHI�g�[�N�����Ȃ���Αō��͍s��Ȃ�
    return;
  }
  
  var txKigyoId1 = "keycre7127";
 // var url = "https://atnd.ak4.jp/api/cooperation/" + txKigyoId1 + "/stamps?token=" + token + "&start_date=20200918000000&end_date=20200919000000"
  //var url = "https://atnd.ak4.jp/api/cooperation/" + txKigyoId1 + "/staffs?token=" + token + "&page=2";
  //var url = "https://atnd.ak4.jp/api/cooperation/" + txKigyoId1 + "/stamps?token=" + token + "&type=12&stampedAt=2020/09/23 18:30:00";
  var url = "https://atnd.ak4.jp/api/cooperation/" + txKigyoId1 + "/stamps";
  let txTime1 = date1.format("YYYY/MM/DD HH:mm:ss");

//  11 : �o��
//  12 : �ދ�
//  21 : ���s
//  22 : ���A
//  31 : �x�e��
//  32 : �x�e��
  let tyType1 = "";
  
  if(txType1 == "�ދ�"){
    // �ދ�
    tyType1 = "12";
  }
  else{
    // �o��
    tyType1 = "11";
  }
  
  var payload = {
    "token" : txToken1,
    "type" : tyType1,
    "stampedAt" : txTime1
  };
  
  var params = {
    "method" : "post",
    "payload" : payload
  };
  
  try{ 
    // Slack�ɓ��e����
    let res1 = UrlFetchApp.fetch(url, params);
    
    const resInfo1 = JSON.parse(res1);
    if(resInfo1.success){
      let txSlkId1 = GetMemVal(idMemRow1, "Slack ID");
      if(txSlkId1 != ""){                               
        PostMessage("AKASHI��" + txType1 + "�ō�������܂����B\n" + txTime1, "@" + txSlkId1);
      }
    }
  }
  catch(e){
    WrtErrLog(e);
  }
}

// ============================================================================
// ����8���ɑދ΂̃`�F�b�N���s��
// ============================================================================
function KinChk()
{
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�ΑӊǗ�");
  
  var dateNow1 = Moment.moment();
  
  // ����̗���擾
  let idCol1 = GetDateCol(sheet1, dateNow1.clone().subtract(1, "d"));
  if(idCol1 >= 0){
    
    let idRow1;
    for(idRow1 = ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1 += N_ROW_DATA){
      let txHis1 = sheet1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).getValue();
      if(txHis1 != ""){
        let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).getValue();
        let txEnd1 = sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).getValue();
        if(txSta1 != "" && txEnd1 != ""){
          // �ō���������Ă���
        }
        else{
          // �ō���������Ă��Ȃ�
          let txSlackId1 = sheet1.getRange(1 + idRow1, 1 + COL_SLK_ID).getValue();
          if(txSlackId1 != ""){
            // SLACKID�������DM�𑗐M
            let txMsg1 = "############������e�X�g�ő��M���Ă��܂��̂Ŗ������Ă��������B##############\n"
            if(txSta1 == "" && txEnd1 == ""){
              txMsg1 = txMsg1 + "�o�΂Ƒދ΂�����Ă��܂���B\n";
            }
            else if(txSta1 == ""){
              txMsg1 = txMsg1 + "�o�΂�����Ă��܂���B\n";
            }
            else{
              txMsg1 = txMsg1 + "�ދ΂�����Ă��܂���B\n";
            }
            
            let txTimeSta1 = txHis1.slice(1, 6);
            let txDate2 = sheet1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
            let dateOut1 = Moment.moment(txDate2);
            
            let txTimeEnd1 = dateOut1.format("HH:mm");
            
            txMsg1 = txMsg1 + "�o��:" + txTimeSta1 + "\n"
            txMsg1 = txMsg1 + "�ދ�:" + txTimeEnd1 + "\n"
            
            PostMessage(txMsg1, "@"+txSlackId1);
          }
        }
      }
    }
  }
}

// ============================================================================
// �w�肵�����Ԃ�0������̌o�ߕ������擾
// ============================================================================
function GetDayMnt(date1)
{
  let ctDayMnt1 = date1.clone().startOf('day').diff(date1, 'minutes');
  return ctDayMnt1;
}

// ============================================================================
// �w��̓��t�̗�ԍ����擾
// ============================================================================
function GetDateCol(sheet1, date1){
  let txDate1 = date1.clone().add(-6, "h").format("MM/DD");
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
    idCol1 = -1;
  }
  return(idCol1);
  
}

const ID_ROW_CPT = 1 - 1;
const ID_ROW_DATA = ID_ROW_CPT + 1;
const ID_COL_DATA = 3 - 1;

// ============================================================================
// ID�Ǘ��V�[�g�s�ԍ����擾
// ============================================================================
function GetMemRow(txCol1, txVal1)
{
  WrtErrLog(txVal1);
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("ID�Ǘ�");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    let txCol2 = sheet1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){
      break;
    }
  }
  
  //WrtErrLog(txVal1);
  if(idCol1 >= sheet1.getLastColumn()){
    // �w��̗񂪂���܂���B
    WrtErrLog("��");
    return(-1);
  }
    
  let idRow1;
  for(idRow1 = ID_ROW_DATA; idRow1 < sheet1.getLastRow(); idRow1++){
    let txVal2 = "'" + sheet1.getRange(1 + idRow1, 1 + idCol1).getValue();
    
    //WrtErrLog(txVal2 +" == "+txVal1);
    
    if(txVal2 == txVal1){
      break;
    }
  }
  
  if(idRow1 >= sheet1.getLastRow()){
    // �w���ID������܂���B
    WrtErrLog("�s");
    return(-1);
  }
  
  return(idRow1);
}
    
// ============================================================================
// ID�Ǘ��V�[�g�p�����[�^���擾
// ============================================================================
function GetMemVal(idMemRow1, txCol1)
{
  if(idMemRow1 < 0){
    // �w��̍s������܂���B
    return("");
  }

  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("ID�Ǘ�");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sheet1.getLastColumn(); idCol1++){
    let txCol2 = sheet1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){

      break;
    }
  }
  
  if(idCol1 >= sheet1.getLastColumn()){
    // �w��̗񂪂���܂���B
    return("");
  }
  
  let txVal1 = sheet1.getRange(1 + idMemRow1, 1 + idCol1).getValue()
  
  return(txVal1);
}

// ============================================================================
// �`�����l���ړ����s
// ============================================================================
function ChnSftExe(sheet1, date1, idRow1, txAftSvr1, txAftChn1, txName1){
  try{
    // �ō����t�͗�����6:00����؂�Ƃ���
    let txDate1 = date1.clone().subtract(6, "h").format("MM/DD");
    // subtract����ƂȂɂ��date1���̂��ω�����݂����Ȃ̂ŁA�߂��Ă��
//    date1.add(6, "h")
    //let txDate1 = date1.format("MM/DD");
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
    if(GetDayMnt(date1) >= 10 * 60){
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
    
    let txId3 = "'" + sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    let idMemRow1 = GetMemRow("Discord ID", txId3);
    if(txAftChn1 == "�ދ�"){
      // �ދ�
      // �ދΎ��͋C�ɂ����ދΑō����s��
      sheet1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).setValue(txTime1);
      //let txId3 = "'" + sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      //WrtErrLog(txId3);
      //let idMemRow1 = GetMemRow("Discord ID", txId3);
      AkashiDakoku(idMemRow1, "�ދ�", date1);
    }
    else if(txAftChn1 == "�ގ�"){
      // �ގ�
      // �ގ��͑ō����͍X�V����
    }
//    else{
    else if(txAftChn1 == "�o��" || txAftChn1 == "�e�����[�N�J�n"){
      // �ދΈȊO�͏o�Έ���
      // �o��
      let txSta1 = sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
      if(txSta1 == ""){
        // �o�Αō��Ȃ��Ȃ�o�Αō����s��
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).setValue(txTime1);
        //let txId3 = "'" + sheet1.getRange(1 + idRow1, 1 + COL_ID).getValue();
        //let idMemRow1 = GetMemRow("Discord ID", txId3);
        AkashiDakoku(idMemRow1, "�o��", date1);
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
      // Iruca�X�e�[�^�X���u�ݐȁv�ɕύX
      Iruca_WorkStartOffice(idMemRow1);
    }
    else if(txAftChn1 == "�e�����[�N�J�n"){
      // �ꏊ���X�V
      if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() == ""){
        PostMessage(txName1 + "���e�����[�N���J�n���܂����B", txChn1);
        sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("�e�����[�N��");
      }
      // Iruca�X�e�[�^�X���u�ݐȁv�e�����[�N�ɕύX
      Iruca_WorkStartHome(idMemRow1);
    }
    else if(txAftChn1 == "�ދ�"){
      if(sheet1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        PostMessage(txName1 + "���e�����[�N���I�����܂���", txChn1);
        sheet1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
      // Iruca�X�e�[�^�X���u�x�Ɂv�ɕύX
      Iruca_WorkEnd(idMemRow1);
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
      
      // Iruca���b�Z�[�W���X�V
      Iruca_SetMessage(idMemRow1, txAftChn1);
    }
  }
  catch(e){
    WrtErrLog(e);
  }
}

function WrtErrLog(log)
{
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheetErr1 = spreadSheet1.getSheetByName("�G���[���O");
  sheetErr1.getRange(1 + sheetErr1.getLastRow(), 1).setValue(log);
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




// ============================================================================
// �C���J����֐�
// ============================================================================

// �w�胋�[���̃����o�[���擾
function Iruca_getMenbers(roomid){
  
  // �����o�[���X�g�擾API
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members';
  
  // API�Ƀ��N�G�X�g��JSON�f�[�^���󂯎��
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() >= 400) {
    // �G���[
    Logger.log('Error: status = ' + response.getResponseCode());
    return null;
  }
  else{
    //Logger.log(response);
    return JSON.parse(response.getContentText());
  }
}

// �l�P�ʂ̃����o�[���擾
function Iruca_getMenber(roomid, memberid){
  
  if( roomid == "" ) return;
  if( memberid == "" ) return;
  
  // �����o�[�擾API
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members/' + memberid;
  
  // API�Ƀ��N�G�X�g��JSON�f�[�^���󂯎��
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() >= 400) {
    // �G���[
    Logger.log('Error: status = ' + response.getResponseCode());
    return null;
  }
  else{
    //Logger.log(response);
    return JSON.parse(response.getContentText());
  }  
}

// �����o�[��Ԃ�ύX����
function Iruca_setMemberStatus( roomid, id, status, msg ){
  
  if( roomid == "" ) return;
  if( id == "" ) return;
  
  // �����o�[���X�VAPI
  var url = 'https://iruca.co/api/rooms/' + roomid + '/members/' + id;
  
  var payload = {
    "status":status,
    "message": msg
  };
  var params = {
    "method": "put",
    "contentType" : "application/json", //�f�[�^�̌`�����w��
    "payload" : JSON.stringify(payload),
     muteHttpExceptions : true
  };
  
  var response = UrlFetchApp.fetch(url,params);
  if (response.getResponseCode() >= 400) {
    // �G���[
    Logger.log('Error: SetMemverStatus ErrSts = ' + response.getResponseCode());
  }
}

// �o��
function Iruca_WorkStartOffice( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",�o��" );
    Iruca_setMemberStatus( room_id, member_id, "�ݐ�", "");
  }
}

// �e�����[�N
function Iruca_WorkStartHome( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",�Ă��[��");
    Iruca_setMemberStatus( room_id, member_id, "�ݐ�", "[�e�����[�N]" );
  }
}

// �ދ�
function Iruca_WorkEnd( idRow ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    //WrtErrLog( idRow + "," + room_id+ "," + member_id + ",�ދ�");
    Iruca_setMemberStatus( room_id, member_id, "�x��", "" );
  }
}

// ���b�Z�[�W�i�ꌾ�j�ݒ�
function Iruca_SetMessage( idRow, msg ){
  if( idRow > 0 ){
    let room_id = GetMemVal(idRow, "iruca ROOM ID");
    let member_id = GetMemVal(idRow, "iruca Member ID");
    
    if(( room_id != "" ) && (member_id != "") ){
      var member_inf = Iruca_getMenber( room_id, member_id );
      //WrtErrLog( idRow + "," + room_id+ "," + member_id + "," + member_inf.message);
      if( member_inf != null ){
        if( member_inf.message.includes("[�e�����[�N]") ){
          Iruca_setMemberStatus( room_id, member_id, member_inf.status , "[�e�����[�N]"+ msg);
        }
        else{
          Iruca_setMemberStatus( room_id, member_id, member_inf.status , msg );
        }
      }
    }
  }
}


// �����o�[�̏�Ԃ��f�o�b�O�\��
function Iruca_writeMenberList(members){
  /*
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  let sheet1 = spreadSheet1.getSheetByName("�f�o�b�O�V�[�g");

  if( members != null ){
    // �����o�[�̖��O,�󋵂��擾
    for( i=0; i<members.length; i++ ){
      if( members[i] != null ){
        sheet1.appendRow([members[i].id ,members[i].name , members[i].status, members[i].message ]);
      }
    }
  }
  */
}
