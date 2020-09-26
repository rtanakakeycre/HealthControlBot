/*
function doPost(e){
  var jsPrm1 = JSON.parse(e.postData.getDataAsString());
  return ContentService.createTextOutput(jsPrm1.challenge);
}
*/


// �G���g���|�C���g
// https://script.google.com/macros/s/AKfycby1lxGnLphhMjy-WrLsKglK5ZgEwBUFHVA_VTLUwD_QBIYrlOU/exec

// ============================================================================
// POST�̃n���h��
// ============================================================================
function doPost(e) {
   
  try{    
    // �|�X�g�f�[�^����p�����[�^���擾
    const jsPrm1 = JSON.parse(e.postData.getDataAsString());
    
    if(jsPrm1.type == "SheetWrite"){
      // �V�[�g�������ݗv��
      UpdKintai(jsPrm1.name, jsPrm1.id, jsPrm1.AftSvr, jsPrm1.AftChn, jsPrm1.BefSvr, jsPrm1.BefChn);
    }
    else{
      // slack����̏��
      UpdTaion(jsPrm1);
    }    

    return ContentService.createTextOutput(jsPrm1.challenge);
    
  }catch(err){
    UpdTaion(err);
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
  //UpdKintai("�c���Ǖ�", "730250456168792000", "", "", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��");
  UpdKintai("�c���Ǖ�", "730250456168792000", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��2", "", "");
}

function Test2()
{
  UpdKintai("�c���Ǖ�", "730250456168792000", "", "", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��");
  //UpdKintai("�c���Ǖ�", "730250456168792000", "�e�X�g�T�[�o�[", "�e�X�g�`�����l��2", "", "");
}

// ============================================================================
// �Αӏ��̉�͂ƍX�V
// ============================================================================
function UpdKintai(txName1, txDscId1, txAftSvr1, txAftChn1, txBefSvr1, txBefChn1){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�ΑӊǗ�");

  // DiscordID�́u_�v�t���ŊǗ����Ă��邽�߁u_�v��t��
  let txId1 = "_" + txDscId1;
   
  let flNew1 = true;
  // ID����Ή������o�[�̍s���擾
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1 += N_ROW_DATA){
    // ID�񂩂�ID���擾
    let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
    if(txId2 == txId1){
      // ID�����łɑ���
      flNew1 = false;
      break;
    }
  }
  
  if(flNew1){
    // ID���L��
    sht1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
    
    // ���o�����L��
    sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + COL_CPT).setValue("�Αӗ���");
    sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + COL_CPT).setValue("�o��");
    sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + COL_CPT).setValue("�ދ�");
    
    // �s�̃O���[�v�����s���܂��B
    sht1.getRange(1 + idRow1 + 1, 1, N_ROW_DATA - 1).shiftRowGroupDepth(1);
  }
  // ���O���X�V
  sht1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
    
  // �����̓��t���擾
  var dateNow1 = Moment.moment();

  if(txAftChn1 == ""){
    let txNowChn1 = sht1.getRange(1 + idRow1, 1 + COL_STT).getValue();
    if(txNowChn1 == "�ދ�"){
      // �ދ΂���̑ގ��͖���
    }
    else{
      // �ގ��̏ꍇ�A�ގ����Ԃ��L�^
      sht1.getRange(1 + idRow1, 1 + COL_OUT).setValue(dateNow1.format("YYYY-MM-DD HH:mm"));
      // ��Ԃ�����
      sht1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
    }
  }
  else{
    // �����̏ꍇ�A�O��̃`�����l�����m�F
    let txNowChn1 = sht1.getRange(1 + idRow1, 1 + COL_STT).getValue();    
    if(txNowChn1 != ""){
      // �O��̃`�����l��������ꍇ�͂����̕����ړ�
      ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
    }
    else{
      // �O�񂪑ގ��̏ꍇ�́u�T�[�o�ړ��v�A�u�ގ��˓����v�A�u�ދ΁ˏo�΁v�̂����ꂩ      
      // �O��̑ގ�����̎��Ԃ��擾
      let txDate2 = sht1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
      let dateOut1 = Moment.moment(txDate2);
      // �ގ�����(��)
      let ctOutTime1 = dateNow1.clone().diff(dateOut1, 'minutes');
      
      if(ctOutTime1 >= 240 && GetDayMnt(dateOut1) < 6 * 60 && 6 * 60 < GetDayMnt(dateNow1)){
        // �ގ�����4���Ԉȏォ��6:00���܂����ł����ꍇ�͑ދΈ���
        ChnSftExe(sht1, dateOut1, idRow1, "", "�ގ�", txName1);
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else if(ctOutTime1 >= 2){
        // �ގ�����2���ȏ�͑ގ�����
        ChnSftExe(sht1, dateOut1, idRow1, "", "�ގ�", txName1);
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
      }
      else{
        // �ގ�����2���ȓ��̓T�[�o�ړ�
        ChnSftExe(sht1, dateNow1, idRow1, txAftSvr1, txAftChn1, txName1);
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
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�G���[���O");

  let txId1 = sht1.getRange(1, 3).getValue();
  let idMemRow1 = GetMemRow("Discord ID", txId1);
  AkashiDakoku(idMemRow1, "�ދ�", date1);
}

// ============================================================================
// AKASHI�ɑō����s��
// ============================================================================
function AkashiDakoku(idMemRow1, txType1, date1)
{
  try{ 
    if(idMemRow1 < 0){
      return;
    }  
    
    var txTkn1 = GetMemVal(idMemRow1, "AKASHI Token");
    
    if(txTkn1 == ""){
      // AKASHI�g�[�N�����Ȃ���Αō��͍s��Ȃ�
      return;
    }
    
    var txKigyoId1 = "keycre7127";
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
    
    var jsPyl1 = {
      "token" : txTkn1,
      "type" : tyType1,
      "stampedAt" : txTime1
    };
    
    var jsPrm1 = {
      "method" : "post",
      "payload" : jsPyl1
    };
  
    // Slack�ɓ��e����
    let res1 = UrlFetchApp.fetch(url, jsPrm1);
    
    const resInfo1 = JSON.parse(res1);
    if(resInfo1.success){
      let txSlkId1 = GetMemVal(idMemRow1, "Slack ID");
      if(txSlkId1 != ""){                               
        SendSlkMsg("AKASHI��" + txType1 + "�ō�������܂����B\n" + txTime1, "@" + txSlkId1);
      }
    }
  }
  catch(e){
    AddLog(e);
  }
}

// ============================================================================
// ����8���ɑދ΂̃`�F�b�N���s��
// ============================================================================
function ChkKintai()
{
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�ΑӊǗ�");
  
  // ���݂̓��t���擾
  var dateNow1 = Moment.moment();
  
  // ����̗���擾
  let idCol1 = GetDateCol(sht1, dateNow1.clone().subtract(1, "d"));
  if(idCol1 >= 0){
    
    let idRow1;
    for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1 += N_ROW_DATA){
      let txHis1 = sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol1).getValue();
      if(txHis1 != ""){
        let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol1).getValue();
        let txEnd1 = sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol1).getValue();
        if(txSta1 != "" && txEnd1 != ""){
          // �ō���������Ă���
        }
        else{
          // �ō���������Ă��Ȃ�
          let txSlackId1 = sht1.getRange(1 + idRow1, 1 + COL_SLK_ID).getValue();
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
            let txDate2 = sht1.getRange(1 + idRow1, 1 + COL_OUT).getValue();
            let dateOut1 = Moment.moment(txDate2);
            
            let txTimeEnd1 = dateOut1.format("HH:mm");
            
            txMsg1 = txMsg1 + "�o��:" + txTimeSta1 + "\n"
            txMsg1 = txMsg1 + "�ދ�:" + txTimeEnd1 + "\n"
            
            SendSlkMsg(txMsg1, "@"+txSlackId1);
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
function GetDateCol(sht1, date1){
  let txDate1 = date1.clone().add(-6, "h").format("MM/DD");
  // �ō����t���擾
  let idCol1 = 0;
  // ���t�s���E���猟��
  for(idCol1 = sht1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
    // ���t�s������t���擾
    let txDate2 = sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
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
  AddLog(txVal1);
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("ID�Ǘ�");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    let txCol2 = sht1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){
      break;
    }
  }
  
  //AddLog(txVal1);
  if(idCol1 >= sht1.getLastColumn()){
    // �w��̗񂪂���܂���B
    AddLog("��");
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
    // �w���ID������܂���B
    AddLog("�s");
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

  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("ID�Ǘ�");
  let idCol1;
  for(idCol1 = ID_COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    let txCol2 = sht1.getRange(1 + ID_ROW_CPT, 1 + idCol1).getValue();
    if(txCol2 == txCol1){

      break;
    }
  }
  
  if(idCol1 >= sht1.getLastColumn()){
    // �w��̗񂪂���܂���B
    return("");
  }
  
  let txVal1 = sht1.getRange(1 + idMemRow1, 1 + idCol1).getValue()
  
  return(txVal1);
}

// ============================================================================
// �`�����l���ړ����s
// ============================================================================
function ChnSftExe(sht1, date1, idRow1, txAftSvr1, txAftChn1, txName1){
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
    for(idCol1 = sht1.getLastColumn() - 1; idCol1 >= COL_DATA2; idCol1--){
      // ���t�s������t���擾
      let txDate2 = sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
      if(txDate2 == txDate1){
        // ����
        break;
      }
    }
    
    if(idCol1 < COL_DATA2){
      // ���t���Ȃ���ΉE�[�ɒǉ�
      idCol1 = sht1.getLastColumn();
      sht1.getRange(1 + ROW_DATE, 1 + idCol1).setValue("'" + txDate1);
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
        let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
        if(txSta1 != ""){
          // �o�Бō�����      
          let txEnd1 = sht1.getRange(1 + idRow1 + ROW_DATA_END, 1 + idCol2).getValue();
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
      //AddLog(txId3);
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
      let txSta1 = sht1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).getValue();
      if(txSta1 == ""){
        // �o�Αō��Ȃ��Ȃ�o�Αō����s��
        sheet1.getRange(1 + idRow1 + ROW_DATA_STA, 1 + idCol2).setValue(txTime1);
        //let txId3 = "'" + sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
        //let idMemRow1 = GetMemRow("Discord ID", txId3);
        AkashiDakoku(idMemRow1, "�o��", date1);
      }
    }
    
    //let txChn1 = "C01AG9H3GBF";
    let txChn1 = "C01805HS02F";
    
    if(txAftChn1 == "�o��"){
      // �ꏊ���X�V
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        SendSlkMsg(txName1 + "���e�����[�N���I�����A�o�Ђ��܂����B", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
      // Iruca�X�e�[�^�X���u�ݐȁv�ɕύX
      Iruca_WorkStartOffice(idMemRow1);
    }
    else if(txAftChn1 == "�e�����[�N�J�n"){
      // �ꏊ���X�V
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() == ""){
        SendSlkMsg(txName1 + "���e�����[�N���J�n���܂����B", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("�e�����[�N��");
      }
      // Iruca�X�e�[�^�X���u�ݐȁv�e�����[�N�ɕύX
      Iruca_WorkStartHome(idMemRow1);
    }
    else if(txAftChn1 == "�ދ�"){
      if(sht1.getRange(1 + idRow1, 1 + COL_PLACE).getValue() != ""){
        SendSlkMsg(txName1 + "���e�����[�N���I�����܂���", txChn1);
        sht1.getRange(1 + idRow1, 1 + COL_PLACE).setValue("");
      }
      // Iruca�X�e�[�^�X���u�x�Ɂv�ɕύX
      Iruca_WorkEnd(idMemRow1);
    }
    
    let txHis1 = sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).getValue();
    
    if(txHis1 != ""){
      txHis1 = txHis1 + "��";
    }
    
    // ���Ԃ��L�^
    txHis1 = txHis1 + "["+ txTime1 + "]";      
    let txNowSvr1 = sht1.getRange(1 + idRow1, 1 + COL_SVR).getValue(); 
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
    sht1.getRange(1 + idRow1 + ROW_DATA_HIS, 1 + idCol2).setValue(txHis1);
    
    if(txAftSvr1 != ""){
      // ��Ԃ��X�V
      sht1.getRange(1 + idRow1, 1 + COL_STT).setValue(txAftChn1);
      // �T�[�o���X�V
      sht1.getRange(1 + idRow1, 1 + COL_SVR).setValue(txAftSvr1);
      
      // Iruca���b�Z�[�W���X�V
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
  let sheetErr1 = book1.getSheetByName("�G���[���O");
  sheetErr1.getRange(1 + sheetErr1.getLastRow(), 1).setValue(log);
}

// ============================================================================
// �̉����X�V
// ============================================================================
function UpdTaion(jsPrm1){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�̉��Ǘ�");

  let txName1 = getUserName(jsPrm1.event.user);
  let txId1 = jsPrm1.event.user;

  if(txName1 == "KintaiKanri"){
    return;
  }
  
  // �e�L�X�g����̉��̂ݒ��o
  let txVal1 = jsPrm1.event.text;
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
    SendSlkMsg("�u"+ txVal1 + "�v"+ "\n���͂������ł��B\n��:36.2�A36�A362", "@" + txId1);
    return;
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1++){
    // ID�񂩂�ID���擾
    let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
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
  for(idCol1 = COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = "'" + sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  // ���t���X�V
  sht1.getRange(1 + ROW_DATE, 1 + idCol1).setValue(txDate1);
  // ���O���X�V
  sht1.getRange(1 + idRow1, 1 + COL_NAME).setValue(txName1);
  // �O�̂���ID���X�V
  sht1.getRange(1 + idRow1, 1 + COL_ID).setValue(txId1);
  // �����̑̉����X�V
  sht1.getRange(1 + idRow1, 1 + idCol1).setValue(txVal1);

  // �f�o�b�O�p
  //sht1.getRange(1 + ROW_DATE, 1 + COL_NAME).setValue(jsPrm1);
}


// ============================================================================
// Slack��ID����Slack�̕\�������擾
// ============================================================================
function getUserName(txSlkId1){
  const txTkn1 = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+txTkn1+"&user="+txSlkId1).getContentText();

  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�̉��Ǘ�");

  const userInfo = JSON.parse(userData).user;  
  const userProf =userInfo.profile;
  const userName1 = userProf.display_name;
  const userName2 = userInfo.real_name;

  return userName1 ? userName1 : (userName2 ? userName2 : txSlkId1); 
}

// ���b�Z�[�W�𑗐M
function SendMessage(){
  let book1 = SpreadsheetApp.getActiveSpreadsheet();
  let sht1 = book1.getSheetByName("�̉��Ǘ�");

  // �����̓��t���擾
  var date = new Date();
  let txDate1 = "'" + Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');
  
  // �����̓��t����擾
  let idCol1 = COL_DATA;
  for(idCol1 = COL_DATA; idCol1 < sht1.getLastColumn(); idCol1++){
    // ���t�s������t���擾
    let txDate2 = "'" + sht1.getRange(1 + ROW_DATE, 1 + idCol1).getValue();
    if(txDate2 == txDate1){
      break;
    }
  }
  
  let idRow1;
  for(idRow1 = ROW_DATA; idRow1 < sht1.getLastRow(); idRow1++){
    // ���̗񂩂疼�̂��擾
    
    let txOndo2 = sht1.getRange(1 + idRow1, 1 + idCol1).getValue();
    if(txOndo2 == ""){
      // ���x���L������Ă��Ȃ�
      let txName2 = sht1.getRange(1 + idRow1, 1 + COL_NAME).getValue();
      let txId2 = sht1.getRange(1 + idRow1, 1 + COL_ID).getValue();
      SendSlkMsg("�̉�����͂��Ă�������", "@"+txId2);
    }
  }
}

// ============================================================================
// Slack�Ƀ��b�Z�[�W�𑗐M
// txChn1:@���[�UID���w�肷���DM�𑗐M�ł��܂��B
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
    
    // Slack�ɓ��e����
    let res1 = UrlFetchApp.fetch(url, jsPrm1);
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
