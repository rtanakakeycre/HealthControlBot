function doPost(e) {
  try{
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
    const slackApp = SlackApp.create(token);
    
    const params = JSON.parse(e.postData.getDataAsString());
    writeLog(params);

    return ContentService.createTextOutput(params.challenge);
    
  }catch(err){
    writeLog(err);
  }
}


function writeLog(text){
  let spreadSheet1 = SpreadsheetApp.getActiveSpreadsheet();
  const channel = text.event.channel;

  
  let sheet1 = spreadSheet1.getSheetByName("ëÃâ∑ä«óù");
  Logger.log(text);

  sheet1.appendRow([new Date(), getUserName(text.event.user), text.event.text]); 
}


function getUserName(userId){
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');  
  const userData = UrlFetchApp.fetch("https://slack.com/api/users.info?token="+token+"&user="+userId).getContentText();

  const userName = JSON.parse(userData).user.real_name;
  Logger.log(userId,userName);

  return userName ? userName : userId; 
}