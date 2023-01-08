function getStarredThreads() {
  const threads = GmailApp.getStarredThreads();
  let threadIds = [];
  let threadSubjects = [];

  threads.forEach(thread => {
    const threadId = thread.getId();
    threadIds.push(threadId);
    const threadSubject = GmailApp.getThreadById(threadId).getFirstMessageSubject();
    threadSubjects.push(threadSubject)
  });
  console.log(threadIds);
  console.log(threadSubjects);
}


function removeQuotes_(text){
  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //const text = sheet.getRange('A2').getValue();
  let cleanText = text.replace(/^>.*\n?/gm, '').trim();
  console.log("clean:", cleanText);
  return cleanText
}

function replyToThread() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const threadId = "184648512fcffa88";
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();
  for(i=1; i<messages.length; i++){
    const messageOnly = removeQuotes_(messages[i].getPlainBody());
    sheet.getRange(i+1,1).setValue(messageOnly);
  }

  const data = sheet.getDataRange().getValues();

  const lastColumn = sheet.getLastColumn();
  const headersText = sheet.getRange(1,1,1,lastColumn).getValues();
  const headers = {};
  for (let i = 0; i < headersText[0].length; i++) {
    const header = headersText[0][i];
    headers[`${header}_Col`] = data[0].indexOf(header);
  }
  //console.log("header:index", headers);
  /*
  { email_Col: 0,
    message_Col: 1,
    'message translated_Col': 2,
    status_Col: 3 }
  */

  for(i=1; i<data.length; i++){
    let emailAddress = data[i][headers['email_Col']];
    let reply = data[i][headers['reply_Col']];
    let replyJapanese = googleTranslate(reply);
    let status = data[i][headers['status_Col']];

    let messageAll = reply
    messageAll += '\n'
    messageAll += replyJapanese
    console.log("message:", messageAll);

    //GmailApp.sendEmail(emailAddress,null,messageAll,{threadId:emailThreadId});
  }
  
}

function googleTranslate(text){
  let textToJapanese = LanguageApp.translate(text,'en','ja');
  return textToJapanese;
}