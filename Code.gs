function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Import Email Thread')
      .addItem('1. Get email thread in Starred items', 'menuItem1_')
      .addItem('2. Copy thread id, update script', 'menuItem1_')
      .addSeparator()
      .addItem('3. List messages', 'listMessages')
      .addItem('4. Add reply next to message', 'menuItem2_')
      .addItem('5. Send replies', 'replyToThread')
      .addToUi();
}

function menuItem1_() {
  SpreadsheetApp.getUi().alert('Open App Script');
}

function menuItem2_(){
  SpreadsheetApp.getUi().alert('Enter reply in column B');
}

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
  let cleanText = text.replace(/^>.*\n?/gm, '').trim();
  console.log("cleaned:", cleanText);
  return cleanText
}

function getThreadMessages_(){
  const threadId = "1859533f39c4783e";
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();
  return messages
}

function listMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const messages = getThreadMessages_();
  for(i=1; i<messages.length; i++){
    const messageOnly = removeQuotes_(messages[i].getPlainBody());
    sheet.getRange(i+1,1).setValue(messageOnly);
  }
  SpreadsheetApp.flush();
}

function replyToThread() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const messages = getThreadMessages_();

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
  { message_Col: 0,
    reply_Col: 1,
    'reply translated_Col': 2,
    status_Col: 3 }
  */
  
  for(i=1; i<data.length; i++){
    //let message = data[i][headers['message_Col']];
    //console.log("message:", message)
    let reply = data[i][headers['reply_Col']];
    let replyJapanese = googleTranslate_(reply);
    sheet.getRange(i+1, headers['reply translated_Col']+1).setValue(replyJapanese);
    let status = data[i][headers['status_Col']];

    let messageAll = "";
    if(reply !== '' && status == ''){
      console.log("reply:", reply);
      messageAll += reply;
      messageAll += '\n\n';
      messageAll += replyJapanese;
      console.log("message:", messageAll);
      messages[i].reply(messageAll);
      sheet.getRange(i+1, headers['status_Col']+1).setValue("Sent")
    }
  }
}

function googleTranslate_(text){
  let textToJapanese = LanguageApp.translate(text,'en','ja');
  if(text){
    console.log("translation:", textToJapanese);
  }
  return textToJapanese;
}
