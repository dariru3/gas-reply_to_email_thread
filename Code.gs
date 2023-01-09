/**
 * Add menu to menu bar.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Instructions as well as function buttons.
  ui.createMenu('Import Email Thread')
      .addItem('1. Get email thread in Starred items', 'menuItem1_')
      .addItem('2. Copy thread id, update script', 'menuItem1_')
      .addSeparator()
      .addItem('3. List messages', 'listMessages')
      .addItem('4. Add reply next to message', 'menuItem2_')
      .addItem('5. Send replies', 'replyToThread')
      .addToUi();
}

/**
 * Placeholder function. Future update: add sidebar.
 * Instructions to open script to get email thread
 * and update thread id.
 */
function menuItem1_() {
  SpreadsheetApp.getUi().alert('Open App Script');
}

/**
 * Placeholder function. Future update: add sidebar.
 * Continue how-to instructions.
 */
function menuItem2_(){
  SpreadsheetApp.getUi().alert('Enter reply in column B');
}

/**
 * List thread subjects and thread ideas
 * from Starred messages into console log.
 * Future update: list in sidebar.
 */
function getStarredThreads() {
  // get threads from starred messages
  const threads = GmailApp.getStarredThreads();
  let threadIds = [];
  let threadSubjects = [];
  // loop through each messages for id and subject line
  threads.forEach(thread => {
    const threadId = thread.getId();
    threadIds.push(threadId);
    const threadSubject = GmailApp.getThreadById(threadId).getFirstMessageSubject();
    threadSubjects.push(threadSubject)
  });
  // view and choose in console log.
  console.log(threadSubjects);
  console.log(threadIds);  // copy thread and paste into getThreadMessages_()

}

/**
 * Helper function to shorted imported email body.
 * Removes quotes of previous emails by searching for lines
 * beginning with ">"
 * @param text {string} Text containing email body
 * @returns Email body without quotes, lines starting with ">"
 */
function removeQuotes_(text){
  let cleanText = text.replace(/^>.*\n?/gm, '') // search for line that starts with ">" (and new line)
  .trim(); // remove whitespace before and after
  return cleanText
}

/**
 * Helper function.
 * @returns an array of messages in a given thread.
 */
function getThreadMessages_(){
  const threadId = "1859533f39c4783e"; // get from getStarredThreads()
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();
  return messages
}

/**
 * Lists messages in a thread
 * from the startRow and startColumn.
 */
function listMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const messages = getThreadMessages_();
  const startRow = 1;
  const startColumn = 1;
  for(i=1; i<messages.length; i++){
    const messageCleaned = removeQuotes_(messages[i].getPlainBody());
    sheet.getRange(i+startRow,startColumn).setValue(messageCleaned);
  }
  SpreadsheetApp.flush();
}

function replyToThread() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const messages = getThreadMessages_();
  const data = sheet.getDataRange().getValues();
  // collect column header positions
  const lastColumn = sheet.getLastColumn();
  const headersText = sheet.getRange(1,1,1,lastColumn).getValues();
  const headers = {};
  for (let i = 0; i < headersText[0].length; i++) {
    const header = headersText[0][i];
    headers[`${header}_Col`] = data[0].indexOf(header);
  }
  // end of column header positions

  // loop through each row, starting with row 2
  for(i=1; i<data.length; i++){
    let reply = data[i][headers['reply_Col']];
    let status = data[i][headers['status_Col']];

    let messageAll = "";
    if(reply !== '' && status == ''){ // check for a reply to a certain thread that has not been sent
      let replyJapanese = LanguageApp.translate(reply, 'en', 'ja');
      // optional: update spreadsheet to show translation
      sheet.getRange(i+1, headers['reply translated_Col']+1).setValue(replyJapanese);
      // collect reply and reply translated
      messageAll += reply;
      messageAll += '\n\n';
      messageAll += replyJapanese;
      // send reply to single recipient
      messages[i].reply(messageAll);
      // update spreadsheet to show email has been sent
      sheet.getRange(i+1, headers['status_Col']+1).setValue("Sent")
    }
  }
}