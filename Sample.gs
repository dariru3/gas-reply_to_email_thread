function listMessagesAndReplies() {
  // Get the active spreadsheet and the active sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Get the thread ID from the first column of the active row
  var threadId = sheet.getRange(sheet.getActiveCell().getRow(), 1).getValue();

  // Get the thread using the thread ID
  var thread = GmailApp.getThreadById(threadId);

  // Get the messages in the thread
  var messages = thread.getMessages();

  // Start from the second row of the sheet
  var row = 2;

  // Write the messages and replies to the sheet
  for (var i = 0; i < messages.length; i++) {
    // Write the message to the first column of the current row
    sheet.getRange(row, 1).setValue(messages[i].getPlainBody());

    // Write the reply to the second column of the current row
    sheet.getRange(row, 2).setValue("This is a reply to your message.");

    // Increment the row counter
    row++;
  }

  // Reply to each message
  for (var i = 0; i < messages.length; i++) {
    // Get the reply from the second column of the current row
    var reply = sheet.getRange(i + 2, 2).getValue();

    // Check if the reply cell is empty
    if (reply != "") {
      // Reply to the message
      messages[i].reply(reply);
    }
  }
}
