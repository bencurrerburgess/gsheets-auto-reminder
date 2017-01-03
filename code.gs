// Auto send email reminders for repeated tasks from a google sheet.

// This constant is written in column C for rows for which an email has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";

var DUE = "yes";
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row to process
  var numRows = 30;   // Number of rows to process
  var numColumns = 10; //Number of columns to process
  var startColumn = 1; //First column to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // column A
    var message = row[1];       // column B
    var isDue = row[2];         // column C
    var emailSent = row[3];     // column D
    var daysLeft = row[8];      // column I
    
    if (emailSent != EMAIL_SENT && isDue == DUE) {  // Prevents sending duplicates & only sends if send mail box is true
      var subject = "Regular Website update is Due: " + daysLeft + " Days";
      MailApp.sendEmail(emailAddress,
                        subject,
                        message 
                        +"\n\n---------------------------------------------------------------------------"
                        +"\n(This is an automatic message from: https://docs.google.com/"
                        +"\n---------------------------------------------------------------------------"
                       );
      sheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}

function resetSendEmails() {
   var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row to process
  var numRows = 30;   // Number of rows to process
  var numColumns = 10; //Number of columns to process
  var startColumn = 1; //First column to process

  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns)
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // column A
    var message = row[1];       // column B
    var isDue = row[2];         // column C
    var emailSent = row[3];     // column D
    var daysLeft = row[8];      // column I
    
    sheet.getRange(startRow + i, 4).setValue("") // RESETS THE EMAIL_SENT COLUMN
    SpreadsheetApp.flush();
  }
}

function doIt () {
sendEmails2();
resetSendEmails();
}
