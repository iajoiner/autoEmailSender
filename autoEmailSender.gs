//Copyright Notice
//This file is a modified version of the second example in https://developers.google.com/apps-script/articles/sending_emails, 
//written by Hugo Fierro and released by him under Apache 2.0 License.
//I, Ying Zhou, modifed and released this file under Apache 2.0 License.
//The license is available here: http://www.apache.org/licenses/LICENSE-2.0

//var EMAIL_SENT = "EMAIL_SENT"; //"EMAIL_SENT" can be used in order to make sure that you don't send the emails twice. One example where
//it is used is here: https://developers.google.com/apps-script/articles/sending_emails
var defaultEmailAddress = "yingzhou474@gmail.com"// Please put your email address here
var defaultFirstName = "Ying" // Please put your first name here
var defaultSurname = "Zhou" // Please put your last name here

function sendEmails() {
  var spreadSheet = SpreadsheetApp.openById("<Please put the ID of your spreadsheet here>"); 
  //SpreadsheetApp.setActiveSheet(spreadSheet.getSheets()[1]);
  var sheet = spreadSheet.getSheetByName("<<Please put the name of your sheet in the spreadsheet above here>>"); //
  var startRow = 2;  // First row of data to process, it is usually 2
  var numRows = 13;   // Number of rows to process
  var dataRange = sheet.getRange(startRow, 1, numRows, 4) // Fetch the range of cells A2:D14
  var data = dataRange.getValues();
  var subject = "<Please put your subject here>";
  var doc = DocumentApp.openById("<Please put the ID of the Google Document containing your email except for the initial greetings here>");
  //MailApp.sendEmail(defaultEmailAddress, subject, message);
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[2];  // This is the third column in my sheet, namely the first name
    if (row[3] != "No") // This is the fourth column in my sheet, namely whether an email needs to be sent to the students
    {
      var message = "Hi ".concat(row[0]);
      message = message.concat(",\n\n");
      message = message.concat(doc.getBody().getText());
      MailApp.sendEmail(emailAddress, subject, message);
    }
  }
}
