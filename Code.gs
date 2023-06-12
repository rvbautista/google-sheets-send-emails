// Detect email list then send templated emails
// Link to this file or to https://github.com/rvbautista/google-sheets-send-emails/

function emailMerge() {
  
  var emailsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("emails");
  var subject = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("subjectbody").getRange(2,2).getValue();
  var message = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("subjectbody").getRange(3,2).getValue();
  var startRow = 2; // First row of data to process
  var rows = emailsheet.getDataRange().getValues();
  var numRows = rows.length;

  for(var i=1;i<numRows;i++){
    //Takes data of Q Column
    var sentMark = rows[i][3];
    Logger.log(sentMark);
    if (sentMark == false){

      var emailacct = rows[i][1];
      Logger.log(emailacct);
      var name = rows[i][0];
      Logger.log(name);

      // email information

      MailApp.sendEmail(emailacct, subject, message);

      var cell = emailsheet.getRange(i+1,4); 
      Logger.log(cell);
      cell.setValue(new Date()).setNumberFormat("MM/dd/yyyy hh:mm:ss");
  }

} 
}
