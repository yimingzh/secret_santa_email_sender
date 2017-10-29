function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 6;   // Number of rows to process
  
  var recipientDataRange = sheet.getRange(startRow, 2, numRows, 3);
  var recipientData = recipientDataRange.getValues();
  var santaDataRange = sheet.getRange(2, 2, numRows, 2);
  var santaData = santaDataRange.getValues();
  var usedNumArr = [];
  var randomNum = Math.floor(Math.random() * recipientData.length);
  
  for (i in santaData) {
    while (~usedNumArr.indexOf(randomNum) || (i == randomNum)){
      randomNum = Math.floor(Math.random() * recipientData.length);
    }
    usedNumArr.push(randomNum);
    //send email 
    var row = recipientData[randomNum];
    //var recepEmailAddress = row[1];
    var recepName = row[0];
    var recepList = row[2];
    var santaName = santaData[i][0]
    var message = "Yoyo " + santaName + "! You are the Secret Santa for " + recepName +"!. Reminder that the range is $15-$25. This is their wish list: \n" + recepList;
    var subject = "Your Team Us Secret Santa Is...(open to reveal)";
    var emailAddress = santaData[i][1];
    MailApp.sendEmail(emailAddress, subject, message);
  }
}
