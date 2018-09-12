function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var startCol = 2; // The first column of data to process
  var numRows = 6;   // Number of rows to process
  var fields = ['name', 'email', 'wishList'];

  var participantsRange = sheet.getRange(startRow, startCol, numRows, fields.length);
  var participants = participantsRange.getValues().map(function(p) {
    var pObj = {};
    for (var i = 0; i < fields.length; i++) {
      var fieldName = fields[i];
      pObj[fieldName] = p[i];
    }
    return pObj;
  });
  var usedNumArr = [];
  var randomNum = Math.floor(Math.random() * participants.length);
  
  for (var i = 0; i < participants.length; i++) {
    var santa = participants[i];
    
    while (~usedNumArr.indexOf(randomNum) || (i == randomNum)){
      randomNum = Math.floor(Math.random() * participants.length);
    }
    usedNumArr.push(randomNum);
    
    var recipient = participants[randomNum];
    var recepName = recipient.name;
    var recepList = recipient.wishList;
    var santaName = santa.name;
    var santaEmail = santa.email;
    
    //send email
    var message = "Yoyo " + santaName + "! You are the Secret Santa for " + recepName +"!. Reminder that the range is $15-$25. This is their wish list: \n" + recepList;
    var subject = "Your Team Us Secret Santa Is...(open to reveal)";
    MailApp.sendEmail(santaEmail, subject, message);
  }
}
