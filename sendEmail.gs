function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 6;   // Number of rows to process
  
  var participantsRange = sheet.getRange(startRow, 2, numRows, 3);
  var participants = participantsRange.getValues().map(function(p) {
    return {
      name: p[0],
      email: p[1],
      wishList: p[2]
    };
  });
  var usedNumArr = [];
  var randomNum = Math.floor(Math.random() * participants.length);
  
  for (var i = 0; i < participants.length; i++) {
    var santa = participants[i];
    
    while (~usedNumArr.indexOf(randomNum) || (i == randomNum)){
      randomNum = Math.floor(Math.random() * participants.length);
    }
    usedNumArr.push(randomNum);
    
    //send email 
    var recipient = participants[randomNum];
    //var recepEmailAddress = row[1];
    var recepName = recipient.name;
    var recepList = recipient.wishList;
    var santaName = santa.name;
    var message = "Yoyo " + santaName + "! You are the Secret Santa for " + recepName +"!. Reminder that the range is $15-$25. This is their wish list: \n" + recepList;
    var subject = "Your Team Us Secret Santa Is...(open to reveal)";
    MailApp.sendEmail(santa.email, subject, message);
  }
}
