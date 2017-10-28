# Secret Santa Email Sender!

This is a Google Script function to go through your Secret Santa list! It will randomly generate Secret Santa matches, and it will not match any partipant with themselves. 

## How To Use: 

1. Create a Google Form with the Fields "Name, Email, Wishlist" In that order. They should all be text-input questions. 
2. Save the responses in a Google Sheet. 
3. Open up the Google Scripes for that Sheet, paste in the function `sendEmails()`. You probably will have to change the `numRows` to however many partipants you have. I only needed 6 in any case so I just hardcoded it in. 
4. Click "run". 
