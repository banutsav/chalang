---
layout: post
title:  "Send email with an image in the email body"
date:   2021-07-10 22:30:05 +0530
tags: ['google-app-script', 'email']
---

Firstly the global variables associated with the scripts should be in a separate file

```js
var PIC_ID = '<ID OF THE IMAGE IN GOOGLE DRIVE>';
var SHEET = '<ID OF THE SHEET WHERE DATA COMES FROM>';
var DATA = '<TABNAME WHERE DATA COMES FROM>';
var BCC = 'email1@company.com, email2@company.com';
```

The image which needs to be inserted is kept in a google drive location, the script **should have access** to this location.

The main driving function is shown below - 

```js
// check to see if someone's birthday is today
function send_birthday_email()
{
  // get data
  var sheet = SpreadsheetApp.openById(SHEET).getSheetByName(DATA);
  var rows = sheet.getDataRange().getValues();
  
  // get current day and month
  var now = new Date();
  var month = now.getMonth();
  var day = now.getDate();
  //Logger.log('today = ' + month + ',' + day);
  
  // iterate over data
  for (var i = 1; i < rows.length; i++)
  {
    var row = rows[i];
    var email = row[0];
    
    // end of data
    if (row[0] == '')
      break;
    
    // ignore if email already sent
    if (row[4] != '')
      continue;
    
    var birthday = new Date(row[3]);
    var name = titleCase(row[1]) + ' ' + titleCase(row[2]); // helper.gs
    
    // ignore if birthday not today
    if (!check_if_birthday(month, day, birthday)) // helper.gs
      continue;
    
    //Logger.log('sending email to ' + name);
    
    // send emailwith inline image
    send_a_pretty_email(name, email); // helper.gs
    
    // mark as sent
    sheet.getRange('E' + (i + 1)).setValue(now);
  }
}
```
Once the email is sent, it is marked so in column E of that row - so that it's not repeated again

```js
// mark as sent
sheet.getRange('E' + (i + 1)).setValue(now);
```

It uses some other secondary functions to capitalize the first letter of the person's name and also to check whether today is the person's birthday or not

```js
// uppercase first and lower case the rest of the string
function titleCase(string){
  if (string == '')
    return '';
  return string[0].toUpperCase() + string.slice(1).toLowerCase();
}

// check if today is person's birthday
function check_if_birthday(month, day, birthday)
{
  // check month
  if (birthday.getMonth() != month)
    return false;
  
  // check day
  if (birthday.getDate() != day)
    return false;
  
  return true;
}
```

The actual email is sent out using this function below - notice the **inlineImages: {sampleImage: img}** which is referring to the image as a [blob](https://en.wikipedia.org/wiki/Binary_large_object)

```js
// send email with attached image
function send_a_pretty_email(name, email)
{
  var img = DriveApp.getFileById(PIC_ID).getBlob();
  var html = '<p><h2>Dear ' + name + ',</h2></p>' 
  + '<p><img src="cid:sampleImage" border="0"></p>'
  ;
  
  MailApp.sendEmail({
    to: email,
    noReply: true,
    bcc: BCC,
    subject: "Greetings from the Company",
    htmlBody: html,
    inlineImages: {sampleImage: img}
  });

}
```

