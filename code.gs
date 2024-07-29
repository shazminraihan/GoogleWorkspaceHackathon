/*

Google Workspace Hackathon 2024
Team Imminent
Case Study 2: Automating Leave Tracking and Management

*/

let MySheets  = SpreadsheetApp.getActiveSpreadsheet();
let LoginSheet  = MySheets.getSheetByName("employeeLogin");

function doGet(e) {
    return HtmlService.createTemplateFromFile("index").evaluate();
}

function LoginCheck(pUID, pPassword) {
  let ReturnData = LoginSheet.getRange("A:A").createTextFinder(pUID).matchEntireCell(true).findAll();
    let StartRow = 0;
    ReturnData.forEach(function (range) {
      StartRow = range.getRow();
    });

    let TmpPass = 0;
    if (StartRow > 0) {
        TmpPass = LoginSheet.getRange(StartRow, 2).getValue();
        if (TmpPass == pPassword) {
            return true;
        }
    }
    return false;
}

function OpenPage(PageName) {
    return HtmlService.createHtmlOutputFromFile(PageName).getContent();
}


function sendEmail() {
  var DataSheet = MySheets.getSheetByName("employeeData")   
  var lastRow = DataSheet.getLastRow();
  var calendarId = 'c_0d5f491fe3f3d56bcde7c63de592d14f26c5fd3007fca05486930c7327b46ea7@group.calendar.google.com'; 
  var calendar = CalendarApp.getCalendarById(calendarId);

  for (var i = 2; i <= lastRow; i++) { // Starting from the second row
    var row = DataSheet.getRange(i, 1, 1, 7).getValues()[0]; // Adjust the range if needed
    var name = row[1]; // Employee Name is in the first column 
    var startLeave = new Date(row[2]); 
    var endLeave = new Date(row[3]);  
    var emailAddress = row[4]; 
    var approvalStatus = row[5]; 
    var emailSentStatus = row[6];  

    if (approvalStatus === "Approved" && emailSentStatus !== 'Sent') {
        var subject = "Time off is approved!";
        var message = "Hello " + name + ", your time off has been approved as follows:\n";
        message += "Name: " + name + "\n";
        message += "Email: " + emailAddress + "\n";
        message += "Start Leave: " + startLeave + "\n";
        message += "End Leave: " + endLeave + "\n";

        MailApp.sendEmail(emailAddress, subject, message);
        DataSheet.getRange(i, 7).setValue('Sent'); 
        SpreadsheetApp.flush(); 

       // Add calendar for employees with approved time off
        calendar.createEvent(
          'Leave: ' + name,
          startLeave,
          endLeave,
          {description: 'Approved leave for ' + name}
        );

    } else if (approvalStatus === "Declined" && emailSentStatus !== 'Sent') {
        var subject = "Time off request declined";
        var message = "Hello " + name + ", your time off request has been declined.";
        
        MailApp.sendEmail(emailAddress, subject, message);
        DataSheet.getRange(i, 7).setValue('Sent'); 
        SpreadsheetApp.flush(); 
    }
  }
}
