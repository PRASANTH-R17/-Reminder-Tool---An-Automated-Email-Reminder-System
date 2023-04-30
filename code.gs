let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"); // Access Sheet 01.
let outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Output"); // Access Sheet 02.
let lastCol = ss.getLastColumn();
const now = new Date(); // Get the current date (today's date).
let currentDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0); // Convert the current date to 00:00:00 hours.

// This function creates two custom menus in the toolbar for sending emails and filling data in Sheet 02. This function makes these tasks manual.
function addMenu() {
  var menu = SpreadsheetApp.getUi().createMenu("Custom Menu");
  menu.addItem("Mail Sender", "myFunction");
  menu.addItem("Auto Filler", "autoFill");
  menu.addToUi();
}

// This function makes Data can be auto-filled in Sheet 02.
function autoFill() {
  let badMails = []; // Array containing emails to avoid sending duplicate emails.
  let sslastRow = lastRowFun(ss); // Get the last row of Sheet 2 to write data.
  outSheet.getRange("A2:E" + outSheet.getLastRow() + 1).clearContent(); // Clear previous data.

  let serialNo = 1;
  for (let i = 2; i < sslastRow + 1; i++) {
    writtenRow = outSheet.getLastRow() + 1;
    email = ss.getRange(i, 3).getValue();

    if (badMails.includes(email))
      continue;
    badMails.push(email);

    name = ss.getRange(i, 2).getValue();
    projects = [];

    for (let j = i; j < sslastRow + 1; j++) { // Get the different project details of each person.

      // Get project details before the due date.
      if ((ss.getRange(j, 3).getValue() == email) && (currentDate <= ss.getRange(j, 7).getValue()) && ss.getRange(j, 8).getValue() == false) {
        projects.push(ss.getRange(j, 4).getValue());
        //Logger.log(currentDate + "=" + ss.getRange(j, 7).getValue());
      }
      // Get overdue project details.
      else if ((ss.getRange(j, 3).getValue() == email) && (currentDate > ss.getRange(j, 7).getValue()) && ss.getRange(j, 8).getValue() == false) {
        projects.push(ss.getRange(j, 4).getValue() + " *");
      }
    }
    if (projects.length == 0)
      continue;

    // These lines return data in Sheet 02.
    outSheet.getRange(writtenRow, 1).setValue(serialNo);
    serialNo = serialNo + 1; // Increment the serial number.
    outSheet.getRange(writtenRow, 2).setValue(name);
    outSheet.getRange(writtenRow, 3).setValue(email);
    outSheet.getRange(writtenRow, 4).setValue(projects.join("\n"));
    outSheet.getRange(writtenRow, 5).setValue(projects.length);
  }
}
// main Function - control others functions
function myFunction() {
  autoFill();
  
  // html01 file contains project details for a person with different projects
  let html01 = htmlToString("Html01.html");
  // html02 file contains project details for a person without different or remaining projects
  let html02 = htmlToString("Html02.html");

  let lastRow = lastRowFun(ss);
  for (let i = 2; i < lastRow + 1; i++) {
    let datas = []; // stores each person's data in a nested array
    let status = ss.getRange(i, 8).getValue();
    let dueDate = ss.getRange(i, 7).getValue();
    let dateDiff = dueDate - currentDate;
    dateDiff = dateDiff / (1000 * 60 * 60 * 24);

    // if project status is ticked ( completed ), the program will stop
    if (status) {
      continue;
    } else if (dateDiff != 0 && dateDiff != 2) { // this condition is used to send mail only two days before the due date or on the due date
      
      continue;
    }

    let serialNo = ss.getRange(i, 1).getValue();
    let name = ss.getRange(i, 2).getValue();
    let email = ss.getRange(i, 3).getValue();
    let projectName = ss.getRange(i, 4).getValue();
    
    datas.push([serialNo, name, email, projectName, dueDate, dateDiff]);

    // includes the details of remaining projects in tabular format in the email
    datas = getBalanceData(datas, serialNo, email, lastRow);
    mailTemplate(datas, html01, html02);
  }
}

// Convert HTML file content to a string because it makes it easier to write project details
function htmlToString(fileName) {
  let template = HtmlService.createTemplateFromFile(fileName);
  let htmlString = template.evaluate().getContent();
  return htmlString;
}

function lastRowFun(sheet) {
  let lastRow = 1;
  for (let i = 2; i < sheet.getLastRow(); i++) {
    lastRow = lastRow + 1;
    if (sheet.getRange(i, 1).getValue() == "") {
      break;
    }
  }
  return lastRow - 1;
}

// This function gets the details of remaining projects for each person
function getBalanceData(datas, removeSerialNo, email, lastRow) {
  for (let i = 2; i < lastRow + 1; i++) {
    let status = ss.getRange(i, 8).getValue();
    let currentSerialNo = ss.getRange(i, 1).getValue();
    let currentEmail = ss.getRange(i, 3).getValue();
    let currentDueDate = ss.getRange(i, 7).getValue();
    
    if (!status && !(currentSerialNo == removeSerialNo) && currentDate <= currentDueDate && currentEmail == email) {
      datas.push([currentSerialNo, ss.getRange(i, 2).getValue(), currentEmail, ss.getRange(i, 4).getValue(), currentDueDate]);
    }
  }

  return datas;
}

// This function chooses between two HTML templates (HTML 01 or HTML 02) based on the data and creates an HTML body (mail content).
function mailTemplate(datas, html01, html02) {
  if (datas.length > 1) {
    let name = datas[0][1];
    let email = datas[0][2];
    let projectName = datas[0][3];
    let dateDiff = datas[0][5];
    let content = "";
    
    // This condition executes if a person has any remaining projects (more than 1) - use template 01 (HTML 01 file).
    if (dateDiff == 2) {
      content = "Your project named <b>" + projectName + "</b> is due in 2 days, so please complete it within the given timeframe.";
    } else {
      content = "Your project named <b>" + projectName + "</b> is due today, so please submit it today.";
    }

    let msg = html01;
    let tableContent = "";

    for (let i = 1; i < datas.length; i++) {
      tableContent = tableContent + "<tr><td>" + datas[i][3] + "</td><td>" + dataFormat(datas[i][4]) + "</td></tr>";
      //Logger.log(datas[i][4].toString());
    }

    msg = msg.replace("=name=", name);
    msg = msg.replace("=project detail=", content);
    msg = msg.replace("=table content=", tableContent);

    MailSender(email, msg);
    //Logger.log(msg);
  } 
  // This condition executes if a person doesn't have any remaining projects - use template 02 (HTML 02 file).
  else {
    let name = datas[0][1];
    let email = datas[0][2];
    let projectName = datas[0][3];
    let dateDiff = datas[0][5];
    let content = "";

    if (dateDiff == 2) {
      content = "Your project named <b>" + projectName + "</b> is due in 2 days, so please complete it within the given timeframe.";
    } else {
      content = "Your project named  <b>" + projectName + "</b> is due today, Please submit it before the deadline.";
    }

    let msg = html02;
    let tableContent = "";

    msg = msg.replace("=name=", name);
    msg = msg.replace("=project detail=", content);
    msg = msg.replace("=table content=", tableContent);

    //Logger.log(msg);
    MailSender(email, msg);
  }
}

// This function sends an email to the respective person.
function MailSender(email, msg){

  MailApp.sendEmail({
    to: email,
    subject: "Reminder Mail",
    htmlBody: msg,
  });
}

// Converts a date into this format dd-mm-yyyy for adding this date in mail.
function dataFormat(oldDate) {
  let date = new Date('Sun Apr 30 2023 00:00:00 GMT+0530 (India Standard Time)');
  let day = String(date.getDate()).padStart(2, '0');
  let month = String(date.getMonth() + 1).padStart(2, '0');
  let year = date.getFullYear();
  let formattedDate = `${day}-${month}-${year}`;
  return(formattedDate);
}
