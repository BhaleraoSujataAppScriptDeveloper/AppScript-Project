function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Check if the active sheet is "MySheet"
  if (sheet.getName() == "Email Details") {
    // Check if the edited cell is in column I (9th column)
    if (range.getColumn() == 11) {
      var timestampCell = sheet.getRange(range.getRow(), 12); // Column J (next to I)

      if (range.getValue() != "") {
        // If I cell is not empty, insert current date and time in J cell
        timestampCell.setValue(new Date());
      } else {
        // If I cell is empty, clear the timestamp in J cell
        timestampCell.setValue("");
      }
    }
  }
   if (sheet.getName() == "Inquiry Process") {
    // Check if the edited cell is in column H (8th column)
    if (range.getColumn() == 7) {
      var timestampCell = sheet.getRange(range.getRow(), 21); // Column G (next to H)
      if (range.getValue() == "Sent") {
        // If the cell in column H is "Sent", insert the current date and time in Column G
        var currentDate = new Date();
        timestampCell.setValue(currentDate); // Set the Date object directly
      } else {
        // If the cell in column H is empty, clear the timestamp in Column G
        timestampCell.setValue("");
      }
    }
}
if (sheet.getName() == "Email Details") {
    // Check if the edited cell is in column H (8th column)
    if (range.getColumn() == 18) {
      var timestampCell = sheet.getRange(range.getRow(), 17); // Column G (next to H)
      if (range.getValue() == "DONE") {
        // If the cell in column H is "Sent", insert the current date and time in Column G
        var currentDate = new Date();
        timestampCell.setValue(currentDate); // Set the Date object directly
      } else {
        // If the cell in column H is empty, clear the timestamp in Column G
        timestampCell.setValue("");
      }
    }
 }
}
// Function to send the email with the clickable hyperlink
function sendDownloadLinkViaEmail(emailAddress, pdfUrl) {
  var subject = 'Your Sales Report is Ready';
  var body = 'Your report is ready! You can download it from the following link: ' + 
             '<a href="' + pdfUrl + '">Download PDF Report</a>';
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: body
  });
}

function populateDoer() {
  var form = FormApp.openById('1Ok2UqvEVxcGe3i1e8kUkuv5UxscugFvcBS5ovyYvrzs');  // Replace with your Form ID
  var sheet = SpreadsheetApp.openById('1b0I5poGdAcF__yvg8t5R1cmnbZyQmrPK46cGqPk5940');  // Replace with your Sheet ID
  var range = sheet.getSheetByName('Doer List').getRange('A2:A');  // Adjust range as needed
  var values = range.getValues();
  
  // Search for the dropdown question by its title
  var questions = form.getItems(FormApp.ItemType.LIST);
  var item = null;
  
  // Loop through all questions to find the one with the title you're looking for
  for (var i = 0; i < questions.length; i++) {
    if (questions[i].getTitle() === 'Doer List') {
      item = questions[i].asListItem();
      break;
    }
  }
  if (item) {
    // Flatten array, filter out empty rows, and remove duplicates
    var options = values.map(function(row) {
      return row[0];  // Get the value from the first column
    }).filter(function(option) {
      return option !== '';  // Remove empty options
    });
    // Remove duplicates by converting the array to a Set and back to an array
    options = Array.from(new Set(options));
    // Set the choices to the dropdown list
    item.setChoiceValues(options);
  } else {
    Logger.log('Question not found!');
  }
}

// Function to get the email address for a given sales executive name
function getEmailAddressForSalesExecutive(salesExecutive) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName('Sales Executives List'); // Sheet where sales executive names and emails are stored
  if (!salesSheet) {
    Logger.log("Sales Executives sheet not found!");
    return null;
  }
  var data = salesSheet.getDataRange().getValues(); // Get all data in the Sales Executives sheet
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === salesExecutive) { // Assuming sales executive name is in column A
      return data[i][1]; // Assuming email is in column B
    }
  }
  return null;
}
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) {
    return ""; // Return empty string if invalid date
  }
  var day = date.getDate();
  var month = date.getMonth() + 1; // Months are zero-indexed
  var year = date.getFullYear();
  // Add leading zeros for single-digit days or months
  return (day < 10 ? '0' : '') + day + '/' + 
         (month < 10 ? '0' : '') + month + '/' + 
         year;
}

function generateFilteredReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inquiry Process');
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }
  var salesExecutive = sheet.getRange('W3').getValue(); // Sales executive
  var fromDate = sheet.getRange('W1').getValue(); // From date
  var toDate = sheet.getRange('W2').getValue();   // To date
  var selectedType = sheet.getRange('W4').getValue(); // Type filter
  
  fromDate = new Date(fromDate);
  toDate = new Date(toDate);

  if (isNaN(fromDate) || isNaN(toDate)) {
    Logger.log('Invalid date entered.');
    return;
  }

  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);

  var data = sheet.getRange('A2:AC' + sheet.getLastRow()).getValues();
  var filteredData = data.filter(function(row) {
    var salesExec = row[15]; // Column P
    var type = row[5];       // Column F
    var plannedDate1 = new Date(row[21]);
    var dueDate1 = new Date(row[23]);
    var dueDate2 = new Date(row[25]);
    var dueDate3 = new Date(row[27]);
    var status = row[14];    // Column O
    var sentStatus = row[6]; // Column G

    plannedDate1.setHours(0, 0, 0, 0);
    dueDate1.setHours(0, 0, 0, 0);
    dueDate2.setHours(0, 0, 0, 0);
    dueDate3.setHours(0, 0, 0, 0);

    var statusCondition = status !== "RECEIVED" && sentStatus === "Sent" && status !== "CLOSED";

    // Apply Type logic:
var typeCondition =
  !selectedType ||
  (selectedType === "Tender" && type === "Tender") ||
  (selectedType === "Live Inquiry" && type !== "Tender");
    if (salesExecutive) {
      return (salesExec === salesExecutive) &&
             typeCondition &&
             ((plannedDate1 >= fromDate && plannedDate1 <= toDate) ||
              (dueDate1 >= fromDate && dueDate1 <= toDate) ||
              (dueDate2 >= fromDate && dueDate2 <= toDate) ||
              (dueDate3 >= fromDate && dueDate3 <= toDate)) &&
             statusCondition;
    } else {
      return typeCondition &&
             ((plannedDate1 >= fromDate && plannedDate1 <= toDate) ||
              (dueDate1 >= fromDate && dueDate1 <= toDate) ||
              (dueDate2 >= fromDate && dueDate2 <= toDate) ||
              (dueDate3 >= fromDate && dueDate3 <= toDate)) &&
             statusCondition;
    }
  });
  if (filteredData.length === 0) {
    Logger.log('No data found for the specified criteria.');
    Browser.msgBox('No data found for the selected criteria.');
    return;
  }

  var doc = DocumentApp.create('Sales Report' + (salesExecutive ? ' for ' + salesExecutive : ''));
  var body = doc.getBody();
  doc.setMarginTop(5);
  doc.setMarginBottom(5);
  doc.setMarginLeft(10);
  doc.setMarginRight(1);

  body.appendParagraph('Sales Report' + (salesExecutive ? ' for ' + salesExecutive : ''))
       .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Follow UP Date Range: ' + formatDate(fromDate) + ' to ' + formatDate(toDate))
       .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  var headers = [
    'Quotation Date', 'Company Name', 'Customer Name', 'Mobile Number', 
    'Project', 'Basic Value', 'Quotation Number', 'Sales Executive', 'Remark'
  ];
  headers.forEach(function(header) {
    var headerCell = headerRow.appendTableCell(header);
    headerCell.setBackgroundColor('#D3D3D3');
    headerCell.setBold(true);
    headerCell.setFontSize(8);
  });

  filteredData.forEach(function(row) {
    var tableRow = table.appendTableRow();
    tableRow.appendTableCell(formatDate(row[20])).setFontSize(8); 
    tableRow.appendTableCell(String(row[3])); 
    tableRow.appendTableCell(String(row[2])); 
    tableRow.appendTableCell(String(row[16])); 
    tableRow.appendTableCell(String(row[18])); 
    tableRow.appendTableCell(String(row[12])); 
    tableRow.appendTableCell(String(row[10])); 
    tableRow.appendTableCell(String(row[17])); 
    tableRow.appendTableCell('');
  });

  doc.saveAndClose();

  var pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  var pdfFile = DriveApp.createFile(pdf);
  var pdfUrl = pdfFile.getUrl();
  var email;
  if (salesExecutive) {
    email = getEmailAddressForSalesExecutive(salesExecutive);
    if (!email) {
      Logger.log("Email not found for " + salesExecutive);
      Browser.msgBox('Email not found for the sales executive.');
      return;
    }
  } else {
    email = 'sales@vishalcables.com'; // fallback email
  }

  sendDownloadLinkViaEmail(email, pdfUrl);

  Browser.msgBox('Your report is ready! An email with the download link has been sent. Also, you can copy-paste this URL to see the Report: ' + pdfUrl);
  Logger.log('PDF Created: ' + pdfUrl);
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd-MMM-yyyy");
}


function setDropdownFromColumnI() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inquiry Process"); // Change as needed
  var lastRow = sheet.getLastRow(); // Get last row with data

  if (lastRow < 2) return; // If no data, exit

  var columnIValues = sheet.getRange(5, 9, lastRow - 1, 1).getValues(); // Get all values from Column I
  var columnHCells = sheet.getRange(5, 8, lastRow - 1, 1); // Get all cells in Column H
  var columnHValidations = columnHCells.getDataValidations(); // Get dropdown validations in one go
  var columnHValues = columnHCells.getValues(); // Get current values in Column H

  var newValues = []; // To store updated values for batch processing

  for (var i = 0; i < columnIValues.length; i++) {
    var valueInI = columnIValues[i][0];
    var validation = columnHValidations[i][0];

    if (validation) {
      var dropdownValues = validation.getCriteriaValues()[0]; // Get dropdown list values
      if (dropdownValues.includes(valueInI)) {
        newValues.push([valueInI]); // Set new value for Column H
      } else {
        newValues.push([columnHValues[i][0]]); // Keep the existing value if I's value isn't in dropdown
      }
    } else {
      newValues.push([columnHValues[i][0]]); // Keep existing value if no dropdown
    }
  }

  columnHCells.setValues(newValues); // Batch update in one call (MUCH FASTER)
}


function onOpen() {
  setDropdownFromColumnI();

}

function testSendEmailOnFormSubmit() {
  // Create a mock event object that simulates the data passed on form submit
  var sheetName = 'Email Task Details';  // Replace with your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Simulate a row (for example, row 2) being updated (you can use any valid row)
  var row = 3;

  // Create a mock event object
  var e = {
    range: sheet.getRange(row, 1),  // Get the range for the row (starting from column 1)
    source: SpreadsheetApp.getActiveSpreadsheet(),
  };
Logger.log(row);
  // Manually trigger the sendEmailOnFormSubmit function with the mock event
 // sendEmailOnFormSubmit(e);
 sendEmailForTask(e)
}
function sendEmailForTask(e) {
  var sheetName = 'Email Task Details';  
  var sheet = e.source.getSheetByName(sheetName);
  var row = e.range.getRow();  
  
  var whatsappStatusCol = 8; 
  var doerNameCol = 7; 
  var taskCol = 5;
  var remarkCol = 6;
  var timestampCol = 1;

  Logger.log("Processing row: " + row);

  var timestamp = sheet.getRange(row, timestampCol).getValue();  
  var task = sheet.getRange(row, taskCol).getValue();  
  var whatsappStatus = sheet.getRange(row, whatsappStatusCol).getValue();  
  var doerName = sheet.getRange(row, doerNameCol).getValue();
  var remark = sheet.getRange(row, remarkCol).getValue();

  var plannedDate = "";
  if (timestamp && !isNaN(new Date(timestamp).getTime())) {  
    var date = new Date(timestamp);
    
    // Add 2 days
    date.setDate(date.getDate() + 2);
    
    // If the new date is Friday (getDay() === 5), skip to Saturday
    if (date.getDay() === 5) {
      date.setDate(date.getDate() + 1);
    }

    plannedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  var doerListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doer list");
  var contactRange = doerListSheet.getRange("A:B").getValues(); 
  var doerContact = "";

  for (var i = 0; i < contactRange.length; i++) {
    if (contactRange[i][0] === doerName) {
      Logger.log("Matching Doer Name: " + doerName);
      doerContact = contactRange[i][1]; 
      Logger.log("Doer Contact: " + doerContact);
      break; 
    }
  }

  if (doerContact) {
    var waMsg = "Hello " + doerName + ",\n\n" + 
                "You have been delegated the following task:\n\n" +
                "Task: " + task + "\n" +
                "Remark: " + remark + "\n" +
                "Planned Date: " + plannedDate + "\n" +
                "Task Assigner: By Email\n\n" +
                "Best regards,\nTeam";

    Logger.log("Sending WhatsApp message...");
    sendMessage(doerContact, waMsg);
    sheet.getRange(row, whatsappStatusCol).setValue('Sent');
  }
}

function sendMessage(contact,waMsg){
  //Logger.log(imageLink);
  var URL = "https://app.messageautosender.com/message/new/json";
  var whatSend = {
    "username":"vishalcables_mis",          // Your username
    "password":"MIS@1234",        // Your password
    "receiverMobileNo":contact,  // Receiver's phone number (you can add more numbers separated by commas)
    "message":[waMsg]  // Messages to send
    // Optional: Attach files (if needed)
   
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(whatSend)
  };

  var sendNow = "";
  try {
    // Send POST request
    sendNow = UrlFetchApp.fetch(URL, options);
    
    // Log the response from the API
    var responseCode = sendNow.getResponseCode();
    var responseBody = sendNow.getContentText();
    
    Logger.log("Response Code: " + responseCode);  // Log the status code
    Logger.log("Response Body: " + responseBody);  // Log the response content
    
    // Optionally, you can check if the API returned a successful response (e.g., 200 OK)
    if (responseCode == 200) {
      Logger.log("Message sent successfully!");
    } else {
      Logger.log("Error sending message: " + responseCode);
    }
  } catch (e) {
    // Log any errors that occur
    Logger.log("Error in sending message: " + e);
  }
}
function checkAndSendMessages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inquiry Process');
  var data = sheet.getDataRange().getValues(); // Get all data from sheet
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize time
  var contact = "9175527267";
 // var contact = "7058819446";
  let allTasksMessage = "Hello, Please find today's follow-up Details\n\n"; // Move greeting outside

  // Filter rows where M, O, or Q match today’s date
  var filteredData = data.filter(function(row, index) {
    if (index === 0) return false; // Skip header row
    var followUpDate = row[21] ? new Date(row[21]) : null; // Column V 
    var followUp1Date = row[23] ? new Date(row[23]) : null; // Column X
    var followUp2Date = row[25] ? new Date(row[25]) : null; // Column Z 
    var followUp3Date = row[27] ? new Date(row[27]) : null; // Column AB 

    if (followUpDate && !isNaN(followUpDate)) followUpDate.setHours(0, 0, 0, 0);
    if (followUp1Date && !isNaN(followUp1Date)) followUp1Date.setHours(0, 0, 0, 0);
    if (followUp2Date && !isNaN(followUp2Date)) followUp2Date.setHours(0, 0, 0, 0);
    if (followUp3Date && !isNaN(followUp3Date)) followUp3Date.setHours(0, 0, 0, 0);

    return ((followUpDate && followUpDate.getTime() === today.getTime()) ||
           (followUp1Date && followUp1Date.getTime() === today.getTime()) ||
            (followUp2Date && followUp2Date.getTime() === today.getTime()) ||
            (followUp3Date && followUp3Date.getTime() === today.getTime())) &&
           row[3]; // Ensure Customer Name is not empty (Column C - index 2)
  });

  // Process filtered rows
  let taskID=0;
  filteredData.forEach(function(row, index) {
    var rowIndex = data.indexOf(row) + 1; // Get actual row index
    var messages = [];
    var followUpDate = row[21] ? new Date(row[21]) : null;
    var followUp1Date = row[23] ? new Date(row[23]) : null;
    var followUp2Date = row[25] ? new Date(row[25]) : null;
    var followUp3Date = row[27] ? new Date(row[27]) : null;

 if (followUpDate && isSameDate(followUpDate, today)) {
      taskID++;
     messages.push("Task " + (taskID) + ":\n" + getMessageForRow(sheet, rowIndex, followUpDate));
    }
    if (followUp1Date && isSameDate(followUp1Date, today)) {
      taskID++;
     messages.push("Task " + (taskID) + ":\n" + getMessageForRow(sheet, rowIndex, followUp1Date));
    }
    if (followUp2Date && isSameDate(followUp2Date, today)) {
      taskID++;
      messages.push("Task " + (taskID) + ":\n" + getMessageForRow(sheet, rowIndex, followUp2Date));
    }
    if (followUp3Date && isSameDate(followUp3Date, today)) {
      taskID++;
      messages.push("Task " + (taskID) + ":\n" + getMessageForRow(sheet, rowIndex, followUp3Date));
    }

    if (messages.length > 0) {
      allTasksMessage += messages.join("\n\n") + "\n\n";
    }
  });

  if (allTasksMessage) {
    if(taskID>0)
    {
         Logger.log("Sending message: " + allTasksMessage);
    sendMessage(contact, allTasksMessage);
    }
  }
}

function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

function getMessageForRow(sheet, rowIndex, followUpDate) {
  var CustName = sheet.getRange(rowIndex, 3).getValue();  
  var CompName = sheet.getRange(rowIndex, 4).getValue();  


  var formattedDate = Utilities.formatDate(followUpDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  if (!CustName || !CompName) return ""; // Skip blank messages

  return "Customer Name: " + CustName + "\n" +
         "Company Name: " + CompName + "\n" +
         "Date: " + formattedDate ;
}
       





