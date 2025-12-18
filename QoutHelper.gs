function generateFilteredReportMonthWise() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inquiry Process');
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }
  var salesExecutive = sheet.getRange('Z3').getValue(); // Sales executive
  var fromDate = sheet.getRange('Z1').getValue(); // From date
  var toDate = sheet.getRange('Z2').getValue();   // To date
  var selectedType = sheet.getRange('Z4').getValue(); // Type filter
  
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
    var plannedDate1 = new Date(row[20]);
   // var dueDate1 = new Date(row[23]);
    //var dueDate2 = new Date(row[25]);
   // var dueDate3 = new Date(row[27]);
    var status = row[14];    // Column O
    var sentStatus = row[6]; // Column G

    plannedDate1.setHours(0, 0, 0, 0);
  //  dueDate1.setHours(0, 0, 0, 0);
  //  dueDate2.setHours(0, 0, 0, 0);
   // dueDate3.setHours(0, 0, 0, 0);

    var statusCondition = status !== "RECEIVED" && sentStatus === "Sent" && status !== "CLOSED";

    // Apply Type logic:
var typeCondition =
  !selectedType ||
  (selectedType === "Tender" && type === "Tender") ||
  (selectedType === "Live Inquiry" && type !== "Tender");
    if (salesExecutive) {
      return (salesExec === salesExecutive) &&
             typeCondition &&
             ((plannedDate1 >= fromDate && plannedDate1 <= toDate) 
              &&
             statusCondition);
    } else {
      return typeCondition &&
             ((plannedDate1 >= fromDate && plannedDate1 <= toDate)
              &&
             statusCondition);
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
 

  //sendDownloadLinkViaEmail(email, pdfUrl);

  Browser.msgBox('Your report is ready! You can copy-paste this URL to see the Report: ' + pdfUrl);
  Logger.log('PDF Created: ' + pdfUrl);
}
