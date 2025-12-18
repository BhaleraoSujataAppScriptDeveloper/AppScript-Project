function updateProjectDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Project Report"); // your dashboard sheet
  const poSheet = ss.getSheetByName("PO_Data");

  const project = sheet.getRange("E6").getValue().toString().trim();
  const fromDate = sheet.getRange("E4").getValue();
  const toDate = sheet.getRange("H4").getValue();

  const data = poSheet.getRange(2,1, poSheet.getLastRow()-1, poSheet.getLastColumn()).getValues();

  // Filter by Project and Date
  const filtered = data.filter(row => {
    const poDate = row[2]; // C = index 2
    const proj = row[18] ? row[18].toString().trim() : ""; // S = index 18
    return proj === project && poDate >= fromDate && poDate <= toDate;
  });

  if(filtered.length === 0){
    sheet.getRange("E8").setValue(0);
    sheet.getRange("G8").setValue(0);
    sheet.getRange("I8").setValue(0);
    sheet.getRange("D11:H").clearContent();
    return;
  }

  // Distinct Item count (Item Name + Item Code)
  const distinctItems = new Set(
    filtered.map(r => r[5].toString().trim().toLowerCase() + "|" + r[30].toString().trim().toLowerCase()) // F = 5, AC = 28
  );
  sheet.getRange("E8").setValue(distinctItems.size);

  // Total Quantity = sum of G (Quantity) column
  const totalQty = filtered.reduce((sum,r) => sum + (Number(r[6])||0),0);
  sheet.getRange("G8").setValue(totalQty);

  // Total Order Value = X column = 23
  const totalOrder = filtered.reduce((sum,r) => sum + (Number(r[23])||0),0);
  sheet.getRange("I8").setValue(totalOrder);

  // Item-wise summary table (D11:H)
  const summaryMap = {};
  filtered.forEach(r => {
    const itemName = r[5].toString().trim(); // F
    const itemCode = r[30].toString().trim(); // AC
    const qty = Number(r[6]||0); // G
    const sales = Number(r[23]||0); // X

    const key = itemName.toLowerCase() + "|" + itemCode.toLowerCase();
    if(!summaryMap[key]){
      summaryMap[key] = {item: itemName, code: itemCode, qty: qty, sales: sales};
    } else {
      summaryMap[key].qty += qty;
      summaryMap[key].sales += sales;
    }
  });

  const summaryArray = Object.values(summaryMap).map(o => [o.item, o.code, o.qty, o.sales, project]);
  sheet.getRange("D11:H").clearContent();
  if(summaryArray.length > 0){
    sheet.getRange(11,4,summaryArray.length,5).setValues(summaryArray);
  }
}

function getProjectList() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("PO_Data");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange("S2:S" + lastRow).getValues()
    .flat()
    .map(v => v && v.toString().trim())
    .filter(v => v!== "");
  Logger.log([...new Set(values)]);
  return [...new Set(values)];
}
function onChange_new() {
  var ss=SpreadsheetApp.getActiveSpreadsheet()
  var sheet=ss.getActiveSheet();
  var active=ss.getActiveCell();
  var row=active.getRow();
  var name=sheet.getSheetName();
  var col=active.getColumn();
  var val=active.getValue();
  var date=Utilities.formatDate(new Date(), "IST", "dd/MM/YYYY HH:mm:ss");
  if(val=="Done"){ts=sheet.getRange(row, col-1).getValue();
  	if(ts==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val=="Yes"){ts=sheet.getRange(row, col-1).getValue();//Actual Time column
  	if(ts==""){sheet.getRange(row, col-1).setValues(new Date())}}
  if(val=="No" ){ts=sheet.getRange(row, col-1).getValue();//Actual Time column
  	if(ts==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val==true ){ts=sheet.getRange(row, col-1).getValue();
  	if(ts==""){sheet.getRange(row, col-1).getValue(new Date())}}
}
function submitPOEntry(data) {
  try {
    Logger.log(data);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
    if (!sheet) throw new Error('Sheet "PO_Data" not found');

    if (!data.poNumber || !data.poDate || !data.items || data.items.length === 0) {
      throw new Error("PO Number, PO Date, and at least one item are required.");
    }
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    const allData = sheet.getDataRange().getValues();

    const headerFields = [
      formattedDate,
      data.poNumber,
      data.poDate,
      data.dispatchDate,
      data.buyerName,
      data.buyerAddress,
      data.buyerMobile,
      data.buyerPAN,
      data.buyerGST,
      data.consigneeName,
      data.paymentTerm,
      data.inspectionRequired,
      data.buyerEmail,
      data.project,
      data.consigneeAddress,
      data.consigneePAN,
      data.consigneeGST,
      data.consigneeMobile,
      data.orderValue,
      data.gstValue,
      data.totalAmount,
      data.finalAmount,
      data.creditDays,
      data.departmentName,
      data.orderType,
      data.salesExecutive
    ];

    const unitColumnIndex = 8; // AE column (Unit)
    const goColumnIndex = 9;   // GO number column

    const unitGoCounters = {}; // Track max GO number per unit from existing data
    const unitGoMap = {};      // New GO number per unit in current batch

    // Step 1: Read max GO number per unit from existing sheet
    allData.slice(1).forEach(row => {
      const unit = (row[unitColumnIndex] || "").toString().trim();
      const goValue = (row[goColumnIndex] || "").toString().trim();

      // Match format like: GO(U2)-2025-26/OG-2
      const match = /GO\((.*?)\)-2025-26\/OG-(\d+)/.exec(goValue);
      if (match) {
        const goNum = parseInt(match[2]);  // e.g., 2
        if (!unitGoCounters[unit] || goNum > unitGoCounters[unit]) {
          unitGoCounters[unit] = goNum;
        }
      }
    });
    // Step 2: Assign GO number per unit
    data.items.forEach(item => {
      const unit = item.unit.trim();
      if (!unitGoMap[unit]) {
        const nextGoNum = (unitGoCounters[unit] || 0) + 1;
        unitGoCounters[unit] = nextGoNum;
        const shortUnit = unit.replace(/\s+/g, '').replace(/^unit/i, 'U'); // "Unit 2" → "U2"
        unitGoMap[unit] = `GO(${shortUnit})-2025-26/OG-${nextGoNum}`;
      }
      item.goNumber = unitGoMap[unit];
    });

    // Step 3: Create rows
  const rows = data.items.map(item => {
  return [
    // First 5 header fields
    headerFields[0], // formattedDate
    headerFields[1], // poNumber
    headerFields[2], // poDate
    headerFields[3], // dispatchDate
    headerFields[4], // buyerName
    // Insert item-specific fields here
    item.itemName,
    item.qty,
    item.rate,
    item.unit,
    item.goNumber,
    // Continue remaining header fields
    headerFields[5],  // buyerAddress
    headerFields[6],  // buyerMobile
    headerFields[7],  // buyerPAN
    headerFields[8],  // buyerGST
    headerFields[9],  // consigneeName
    headerFields[10], // paymentTerm
    headerFields[11], // inspectionRequired
    headerFields[12], // buyerEmail
    headerFields[13], // project
    headerFields[14], // consigneeAddress
    headerFields[15], // consigneePAN
    headerFields[16], // consigneeGST
    headerFields[17], // consigneeMobile
    headerFields[18], // orderValue
    headerFields[19], // gstValue
    headerFields[20], // totalAmount
    headerFields[21], // finalAmount
    headerFields[22], // creditDays
    headerFields[23], // departmentName
    headerFields[24], // orderType

    // Finish with remaining item-specific fields
    item.itemCode,
    item.value,
    item.freight,
    item.freightTotalItem,
    data.freightType,
    data.priceVariation,
    data.contactPerson,
    data.salesExecutive
  ];
});
    // Step 4: Write to sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
// save data into client master
// Step 6: Store Buyer info in "Client Master" only if GST not present
const clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Client Master");
if (!clientSheet) throw new Error('Sheet "Client Master" not found');

const gstToCheck = data.buyerGST ? data.buyerGST.toString().trim() : "";

// Only check if GST is not blank
if (gstToCheck !== "") {
  const gstValues = clientSheet.getRange("H2:H" + clientSheet.getLastRow()).getValues().flat(); // GST column (H)
  const exists = gstValues.some(g => g.toString().trim() === gstToCheck);

  if (!exists) {
    const clientRow = [
      new Date(),                // Timestamp
      data.buyerName || "",
      data.buyerMobile || "",
      data.buyerEmail || "",
      data.buyerAddress || "",
      "",                        // City (blank)
      "",                        // State (blank)
      data.buyerGST || "",
      data.buyerPAN || ""
    ];
    clientSheet.appendRow(clientRow);
  }
}
// till here

//setUnitDropdownsForNewRowsOnly();

  // Step 5: If PO number contains "VCPL", log in Master sheet
if (data.poNumber && data.poNumber.toString().toUpperCase().includes("VCPL")) {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  
  if (!masterSheet) throw new Error('Sheet "Master" not found');

  const colC = masterSheet.getRange("C1:C").getValues(); // Entire column C
  let insertRow = colC.findIndex(row => row[0] === ""); // Find first empty row in C

  if (insertRow === -1) {
    insertRow = colC.length + 1; // If C is fully filled, go to next row
  } else {
    insertRow += 1; // Convert 0-based to 1-based row number
  }
  masterSheet.getRange(insertRow, 3).setValue(data.poNumber); // Column C = index 3
}
    return { status: "success", message: `PO submitted successfully with unit-wise GO Numbers.` };
  } catch (e) {
    logToSheet("Error: " + e.message);
    throw new Error("Submission failed: " + e.message);
  }
}

function uploadPODOCS(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO Files");
  const folder = DriveApp.getFolderById("1k1Fj0plo4NXgBwunLkUUsSWGT_4muq1u"); // Replace with your folder ID

  let fileUrl = "";

  if (data.fileData) {
    const blob = Utilities.newBlob(
      new Uint8Array(data.fileData.bytes),
      data.fileData.type,
      `${data.poNumber}_${data.fileData.name}`
    );
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  }
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm");
  // Store just PO number (Tender No) and file link
  sheet.appendRow([timestamp,data.poNumber, fileUrl]);
}
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : 'PO';
  if (page === 'GO') {
    const unitName = e.parameter.unit || '';
    const template = HtmlService.createTemplateFromFile('GO Form');
    template.unit = unitName;
    return template.evaluate()
      .setWidth(1200)
      .setHeight(700)
      .setTitle(unitName ? `Generate PDF - ${unitName}` : 'Generate PDF');
  } 
   if (page === 'INVOICE') {
    const template = HtmlService.createTemplateFromFile('Invoice');
    const unitName = e.parameter.unit || '';
    template.unit = unitName;
    return template.evaluate()
      .setWidth(1200)
      .setHeight(700)
      .setTitle(`Add Invoice - ${unitName}`);
  }
  // Default (PO Form)
  return HtmlService.createHtmlOutputFromFile('PO Form').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setWidth(1200)
    .setHeight(700)
    .setTitle('PO Form');
}
function logToSheet(data) {
  const sheetName = "Debug_Log";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["Timestamp", "Log"]);
  }
  const now = new Date();
  const logEntry = typeof data === "object" ? JSON.stringify(data) : String(data);

  sheet.appendRow([now, logEntry]);
}

function getPODetails(goNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues(); // includes header
  const headers = data[0];
  const goCol = headers.indexOf("GO Number");  // Column AF
  const poCol = headers.indexOf("PO Number");
  const buyerNameCol = headers.indexOf("Buyer Name");
  const buyerMobileCol = headers.indexOf("Buyer Mobile");
 
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][goCol]) === goNumber) {
      return {
        poNumber: data[i][poCol] || "",
        buyerName: data[i][buyerNameCol] || "",
        buyerMobile: data[i][buyerMobileCol] || ""
      };
    }
  }
  return null;
}

function testSubmit() {
  const data = {
    poNumber: "PO123",
    poDate: "2024-05-22",
    dispatchDate: "2024-05-25",
    buyerName: "Test Buyer",
    buyerAddress: "123 Street",
    buyerMobile: "9876543210",
    buyerPAN: "ABCDE1234F",
    buyerGST: "22ABCDE1234F1Z5",
    consigneeName: "Test Consignee",
    paymentTerm: "Advance",
    inspectionRequired: "Yes",
    buyerEmail: "buyer@example.com",
    project: "Project Alpha",
    consigneeAddress: "456 Road",
    consigneePAN: "WXYZE9876L",
    consigneeGST: "22WXYZE9876L1Z9",
    consigneeMobile: "9876501234",
    orderValue: 10000,
    gstValue: 1800,
    freight: 200,
    totalAmount: 12000,
    items: [
      { itemName: "Item A", qty: 10, rate: 100, value: 1000, unit: "pcs" },
      { itemName: "Item B", qty: 5, rate: 200, value: 1000, unit: "pcs" }
    ]
  };
  submitPOEntry(data);
}
function saveInvoiceData(data) {
  const folderId = '1cRQeBEL80winc3GpgAq7fCB4Wl9dGxno'; // 🔁 Replace with your Drive folder ID
  const folder = DriveApp.getFolderById(folderId);

  // Convert base64 to blob and upload to Drive
  const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.contentType, data.fileName);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();

  // Access the target sheet tab
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice Details");
  if (!sheet) {
    return "❌ 'InvoiceDetails' sheet not found.";
  }

  // Prepare the row data
  const newRow = [
    new Date(),             // Timestamp
    data.poNumber,
    data.goNumber,
    data.buyerName,
    data.buyerMobile,
    data.altMobile,
    data.dispatchQty,
    data.invoiceNo,
    data.invoiceDate,
    data.amount,
    data.fileName,
    fileUrl,
    data.driverContact,
    data.status
  ];
  // Append to the next available row
 sheet.appendRow(newRow);
 const lastRow = sheet.getLastRow();
 sendInvoiceToClient(lastRow); // ✅ pass row number
 return "✅ Invoice saved successfully to 'InvoiceDetails'!";
}
let imageLink = '';
function sendInvoiceToClient(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice Details');
  const PONumberCol = 2;
  const InvoiceNumberCol = 8;
  const fileCol = 12;
  const buyerNameCol = 4;
  const InvoiceDateCol = 9;
  const AmountCol = 10;
  const buyerContactCol = 5;
  const AltContactCol=6;
  const DriverMobileCol=13;
  const poNumber = sheet.getRange(row, PONumberCol).getValue();
  const invoiceNumber = sheet.getRange(row, InvoiceNumberCol).getValue();
  const buyerName = sheet.getRange(row, buyerNameCol).getValue();
  const invoiceDate = sheet.getRange(row, InvoiceDateCol).getValue();
  const amount = sheet.getRange(row, AmountCol).getValue();
  const file = sheet.getRange(row, fileCol).getValue();
  const buyerContact = sheet.getRange(row, buyerContactCol).getValue();
 // const buyerContact = "7058819446";
  const altContact = sheet.getRange(row, AltContactCol).getValue();
  const driverMobile = sheet.getRange(row, DriverMobileCol).getValue();
  const formattedInvoiceDate = Utilities.formatDate(new Date(invoiceDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
  let imageLink = "";
  if (file) {
    imageLink = getDirectDownloadUrlFromShareableLink(file);
  }

  var waMsg = 
  "Dear " + buyerName + ",\n\n" +
  "Thank you for choosing  – Vishal Cables Pvt. Ltd. \n\n" +
  "This is an automated message to inform you that your order has been successfully dispatched. Please find your Invoice details and Invoice copy below:\n\n" +
  "🔹 PO Number: " + poNumber + "\n" +
  "🔹 Invoice Number: " + invoiceNumber + "\n" +
  "🔹 Invoice Date: " + formattedInvoiceDate + "\n" +
  "🔹 Amount: ₹" + amount + "\n\n" +
    (driverMobile ? "🔹 Driver's Contact: " + driverMobile + "\n\n" : "");
 
  "For any queries or assistance, feel free to contact us.\n\n" +
  "Warm regards,\nVishal Cables Pvt Ltd.";


  Logger.log("Sending WhatsApp message...");
  if (buyerContact && /^[6-9]\d{9}$/.test(buyerContact))
  {
  sendMessage(buyerContact, imageLink, waMsg); // ✅ send
  sheet.getRange(row, 15).setValue('Buyer Sent'); // status
  }
 if (altContact && /^[6-9]\d{9}$/.test(altContact))
  {
    sendMessage(altContact, imageLink, waMsg); // ✅ send
    sheet.getRange(row, 16).setValue('Alt Sent'); // status
  }

}

function getDirectDownloadUrlFromShareableLink(shareableLink) {
  // Extract the file ID from the shareable link
  var fileId = extractFileIdFromUrl(shareableLink);
  
  if (fileId) {
    // Create the direct download URL
    return "https://drive.google.com/uc?export=download&id=" + fileId;
  } else {
    Logger.log("Invalid link or file ID could not be extracted.");
    return null;
  }
}
function extractFileIdFromUrl(url) {
  Logger.log("URL:  - " + url);  // Log the input URL for debugging
  var fileId = null;
  
  // Update the regex to capture the file ID after "?id="
  var regex = /(?:open\?id=|d\/)([a-zA-Z0-9_-]+)/;
  var matches = url.match(regex);
  
  if (matches && matches[1]) {
    fileId = matches[1];  // Extract the file ID from the URL
  }
  
  Logger.log("File ID: " + fileId);  // Log the extracted file ID for debugging
  return fileId;
}

function sendMessageOLD(contact,imageLink,waMsg){
  Logger.log(imageLink);
  var URL = "https://app.messageautosender.com/message/new/json";
 
  var whatSend = {
    "username":"vishalcables_cc",          // Your username
    "password":"MIS@1234",        // Your password
    "receiverMobileNo":contact,  // Receiver's phone number (you can add more numbers separated by commas)
    "message":[waMsg],  // Messages to send
    // Optional: Attach files (if needed)
    "filePathUrl":[imageLink]
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
function sendMessage(contact, imageLink, waMsg) {
  const payload = {
    contact: contact,
    imageLink: imageLink,
    waMsg: waMsg
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  const webAppUrl = "https://script.google.com/macros/s/AKfycbwc1IdrG5xFmEbKY3svX0tIFVpyNi3IrTj_pEHwJFc1XRGC3X0wfxMGkspno4W3ogW_/exec"; // e.g., https://script.google.com/macros/s/xxxx/exec
  UrlFetchApp.fetch(webAppUrl, options);
}

function generatePendingOrdersReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poDataSheet = ss.getSheetByName("PO_Data");
  const poSummarySheet = ss.getSheetByName("PO Summary");
  const invoiceSheet = ss.getSheetByName("Invoice Details");
  const reportSheet = ss.getSheetByName("Pending Sales Orders-Unit I") || ss.insertSheet("Pending Sales Orders-Unit I");

  // 🔹 Clear old data but keep header
  if (reportSheet.getLastRow() > 1) {
    reportSheet.getRange(4, 1, reportSheet.getLastRow() - 2, 11).clearContent();
  }
  reportSheet.getRange("A:A").breakApart();

  // 1️⃣ Build PO → Pending status map
  const summaryData = poSummarySheet.getRange("C2:M" + poSummarySheet.getLastRow()).getValues();
  const pendingPOs = new Set();
  summaryData.forEach(r => {
    const po = r[0];       // C = PO Number
    const status = r[9];   // M = index 9 (0-based from C)
    if (status === "Pending") pendingPOs.add(po);
  });

  // 2️⃣ Build PO → list of Go Numbers
  const invoiceData = invoiceSheet.getRange("B2:C" + invoiceSheet.getLastRow()).getValues();
  const poGoMap = {}; // PO → array of Go Numbers
  invoiceData.forEach(r => {
    const po = r[0];
    const goNumber = r[1] ? r[1].toString().toUpperCase().trim() : "";
    if (!poGoMap[po]) poGoMap[po] = [];
    poGoMap[po].push(goNumber);
  });

  // 3️⃣ Loop PO_Data and filter Unit 2 items
  const poData = poDataSheet.getRange("A2:Z" + poDataSheet.getLastRow()).getValues();
  let output = [];

  poData.forEach(r => {
    const poNumber = r[1]; // B = PO Number
    const unit = r[8];     // I = Unit
    
    if (!pendingPOs.has(poNumber)) return;

    const goNumbers = poGoMap[poNumber] || [];

    if (goNumbers.length > 0) {
      // ✅ Invoice exists → decide by GO Number
      if (goNumbers.some(g => g.includes("U1")) && unit === "Unit 1") {
        output.push([r[4], r[2], r[1], r[4], r[5], r[7], r[6], "", "", ""]);
      }
    } else {
      // 🚨 No invoice yet → fall back to PO_Data unit column
      if (unit === "Unit 1") {
        output.push([r[4], r[2], r[1], r[4], r[5], r[7], r[6], "", "", ""]);
      }
    }
  });

  // 4️⃣ Write to report
  if (output.length > 0) {
    reportSheet.getRange(4, 1, output.length, output[0].length).setValues(output);

    for (let i = 3; i <= output.length + 2; i++) {
      reportSheet.getRange(i, 8).setFormula(`=F${i}*G${i}`);
      reportSheet.getRange(i, 9).setFormula(`=H${i}*0.18`);
      reportSheet.getRange(i, 10).setFormula(`=H${i}+I${i}`);
      reportSheet.getRange(i, 11).insertCheckboxes();
    }

    // 🔹 Buyer merge + totals
    let lastRow = reportSheet.getLastRow();
    let buyerCol = reportSheet.getRange(4, 1, lastRow - 2, 1).getValues();
    let startRow = 4;
    for (let i = 5; i <= lastRow + 1; i++) {
      if (i > lastRow || buyerCol[i - 3][0] !== buyerCol[i - 4][0]) {
        let endRow = i - 1;
        let total = reportSheet.getRange(startRow, 10, endRow - startRow + 1, 1).getValues().flat()
          .reduce((a, b) => a + (Number(b) || 0), 0);
        if (endRow - startRow + 1 >= 1) {
          reportSheet.getRange(startRow, 1, endRow - startRow + 1, 1).mergeVertically()
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontWeight("bold")
            .setValue(`${buyerCol[startRow - 3][0]} (${total})`);
        }
        startRow = endRow + 1;
        lastRow = reportSheet.getLastRow();
        buyerCol = reportSheet.getRange(3, 1, lastRow - 2, 1).getValues();
      }
    }
  }
}

function generatePendingOrdersReportUnit2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poDataSheet = ss.getSheetByName("PO_Data");
  const poSummarySheet = ss.getSheetByName("PO Summary");
  const invoiceSheet = ss.getSheetByName("Invoice Details");
  const reportSheet = ss.getSheetByName("Pending Sales Orders-Unit II") || ss.insertSheet("Pending Sales Orders-Unit II");

  // 🔹 Clear old data but keep header
  if (reportSheet.getLastRow() > 1) {
    reportSheet.getRange(4, 1, reportSheet.getLastRow() - 2, 11).clearContent();
  }
  reportSheet.getRange("A:A").breakApart();

  // 1️⃣ Build PO → Pending status map
  const summaryData = poSummarySheet.getRange("C2:M" + poSummarySheet.getLastRow()).getValues();
  const pendingPOs = new Set();
  summaryData.forEach(r => {
    const po = r[0];       // C = PO Number
    const status = r[9];   // M = index 9 (0-based from C)
    if (status === "Pending") pendingPOs.add(po);
  });

  // 2️⃣ Build PO → list of Go Numbers
  const invoiceData = invoiceSheet.getRange("B2:C" + invoiceSheet.getLastRow()).getValues();
  const poGoMap = {}; // PO → array of Go Numbers
  invoiceData.forEach(r => {
    const po = r[0];
    const goNumber = r[1] ? r[1].toString().toUpperCase().trim() : "";
    if (!poGoMap[po]) poGoMap[po] = [];
    poGoMap[po].push(goNumber);
  });

  // 3️⃣ Loop PO_Data and filter Unit 2 items
  const poData = poDataSheet.getRange("A2:Z" + poDataSheet.getLastRow()).getValues();
  let output = [];

  poData.forEach(r => {
    const poNumber = r[1]; // B = PO Number
    const unit = r[8];     // I = Unit
    if (!pendingPOs.has(poNumber)) return;
    const goNumbers = poGoMap[poNumber] || [];
    if (goNumbers.length > 0) {
      // ✅ Invoice exists → decide by GO Number
      if (goNumbers.some(g => g.includes("U2")) && unit === "Unit 2") {
        output.push([r[4], r[2], r[1], r[4], r[5], r[7], r[6], "", "", ""]);
      }
    } else {
      // 🚨 No invoice yet → fall back to PO_Data unit column
      if (unit === "Unit 2") {
        output.push([r[4], r[2], r[1], r[4], r[5], r[7], r[6], "", "", ""]);
      }
    }
  });
  // 4️⃣ Write to report
  if (output.length > 0) {
    reportSheet.getRange(4, 1, output.length, output[0].length).setValues(output);
    for (let i = 4; i <= output.length + 2; i++) {
      reportSheet.getRange(i, 8).setFormula(`=F${i}*G${i}`);
      reportSheet.getRange(i, 9).setFormula(`=H${i}*0.18`);
      reportSheet.getRange(i, 10).setFormula(`=H${i}+I${i}`);
      reportSheet.getRange(i, 11).insertCheckboxes();
    }
    // 🔹 Buyer merge + totals
    let lastRow = reportSheet.getLastRow();
    let buyerCol = reportSheet.getRange(4, 1, lastRow - 2, 1).getValues();
    let startRow = 4;
    for (let i = 5; i <= lastRow + 1; i++) {
      if (i > lastRow || buyerCol[i - 3][0] !== buyerCol[i - 4][0]) {
        let endRow = i - 1;
        let total = reportSheet.getRange(startRow, 10, endRow - startRow + 1, 1).getValues().flat()
          .reduce((a, b) => a + (Number(b) || 0), 0);
        if (endRow - startRow + 1 >= 1) {
          reportSheet.getRange(startRow, 1, endRow - startRow + 1, 1).mergeVertically()
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontWeight("bold")
            .setValue(`${buyerCol[startRow - 3][0]} (${total})`);
        }
        startRow = endRow + 1;
        lastRow = reportSheet.getLastRow();
        buyerCol = reportSheet.getRange(3, 1, lastRow - 2, 1).getValues();
      }
    }
  }
}

function showLoader() {
  const html = HtmlService.createHtmlOutputFromFile('Loading')
      .setWidth(250)
      .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Please wait');
}
function showLoaderUnit2() {
  const html = HtmlService.createHtmlOutputFromFile('LoadingU2')
      .setWidth(250)
      .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Please wait');
}
function getPODataByBuyer(buyerName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === buyerName) { // Column E = index 4
      const rowData = {};
      headers.forEach((header, idx) => {
        rowData[header] = data[i][idx];
      });
      results.push(rowData);
    }
  }
  return results;
}


function hideFormDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Invoice Details"); // replace with your sheet name
  if (dataSheet) {
    dataSheet.hideSheet(); // Hides the sheet from normal users
  }
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // BMP Formulas menu
 const bmpMenu= ui.createMenu('BMP Formulas')

      .addItem('⚙️ Initial Setup (Run First Time)', 'initialSetup')

      .addSeparator()

      .addSubMenu(ui.createMenu('📊 FMS to DB Conversion')

          .addItem('1️⃣ Setup Looker Studio Sheets', 'createSetup')

          .addItem('2️⃣ Create Form with Prefilled URL', 'createFormWithPrefilledURL')

          .addItem('3️⃣ Convert FMS to DB Format', 'convertFMStoDBWithDoers'))

      .addSeparator()

      .addItem('📥 Assisted ImportRange', 'importRangeFormula')

      .addSeparator()

      .addSubMenu(ui.createMenu('📐 FMS Formulas')

          .addItem('TAT with Working Hours', 'plannedwwh')

          .addItem('TAT in Days', 'plannedindays')

          .addItem('T-x Formula', 'plannedlead')

          .addItem('Specific Time', 'specificTime')

          .addItem('Show Planned Only When Status is NO', 'tatifno')

          .addItem('Show Planned Only When Status is YES', 'tatifyes')

          .addItem('Set Actual Time', 'actualTime')

          .addItem('Time Delay Formula', 'timeDelay')

          .addItem('Vlookup Wizard', 'vlookupFormula'))
  // Purchase Order menu
  const poMenu = ui.createMenu('O2D Menus')
  //  .addItem('PO Form', 'showPOForm')
   // .addSeparator()
    .addItem('Download GO - Unit 1', 'showPOFormUnit1')
    .addSeparator()
    .addItem('Download GO - Unit 2', 'showPOFormUnit2')
    .addSeparator()
    .addItem('Add Invoice - Unit 1', 'submitInvoiceUnit1')
    .addSeparator()
    .addItem('Add Invoice - Unit 2', 'submitInvoiceUnit2');
    bmpMenu.addToUi();
    poMenu.addToUi();
}
function showPOFormUnit1() {
  showPOFormByUnit("Unit 1");
}
function showPOFormUnit2() {
  showPOFormByUnit("Unit 2");
}
function submitInvoiceUnit1()
{
  addInvoiceUnit("Unit 1");
}
function submitInvoiceUnit2()
{
  addInvoiceUnit("Unit 2");
}

function getNextVCPLPO() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat().filter(String);
  const prefix = "VCPL/2025-26/";
  
  // Extract last number
  const numbers = data
    .filter(po => po.startsWith(prefix))
    .map(po => parseInt(po.replace(prefix, ""), 10))
    .filter(n => !isNaN(n))
    .sort((a, b) => b - a); // Descending

  const lastNumber = numbers.length ? numbers[0] : 0;
  const nextNumber = String(lastNumber + 1).padStart(3, "0");
  return prefix + nextNumber;
}

function setUnitDropdownsForOnlyNewRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const optionsU1 = ["Done"];
  const optionsU2 = ["Done"];

  const startRow = 7; // Start scanning from row 7
  for (let row = startRow; row <= lastRow; row++) {
    const dValue = sheet.getRange(row, 4).getValue(); // Column D
    const zCell = sheet.getRange(row, 26); // Column Z
    const adCell = sheet.getRange(row, 30); // Column AD
    const zHasValidation = zCell.getDataValidation() !== null;
    const adHasValidation = adCell.getDataValidation() !== null;

    // Skip if Z or AD already has validation
    if (zHasValidation || adHasValidation) continue;

    if (typeof dValue !== "string" || dValue === "") continue;

    // Clear previous validations only
    zCell.clearDataValidations();
    adCell.clearDataValidations();

    if (dValue.includes("U1")) {
      const ruleZ = SpreadsheetApp.newDataValidation()
        .requireValueInList(optionsU1, true)
        .setAllowInvalid(false)
        .build();
      zCell.setDataValidation(ruleZ);
    } else if (dValue.includes("U2")) {
      const ruleAD = SpreadsheetApp.newDataValidation()
        .requireValueInList(optionsU2, true)
        .setAllowInvalid(false)
        .build();
      adCell.setDataValidation(ruleAD);
    }
  }
}
function addInvoiceUnit(unitName) {
  const template = HtmlService.createTemplateFromFile('Invoice');
  template.unit = unitName;
  const html = template.evaluate().setWidth(1000).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, `Add Invoice - ${unitName}`);
}
function CreatePackingForm_UnitI() {
  const template = HtmlService.createTemplateFromFile('PackingList_Unit I');
  template.unit = "UNITI";  // 👈 value set kar di
  const html = template.evaluate().setWidth(1000).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, `Create Packing List Unit I`);
}
function CreatePackingForm_UnitII() {
  const template = HtmlService.createTemplateFromFile('PackingList_Unit II');
  template.unit = "UNITII";  // 👈 value set kar di
  const html = template.evaluate().setWidth(1000).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, `Create Packing List Unit II`);
}
function showPOFormByUnit(unitName) {
  const template = HtmlService.createTemplateFromFile('GO Form');
  template.unit = unitName;
  const html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, `Generate PDF - ${unitName}`);
}
function populatePONumbers() {
  var form = FormApp.openById('1PmUHr8N7yOevB_E46T4xK8M_uSZGoz3Fe5wgLDXjXYM');  // Replace with your Form ID
  var sheet = SpreadsheetApp.openById('1ZCjxuaee-VOTYVIv83iACqV2XO3ZKLtbQup_m2MJpJw');  // Replace with your Sheet ID
  var range = sheet.getSheetByName('PO_Data').getRange('B7:B');  // Adjust range as needed
  var values = range.getValues();
  
  // Search for the dropdown question by its title
  var questions = form.getItems(FormApp.ItemType.LIST);
  var item = null;
  Logger.log(questions);
  // Loop through all questions to find the one with the title you're looking for
  for (var i = 0; i < questions.length; i++) {
    if (questions[i].getTitle().toLowerCase() === "po number") {  // Replace with your actual question title
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
function getUniquePOs(unit) {
  Logger.log(unit);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  const pos = new Set();

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === unit) {
      pos.add(data[i][9]); // PO Number from column B
    }
  }
  return [...pos].sort();
}
function test1() {
  const result = getPODetailsByUnit("Unit 1");
  Logger.log("Result: " + JSON.stringify(result));
}
function getUniqueGOs(unit) {
  const poDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const poSummarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO Summary");

  const poData = poDataSheet.getDataRange().getValues();
  const poSummary = poSummarySheet.getDataRange().getValues();

  // Create a Set of completed PO numbers from PO Summary
  const completedPOs = new Set();
  for (let i = 1; i < poSummary.length; i++) {
    if (poSummary[i][11] != "Completed") {  // Column M = index 12
      completedPOs.add(poSummary[i][2]);     // Column C = index 2 (PO Number)
    }
  }
  // Now collect GO numbers from PO_Data only if unit matches and PO is completed
  const goNumbers = new Set();
  for (let i = 1; i < poData.length; i++) {
    const poNumber = poData[i][1];   // Column B = index 1
    const currentUnit = poData[i][8]; // Column I = index 8
    const goNumber = poData[i][9];   // Column J = index 9

    if (currentUnit === unit && completedPOs.has(poNumber)) {
      goNumbers.add(goNumber);
    }
  }
//Logger.log([...goNumbers].sort());
  return [...goNumbers].sort();
}

function getPODetailsByUnit(unit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PO_Data');
  const data = sheet.getDataRange().getValues();
  const poList = [];
  const seenPOs = new Set(); // to store unique PO numbers

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === unit) {
      const poNumber = String(data[i][1]).trim();
      if (!seenPOs.has(poNumber)) {
        seenPOs.add(poNumber);
        poList.push({
          poDate: formatDate(data[i][2]),
          poNumber: poNumber,
          buyerName: String(data[i][4]).trim(),
          goNumber: String(data[i][9]).trim()
        });
      }
    }
  }
  return poList;
}

// Utility to format date safely
function formatDate(dateVal) {
  if (Object.prototype.toString.call(dateVal) === '[object Date]') {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd-MM-yyyy");
  }
  return String(dateVal);
}

function getPODetailsByUnittest(unit) {
  return [
    {
      poDate: "10-06-2025",
      poNumber: "PO123",
      buyerName: "ABC Corp",
      goNumber: "GO001"
    }
  ];
}

function getBuyerDetails(buyerName) {
 // const buyerName="PARTH ENTERPRISES";
  Logger.log("Buyer Name received: " + buyerName);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues(); // includes headers in row 1
  const headers = data[0];
  const buyerIndex = headers.indexOf("Buyer Name");
  Logger.log("Buyer Index: " + buyerIndex);
  if (buyerIndex === -1) return null;

  const matchedRow = data.slice(1).find(row =>
    row[buyerIndex] && row[buyerIndex].toString().trim().toLowerCase() === buyerName.toLowerCase().trim()
  );
  if (!matchedRow) return null;
  Logger.log(matchedRow);
  const details = {};
  headers.forEach((header, i) => {
    details[header] = matchedRow[i];
  });
  return details; // returns object like { "Buyer Name": ..., "Address": ..., ... }
}

function generatePDFForPO(poNumber, unit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFMS = ss.getSheetByName("FMS");
  if (!sheetFMS) return;
  const startRow = 7; // data starts from row 7
  const lastRow = sheetFMS.getLastRow();
  const poNumbers = sheetFMS.getRange(7, 4, lastRow - 1).getValues(); // Column D (PO numbers)

  for (let i = 0; i < poNumbers.length; i++) {
    if (String(poNumbers[i][0]).trim() === poNumber) {
      const row = startRow + i; // because data starts from row 2
      if (poNumber.includes("U1")) {
        sheetFMS.getRange(row, 26).setValue("Done"); // Column Z = 26
      } else if (poNumber.includes("U2")) {
        sheetFMS.getRange(row, 30).setValue("Done"); // Column AD = 30
      }
      break; // exit after first match
    }
  }

  const TEMPLATE_ID = '1ft6olX22SCtmxBCjv3qdpp1UqsZ0ogtyje4lpmRo1lU';
  const DEST_FOLDER_ID = '1qw0GRHM6vNFVGgLhHrkfir69OohwN7qN';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  const filteredRows = data.slice(1).filter(row => row[9] === poNumber && row[8] === unit);
  if (filteredRows.length === 0) return "No matching data.";

  const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`PO_${poNumber}_${unit}`, DriveApp.getFolderById(DEST_FOLDER_ID));
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  const row = filteredRows[0]; // Assume common fields from first matching row
  const todaysDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  const poDate = Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd-MMM-yyyy");
  const goDate = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd-MMM-yyyy");
  let deliveryDate = "";
if (row[3] instanceof Date && !isNaN(row[3].getTime())) {
  deliveryDate = Utilities.formatDate(row[3], Session.getScriptTimeZone(), "dd-MMM-yyyy");
} else {
  deliveryDate = "";  // or any default value you want when date is missing
}

  // Fill in static fields
  body.replaceText('{{TODAYS_DATE}}', goDate);
  body.replaceText('{{UNIT}}', unit);
  body.replaceText('{{GO_NUMBER}}', row[9]);
  body.replaceText('{{BUYER_NAME}}', row[4]);
  body.replaceText('{{BUYER_ADDRESS}}', row[10]);
  body.replaceText('{{BUYER_MOBILE}}', row[11]);
  body.replaceText('{{BUYER_EMAIL}}', row[17]);
  body.replaceText('{{CONSIGNEE_NAME}}', row[14]);
  body.replaceText('{{CONSIGNEE_ADDRESS}}', row[19]);
  body.replaceText('{{CONSIGNEE_MOBILE}}', row[22]);
  body.replaceText('{{PROJECT}}', row[18]);
  body.replaceText('{{C_PAN}}', row[20]);
  body.replaceText('{{C_GST}}', row[21]);
  body.replaceText('{{PO_DATE}}', poDate);
  body.replaceText('{{PAN}}', row[12]);
  body.replaceText('{{GST_NO}}', row[13]);
  body.replaceText('{{PAYMENT_MOTHOD}}', row[15]);
  body.replaceText('{{ISP}}', row[16]);
  body.replaceText('{{DELIVERY_DATE}}', deliveryDate);
  body.replaceText('{{PURCHASE_ORDER}}', row[1]);
 // body.replaceText('{{FREIGHT}}', row[20]);
  body.replaceText('{{CREDIT_DAYS}}', row[27]);
  body.replaceText('{{INST_TYPE}}', row[28]);
  body.replaceText('{{FREIGHT_TYPE}}', row[34]);
  body.replaceText('{{PRICE_VAR}}', row[35]);
  body.replaceText('{{REFERENCE}}',row[37]);
  
  // --- Calculate Order Value, GST, Freight, Total ---
  let orderValue = 0;
  let freightValueUnit=0;
  filteredRows.forEach(item => {
    const qty = parseFloat(item[6]) || 0;
    const rate = parseFloat(item[7]) || 0;
    const freight=parseFloat(item[32])||0;
    orderValue += qty * rate;
    freightValueUnit+=freight;
  });

 // const freightValue =row[20];
  const totalValue = orderValue  + freightValueUnit;
  const gstValue = totalValue*0.18;
  const finalValue =Math.round(totalValue+gstValue);

  body.replaceText('{{ORDER_VALUE}}', orderValue.toFixed(2));
  body.replaceText('{{FREIGHT}}', freightValueUnit.toFixed(2));
  body.replaceText('{{TOTAL_AMOUNT}}', totalValue.toFixed(2));
  body.replaceText('{{GST}}', gstValue.toFixed(2));
  body.replaceText('{{FINAL_AMOUNT}}', finalValue.toFixed(2));

  // --- Insert Item Table ---
  const found = body.findText('{{ITEM_TABLE}}');
  if (found) {
    const element = found.getElement();
    const parent = element.getParent();
    parent.asParagraph().clear(); // Remove placeholder

    const table = parent.getParent().insertTable(
      parent.getParent().getChildIndex(parent) + 1,
      [['No.', 'Items', 'Code', 'Qty', 'Rate', 'Value','Freight / Mtr','Total Freight']]
    );
    filteredRows.forEach((item, index) => {
      const row = table.appendTableRow();
      row.appendTableCell(String(index + 1));
      row.appendTableCell(item[5]);
      row.appendTableCell(item[30]);
      row.appendTableCell(Number(item[6]));
      row.appendTableCell(Number(item[7]).toFixed(2));
      row.appendTableCell(Number(item[31]).toFixed(2));
      row.appendTableCell(Number(item[32]).toFixed(2));
      row.appendTableCell(Number(item[33]).toFixed(2));
    });
      setColumnWidths(table);
  }
  doc.saveAndClose();
  Utilities.sleep(3000);
  const file = DriveApp.getFileById(doc.getId());
  return file.getUrl();
}
function setColumnWidths(table) {
  const widths = [40, 150, 50, 50, 60, 60, 60, 60]; // Adjust widths as needed
  const numRows = table.getNumRows();
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    for (let c = 0; c < widths.length; c++) {
      row.getCell(c).setWidth(widths[c]);
    }
  }
}
function checkPONumberExists(poNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FMS");
  const data = sheet.getRange("C7:C" + sheet.getLastRow()).getValues().flat().filter(String);
  return data.includes(poNumber);
}
function checkInvoiceExist(invNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment - Dispatch");
  const data = sheet.getRange("E7:E" + sheet.getLastRow()).getValues().flat().filter(String);
  return data.includes(invNumber);
}
function getItemNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const values = sheet.getRange("A2:A" + sheet.getLastRow()).getValues(); // Skip header
  return values.flat().filter(String); // Flatten and remove empty values
}
function getItemCode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const values = sheet.getRange("B2:B" + sheet.getLastRow()).getValues(); // Skip header
  return values.flat().filter(String); // Flatten and remove empty values
}
function showPOForm() {
  const html = HtmlService.createHtmlOutputFromFile("PO Form")
    .setWidth(1500)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, "PO Entry Form");
}
function createTrigger() { 
  removeTrigger()
  removeTrigger()
  removeTrigger()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange_new').forSpreadsheet(ss).onChange().create();  
}
function removeTrigger() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}

function findAndEditFormSheet() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ui = SpreadsheetApp.getUi();

 

  // Get all sheets in the spreadsheet

  var sheets = ss.getSheets();

  var formResponseSheet = null;

 

  // Look for the sheet with matching headers

  for (var i = 0; i < sheets.length; i++) {

    var sheet = sheets[i];

   

    // Skip hidden sheets and the form_config sheet

    if (sheet.isSheetHidden() || sheet.getName() === 'form_config') {

      continue;

    }

   

    // Check if sheet has at least 3 columns

    if (sheet.getLastColumn() >= 3) {

      // Get the first 3 header values

      var headers = sheet.getRange(1, 1, 1, 3).getValues()[0];

     

      // Check if headers match: Timestamp, Unique Key, Step Code

      if (headers[0] === 'Timestamp' &&

          headers[1] === 'Unique Key' &&

          headers[2] === 'Step Code') {

        formResponseSheet = sheet;

        break;

      }

    }

  }

 

  // If no matching sheet found

  if (!formResponseSheet) {

    ui.alert('Error: No form response sheet found with the expected headers (Timestamp, Unique Key, Step Code).');

    return null;

  }

 

  Logger.log('Found form response sheet: ' + formResponseSheet.getName());

 

  // Rearrange columns: Move Timestamp (column 1) to the end (column 3)

  // This will result in: Unique Key, Step Code, Timestamp

 

  // Method: Move column A (Timestamp) to after column C

  formResponseSheet.moveColumns(formResponseSheet.getRange("A:A"), 3);

  formResponseSheet.moveColumns(formResponseSheet.getRange("C:C"), 2);

 

  Logger.log('Columns rearranged successfully');

 

  // Rename the sheet to "Data"

  // Check if a sheet named "Data" already exists

  var existingDataSheet = ss.getSheetByName('Data');

  if (existingDataSheet && existingDataSheet.getSheetId() !== formResponseSheet.getSheetId()) {

    // If "Data" sheet exists and it's not the current sheet, rename it first

    var timestamp = new Date().getTime();

    existingDataSheet.setName('Data_old_' + timestamp);

    Logger.log('Existing "Data" sheet renamed to: Data_old_' + timestamp);

  }

 

  formResponseSheet.setName('Data');

  Logger.log('Sheet renamed to: Data');

 

  // Get the prefilled URL from form_config sheet

  var configSheet = ss.getSheetByName('form_config');

 

  if (!configSheet) {

    ui.alert('Warning: form_config sheet not found. Sheet has been renamed but cannot retrieve prefilled URL.');

    return null;

  }

 

  // Get the last row in form_config (most recent form)

  var lastRow = configSheet.getLastRow();

 

  if (lastRow < 2) {

    ui.alert('Warning: No form configuration found in form_config sheet.');

    return null;

  }

 

  // Get the prefilled URL from the last row (most recent form)

  var prefilledUrl = configSheet.getRange(lastRow, 2).getValue();

 

  Logger.log('Prefilled URL retrieved: ' + prefilledUrl);

 

  // Show success message

  ui.alert(

    'Form Sheet Updated Successfully!\n\n' +

    'Sheet renamed to: Data\n' +

    'Column order: Unique Key, Step Code, Timestamp\n\n' +

    'Prefilled URL retrieved from form_config.'

  );

 

  return prefilledUrl;

}



function createFormWithPrefilledURL() {

  // Get the active spreadsheet

  var ss = SpreadsheetApp.getActiveSpreadsheet();

 

  // Ask user for form name

  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(

    'Create New Form',

    'Please enter the name for the new form:',

    ui.ButtonSet.OK_CANCEL

  );

 

  // Check if user clicked OK

  if (response.getSelectedButton() != ui.Button.OK) {

    ui.alert('Form creation cancelled.');

    return;

  }

 

  var formName = response.getResponseText().trim();

 

  // Validate form name

  if (!formName || formName === '') {

    ui.alert('Error: Form name cannot be empty. Please try again.');

    return;

  }

 

  // Create a new form with the user-provided name

  var form = FormApp.create(formName);

 

  // Add questions to the form

  var uniqueKeyItem = form.addTextItem();

  uniqueKeyItem.setTitle('Unique Key')

               .setRequired(true);

 

  var stepCodeItem = form.addTextItem();

  stepCodeItem.setTitle('Step Code')

              .setRequired(true);

 

  // Link the form to the spreadsheet

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

 

  // Publish the form

  form.setAcceptingResponses(true);

 

  // Create prefilled URL using the more reliable method

  var prefilledResponse = form.createResponse()

    .withItemResponse(uniqueKeyItem.createResponse('u123u'))

    .withItemResponse(stepCodeItem.createResponse('k123k'))

    .toPrefilledUrl();

 

  // Get the regular form URL

  var formUrl = form.getPublishedUrl();

 

  // Create or get the hidden config sheet

  var configSheet = ss.getSheetByName('form_config');

 

  if (!configSheet) {

    configSheet = ss.insertSheet('form_config');

    // Add headers

    configSheet.getRange(1, 1, 1, 2).setValues([['Form Name', 'Prefilled URL']]);

    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold');

  }

 

  // Hide the config sheet

  configSheet.hideSheet();

 

  // Add the form configuration data

  var lastRow = configSheet.getLastRow();

  configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[

    formName,

    prefilledResponse

  ]]);

 

  // Auto-resize columns for better readability

  configSheet.autoResizeColumns(1, 2);

 

  // Log the results

  Logger.log('Form created: ' + formName);

  Logger.log('Form URL: ' + formUrl);

  Logger.log('Prefilled URL: ' + prefilledResponse);

 

  // Show a message to the user

  ui.alert(

    'Form Created Successfully!\n\n' +

    'Form Name: ' + formName + '\n\n' +

    'Configuration saved in hidden sheet "form_config"\n\n' +

    'Next Step: Run "3️⃣ Convert FMS to DB Format" to complete the setup.'

  );

 

  return {

    formUrl: formUrl,

    prefilledUrl: prefilledResponse,

    formName: formName

  };

}





// Combined initial setup function

function initialSetup() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var openingTime = ui.prompt('Enter Opening Time (e.g., 9 for 9 AM)', ui.ButtonSet.OK_CANCEL);

  if (openingTime.getSelectedButton() !== ui.Button.OK) return;

 

  var closingTime = ui.prompt('Enter Closing Time (e.g., 18 for 6 PM)', ui.ButtonSet.OK_CANCEL);

  if (closingTime.getSelectedButton() !== ui.Button.OK) return;

 

  // Enable iterative calculation

  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);

  ss.setIterativeCalculationEnabled(true);

  ss.setMaxIterativeCalculationCycles(1);

  ss.setIterativeCalculationConvergenceThreshold(0.05);

 

  // Setup hidden configuration row

  var sheet = ss.getActiveSheet();

  sheet.getRange('A1').setFormula('=now()');

  sheet.getRange('C1').setValue(parseFloat(openingTime.getResponseText()));

  sheet.getRange('D1').setValue(parseFloat(closingTime.getResponseText()));

  sheet.getRange('B1').setFormula('=D1/24-C1/24');

  sheet.hideRows(1);

 

  ui.alert('✅ Setup Complete!', 'Your spreadsheet is now configured with working hours.', ui.ButtonSet.OK);

}


// Improved FMS to DB conversion with Doer Emails sheet

function convertFMStoDBWithDoers() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ui = SpreadsheetApp.getUi();

 

  // First, ensure setup sheets exist

  createSetup();

 

  var sheet = ss.getSheetByName("Steps Directory");

  var lastRow = sheet.getLastRow() - 1;

 

  if (lastRow < 1) {

    ui.alert('⚠️ No Data Found', 'Please add steps to the "Steps Directory" sheet first.', ui.ButtonSet.OK);

    return;

  }

 

  var data = sheet.getRange(2, 1, lastRow, 6).getValues();

  var arr = [];

  var stepNames = [];

 

  // Build queries and collect step names

  for (var i = 0; i < data.length; i++) {

    var row = data[i];

    var step = row[0];

    var how = row[1];

    var keyrange = row[4];

    var range = row[3];

   

    stepNames.push([step]); // Collect step names for Doer Emails sheet

   

    var toChange = sheet.getRange(i + 2, 6);

    toChange.setValue('select Col1,Col2,\'' + step + '\',\'' + how + '\' where Col3 is null and Col1 is not null and Col2 is not null label \'' + step + '\' \'\',\'' + how + '\' \'\'');

   

    var formula = "ifna(query({" + keyrange + "," + range + "},'Steps Directory'!" + toChange.getA1Notation() + "),{\"\",\"\",\"\",\"\"})";

    if (i < data.length - 1) {

      formula += ";";

    }

    arr.push(formula);

  }

 

  // Build and set the main query formula

  var stringToSet = "=sort({" + arr.join("") + "},1,1)";

  ss.getSheetByName("DB_Format").getRange("A2").setFormula(stringToSet);

 

  // Create Doer Emails sheet

  createDoerEmailsSheet(stepNames);

 

  // Add Doer Email and Doer Name columns with VLOOKUP

  addDoerColumnsToDBFormat();

 

  // Get prefilled form URL from findAndEditFormSheet

  var prefilledLink = findAndEditFormSheet();

 

  if (!prefilledLink) {

    ui.alert('⚠️ Error', 'Could not retrieve prefilled URL. Please ensure you have run "2️⃣ Create Form with Prefilled URL" first.', ui.ButtonSet.OK);

    return;

  }

 

  var finalLink = prefilledLink.replace("u123u", "\"&$A$2:$A&\"").replace("k123k", "\"&vlookup($C$2:$C,'Steps Directory'!$A2:$C,3,false)))");

  var arrayFormula = "=ArrayFormula(if(isblank($A$2:$A),\"\",\"" + finalLink;

  var pcFormula = "=ArrayFormula(if(isblank($B$2:$B),\"\",if($B$2:$B<$G$1,\"Yes\",\"No\")))";

 

  ss.getSheetByName("DB_Format").getRange("E2").setFormula(arrayFormula);

  ss.getSheetByName("DB_Format").getRange("F2").setFormula(pcFormula);

 

  ui.alert('✅ Conversion Complete!', 'DB_Format sheet has been updated with:\n• Form links\n• Doer Email column\n• Doer Name column\n\nYour FMS to DB conversion is complete!', ui.ButtonSet.OK);

}


// Create Doer Emails sheet

function createDoerEmailsSheet(stepNames) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var doerSheet = ss.getSheetByName("Doer Emails");

 

  if (doerSheet != null) {

    doerSheet.clear();

  } else {

    doerSheet = ss.insertSheet();

    doerSheet.setName("Doer Emails");

  }

 

  // Add headers

  doerSheet.appendRow(["Step Name", "Doer Email", "Doer Name"]);

  doerSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");

 

  // Add step names

  if (stepNames.length > 0) {

    doerSheet.getRange(2, 1, stepNames.length, 1).setValues(stepNames);

  }

 

  // Format columns

  doerSheet.setColumnWidth(1, 200);

  doerSheet.setColumnWidth(2, 250);

  doerSheet.setColumnWidth(3, 200);

  doerSheet.setFrozenRows(1);

}


// Add Doer Email and Doer Name columns to DB_Format

function addDoerColumnsToDBFormat() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dbSheet = ss.getSheetByName("DB_Format");

 

  // Update headers

  var headers = dbSheet.getRange("A1:F1").getValues()[0];

  dbSheet.getRange("G1").setValue("Doer Email");

  dbSheet.getRange("H1").setValue("Doer Name");

 

  // Add VLOOKUP formulas

  var doerEmailFormula = "=ArrayFormula(if(isblank($C$2:$C),\"\",iferror(vlookup($C$2:$C,'Doer Emails'!$A:$B,2,false),\"\")))";

  var doerNameFormula = "=ArrayFormula(if(isblank($C$2:$C),\"\",iferror(vlookup($C$2:$C,'Doer Emails'!$A:$C,3,false),\"\")))";

 

  dbSheet.getRange("G2").setFormula(doerEmailFormula);

  dbSheet.getRange("H2").setFormula(doerNameFormula);

 

  // Format headers

  dbSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");

}


function createSetup() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

 

  // Create DB_Format sheet

  var datastudiosheet = ss.getSheetByName("DB_Format");

  if (datastudiosheet == null) {

    datastudiosheet = ss.insertSheet();

    datastudiosheet.setName("DB_Format");

    datastudiosheet.appendRow(["Unique Key", "Planned", "Step", "How", "Link", "For PC"]);

    datastudiosheet.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');

    datastudiosheet.getRange('E:E').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    datastudiosheet.getRange('G1').setFormula("=now()");

    datastudiosheet.getRange("A1:F1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");

  }

 

  // Create Steps Directory sheet

  var stepsDirectory = ss.getSheetByName("Steps Directory");

  if (stepsDirectory == null) {

    stepsDirectory = ss.insertSheet();

    stepsDirectory.setName("Steps Directory");

    stepsDirectory.appendRow(["Step", "How", "Step Code", "Planned/Actual Range", "Unique Key Range", "Query"]);

    stepsDirectory.getRange("A1:F1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");

    stepsDirectory.setFrozenRows(1);

  }

}


// Simplified prompt helper

function getPromptValue(ui, message) {

  var response = ui.prompt(message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  return response.getResponseText();

}


function importRangeFormula() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var donorSheet = getPromptValue(ui, 'Enter URL of the source spreadsheet:');

  if (!donorSheet) return;

 

  var tabName = getPromptValue(ui, 'Enter the sheet/tab name:');

  if (!tabName) return;

 

  var rangeName = getPromptValue(ui, 'Enter the range (e.g., A1:Z100):');

  if (!rangeName) return;

 

  ss.getCurrentCell().setFormula('=importrange("' + donorSheet + '","' + tabName + '!' + rangeName + '")');

}


function vlookupFormula() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var donorColumn = getPromptValue(ui, 'Enter Dependent Cell Column (e.g., A):');

  if (!donorColumn) return;

  donorColumn = donorColumn.replace(/\$/g, '');

 

  var tabColumn = getPromptValue(ui, 'Enter Unique Key Cell Column (e.g., B):');

  if (!tabColumn) return;

  tabColumn = tabColumn.replace(/\$/g, '');

 

  var rangeName = getPromptValue(ui, 'Enter Step Code Cell (e.g., $G$1):');

  if (!rangeName) return;

 

  var sheet = ss.getActiveSheet();

  var cell = sheet.getActiveCell();

  var column = cell.getColumn();

  var startRow = cell.getRow();

  var lastRow = sheet.getLastRow();

 

  var columnValues = sheet.getRange(startRow, column, lastRow - startRow + 1).getValues();

  var formulas = [];

 

  for (var i = 0; i < columnValues.length; i++) {

    var row = startRow + i;

    if (!columnValues[i][0]) {

      formulas.push([

        '=IF(ISBLANK(' + donorColumn + row + '), "", IFNA(VLOOKUP(' + tabColumn + row + '&' + rangeName + ', INDEX({Data!$A:$A&Data!$B:$B, Data!$C:$C}), 2, FALSE), ""))'

      ]);

    } else {

      formulas.push([columnValues[i][0]]);

    }

  }

 

  sheet.getRange(startRow, column, formulas.length, 1).setFormulas(formulas);

}


// Helper function to apply formula to range

function applyFormulaToRange(formula) {

  var ss = SpreadsheetApp.getActive();

  ss.getCurrentCell().setFormula(formula);

  var currentCell = ss.getCurrentCell();

  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

  currentCell.activateAsCurrentCell();

  currentCell = ss.getCurrentCell();

  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

  currentCell.activateAsCurrentCell();

  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

}


function plannedwwh() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

  var ofs = ss.getRange("C1").getValue();

  var ofe = ss.getRange("D1").getValue();

 

  var fromDate = getPromptValue(ui, 'Enter Date Cell (e.g., A2):');

  if (!fromDate) return;

 

  var tatInHours = getPromptValue(ui, 'Enter TAT Cell (e.g., G$5):');

  if (!tatInHours) return;

 

  var formula = '=if(' + fromDate + ',if(and(hour(' + fromDate + '+' + tatInHours + ')>' + ofs + ',(hour(' + fromDate + '+' + tatInHours + ')<' + ofe + ')),' + fromDate + '+' + tatInHours + ',workday.intl(int(' + fromDate + '),1,"0000001")+hour(' + fromDate + '+' + tatInHours + '-$B$1)/24+minute(' + fromDate + ')/1440),"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function plannedindays() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Date Cell (e.g., A2):');

  if (!fromDate) return;

 

  var tatInHours = getPromptValue(ui, 'Enter TAT Cell (e.g., G$5):');

  if (!tatInHours) return;

 

  var formula = '=if(' + fromDate + ',WORKDAY.INTL(' + fromDate + ',' + tatInHours + ',"0000001")+hour(' + fromDate + ')/24+MINUTE(' + fromDate + ')/1440,"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function plannedlead() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Date Cell:');

  if (!fromDate) return;

 

  var leadtime = getPromptValue(ui, 'Enter Lead Time Cell:');

  if (!leadtime) return;

 

  var tatInHours = getPromptValue(ui, 'Enter Number of Days Before Lead Time:');

  if (!tatInHours) return;

 

  var formula = '=if(' + leadtime + ',' + fromDate + '+' + leadtime + '-' + tatInHours + ',"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function specificTime() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Date Cell:');

  if (!fromDate) return;

 

  var leadtime = getPromptValue(ui, 'Enter Number of Days After Previous Planned (0 for same day):');

  if (!leadtime) return;

 

  var tatInHours = getPromptValue(ui, 'Enter Time of Day (hour/24 format):');

  if (!tatInHours) return;

 

  var formula = '=if(' + fromDate + ',workday.intl(int(' + fromDate + '),' + leadtime + ',"0000001",)+' + tatInHours + ',"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function actualTime() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var leadtime = getPromptValue(ui, 'Enter Status Cell:');

  if (!leadtime) return;

 

  var currcella1 = ss.getCurrentCell().getA1Notation();

  var formula = '=if(' + currcella1 + ',' + currcella1 + ',if(' + leadtime + '<>"",$A$1,""))';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function timeDelay() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Planned Cell:');

  if (!fromDate) return;

 

  var leadtime = getPromptValue(ui, 'Enter Actual Cell:');

  if (!leadtime) return;

 

  var formula = '=if(' + fromDate + ',if(' + leadtime + '<>"",if(' + leadtime + '>' + fromDate + ',' + leadtime + '-' + fromDate + ',""),$A$1-' + fromDate + '),"")';

 

  ss.getCurrentCell().setFormula(formula);

  ss.getActiveRangeList().setNumberFormat('[h]:mm:ss');

 

  var spreadsheet = SpreadsheetApp.getActive();

  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();

 

  // Add conditional formatting for delays

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()

    .setRanges([spreadsheet.getActiveRange()])

    .whenFormulaSatisfied('=if(' + leadtime + ',if(' + leadtime + '>' + fromDate + ',1,0),0)')

    .setBackground('#F4C7C3')

    .build());

 

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()

    .setRanges([spreadsheet.getActiveRange()])

    .whenFormulaSatisfied('=if(' + leadtime + ',0,if(' + fromDate + '<$A$1,1,0))')

    .setBackground('#FCE8B2')

    .build());

 

  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

 

  var currentCell = ss.getCurrentCell();

  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

  currentCell.activateAsCurrentCell();

  currentCell = ss.getCurrentCell();

  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

  currentCell.activateAsCurrentCell();

  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

}


function tatifno() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Status Cell:');

  if (!fromDate) return;

 

  var formulaincell = ss.getCurrentCell().getFormula().substr(1);

  var formula = '=if(' + fromDate + '="No",' + formulaincell + ',"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}


function tatifyes() {

  var ss = SpreadsheetApp.getActive();

  var ui = SpreadsheetApp.getUi();

 

  var fromDate = getPromptValue(ui, 'Enter Status Cell:');

  if (!fromDate) return;

 

  var formulaincell = ss.getCurrentCell().getFormula().substr(1);

  var formula = '=if(' + fromDate + '="Yes",' + formulaincell + ',"")';

 

  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  applyFormulaToRange(formula);

}




