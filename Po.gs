function getGSTNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Client Master"); // Change to your sheet name
  const values = sheet.getRange("H2:H").getValues().flat().filter(String); // Read GST numbers
  return values;
}

function getClientByGST(gstNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Client Master");
  const data = sheet.getRange(2, 2, sheet.getLastRow()-1, 8).getValues(); // A2:I
 // sheet.getRange(2,1,sheet.getLastRow()-1,8)
  const headers = ["buyerName","buyerMobile","buyerEmail","buyerAddress","city","state","buyerGST","buyerPAN"];
  for (let row of data) {
    if (row[6] == gstNumber) {  // GST column = H (index 7)
      let result = {};
      headers.forEach((key, i) => result[key] = row[i]);
      return result;
    }
  }
  return null;
}

function getBuyerDraftList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Order Received Details");
  const data = sheet.getDataRange().getValues();

  const quotationIndex = 29; 
  const buyerIndex = 6;    

  // Create "QuotationNo - BuyerName" format, skip blank rows
  const list = data.slice(3) // skip header
    .map(r => r[quotationIndex] && r[buyerIndex] ? `${r[quotationIndex]} - ${r[buyerIndex]}` : null)
    .filter(Boolean);
Logger.log(list);
  // Remove duplicates and sort
  return [...new Set(list)].sort();
  Logger.log("list"+[...new Set(list)].sort());
}
function getQuotationByNumberAndCompany(quotationNo, companyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Order Received Details");
  const data = sheet.getDataRange().getValues();

  // Column C = Quotation No (index 2), Column E = Company Name (index 4)
  const matchingRows = data.slice(3).filter(r => r[29] == quotationNo && r[6] == companyName);
  if (matchingRows.length === 0) return null;

  const firstRow = matchingRows[0];

  const items = matchingRows.map(r => ({
    itemName: r[14],   // N (index 13)
    itemCode: r[15],   // O (index 14)
    uom: r[16],        // P (index 15)
    qty: r[17],        // Q (index 16)
    rate: r[18]        // R (index 17)
  }));

  return {
    quotationNo: firstRow[29],
    companyName: firstRow[6],
    address: firstRow[7],
    state: firstRow[8],
    city: firstRow[9],
    contact: firstRow[10],
    email: firstRow[11],
    items: items
  };
}

function buildPendingOrdersReport() {
  const ss = SpreadsheetApp.getActive();
  const poDataSheet = ss.getSheetByName("PO_Data");
  const summarySheet = ss.getSheetByName("PO Summary");
  const pendingSheet = ss.getSheetByName("Pending Orders");
  
  const fromDate = new Date(pendingSheet.getRange("B1").getValue());
  const toDate   = new Date(pendingSheet.getRange("B2").getValue());

  const lastRowPO = poDataSheet.getLastRow();
  const lastRowSummary = summarySheet.getLastRow();

  // --- Build Status map from PO Summary (C:M) ---
  const summaryData = summarySheet.getRange(2, 3, Math.max(0, lastRowSummary - 1), 11).getValues(); // C:M
  const statusMap = {};
  summaryData.forEach(row => {
    const poNumber = row[0];  // Col C
    const status   = row[10]; // Col M
    if (poNumber != null && poNumber !== "") {
      statusMap[String(poNumber).trim()] = status ? String(status).trim().toLowerCase() : "";
    }
  });

  // --- Get PO_Data ---
  const poData = poDataSheet.getRange(2, 1, Math.max(0, lastRowPO - 1), 38).getValues(); 

  const results = [];
  let lastBuyer = "";

  poData.forEach((row, idx) => {
    const poDate        = row[2];    // Col C (PO Date)
    const poNumber      = row[1];    // Col B
    const buyerName     = row[4];    // Col E
    const salesExec     = row[37];   // Col 38 = Sales Executive (row[37])
    const itemName      = row[5];    // Col F
    const qty           = row[6];    // Col G
    const rate          = row[7];    // Col H
    const totalItemAmtRaw = row[31]; // Col 32 = Total Item Amount
    const finalAmount_fromsheet = row[26];

    if (!(poDate instanceof Date)) return; // skip if PO Date not valid

    const status = statusMap[String(poNumber).trim()] || "";

    if (status === "pending" && poDate >= fromDate && poDate <= toDate) {
      // Safely convert Total Item Amount to number
      const tAmount = Number(totalItemAmtRaw);
      const totalItemAmt = Number.isFinite(tAmount) ? tAmount : 0;

      // Calculate GST (18%) and Final Amount
      const gstValue = Math.round(totalItemAmt * 0.18 * 100) / 100;
      const finalAmount = Math.round((totalItemAmt + gstValue) * 100) / 100;

      // Add actual data row
      results.push([
        Utilities.formatDate(poDate, Session.getScriptTimeZone(), "dd-MM-yyyy"),
        `${buyerName} - (${salesExec || ""}) - (${finalAmount_fromsheet.toFixed(2)})`,
        itemName || "",
        qty || "",
        rate || "",
        totalItemAmt,
        gstValue,
        finalAmount
      ]);

      // If buyer changes in the NEXT row → add a blank row now
      const nextBuyer = (poData[idx + 1] && poData[idx + 1][4]) || "";
      if (buyerName !== nextBuyer) {
        results.push(["", "", "", "", "", "", "", ""]);
      }
      lastBuyer = buyerName;
    }
  });

  // --- Clear old report rows ---
  const startRow = 5;
  const lastPendingRow = pendingSheet.getLastRow();
  if (lastPendingRow >= startRow) {
    pendingSheet.getRange(startRow, 1, lastPendingRow - startRow + 1, 8).clearContent();
  }

  // --- Paste new results ---
  if (results.length > 0) {
    pendingSheet.getRange(startRow, 1, results.length, results[0].length).setValues(results);

    // --- Merge Buyer Name (Column B) by groups ---
    let rowPointer = startRow;
    let groupStart = startRow;
    let prevBuyer = pendingSheet.getRange(startRow, 2).getValue();

    for (let i = startRow + 1; i < startRow + results.length; i++) {
      let currentBuyer = pendingSheet.getRange(i, 2).getValue();
      if (currentBuyer !== "" && currentBuyer !== prevBuyer) {
        // Merge previous buyer group
        if (i - groupStart > 1) {
          pendingSheet.getRange(groupStart, 2, i - groupStart, 1).mergeVertically()
            .setVerticalAlignment("middle");
        }
        groupStart = i;
        prevBuyer = currentBuyer;
      }
    }
    // Merge the last group
    let lastRow = startRow + results.length - 1;
    if (lastRow - groupStart >= 1) {
      pendingSheet.getRange(groupStart, 2, lastRow - groupStart + 1, 1).mergeVertically()
        .setVerticalAlignment("middle");
    }
  }
}

