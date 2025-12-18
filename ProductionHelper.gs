function generateCustomerFGDetail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const fgSheet = ss.getSheetByName("Customer FG");
  const openingSheet = ss.getSheetByName("Opening Stock Details");

  // Get date from B2
  const filterDate = fgSheet.getRange("B2").getValue();
  if (!filterDate) {
    SpreadsheetApp.getUi().alert("Please enter a date in cell B2 of Customer FG sheet.");
    return;
  }
  const filterDateStr = Utilities.formatDate(new Date(filterDate), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  let customerStock = {};

  // --- From Opening Stock ---
  const lastCol = openingSheet.getLastColumn();
  const openHeaders = openingSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const lastRow = openingSheet.getLastRow();
  const openData = openingSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const oIdxItemName = openHeaders.indexOf("Item Name");
  const oIdxItemCode = openHeaders.indexOf("Item Code");
  const oIdxDrum = openHeaders.indexOf("Drum / Bundle");
  const oIdxLength = openHeaders.indexOf("Length");
  const oIdxDept = openHeaders.indexOf("Department");
  const oIdxColor = openHeaders.indexOf("Code");
  const oIdxParty = openHeaders.indexOf("Party / Godown");
  const oIdxDate = openHeaders.indexOf("Production Date");

  openData.forEach(r => {
    if (!r[oIdxDate]) return;
    const rowDateStr = Utilities.formatDate(new Date(r[oIdxDate]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const partyName = String(r[oIdxParty]).trim().toUpperCase();
    if (rowDateStr <= filterDateStr && partyName && partyName !== "GODOWN") {
      const key = normalizeKey(r[oIdxItemName], r[oIdxItemCode], r[oIdxDept], r[oIdxColor], partyName);
      if (!customerStock[key]) customerStock[key] = { drums: {} };
      const drumKey = String(r[oIdxDrum] || "").trim();
      const length = Number(r[oIdxLength] || 0);
      customerStock[key].drums[drumKey] = (customerStock[key].drums[drumKey] || 0) + length;
    }
  });

  // --- From Production/Dispatch ---
  const lastColP = prodSheet.getLastColumn();
  const prodHeaders = prodSheet.getRange(2, 1, 1, lastColP).getValues()[0];
  const lastRowP = prodSheet.getLastRow();
  const prodData = prodSheet.getRange(3, 1, lastRowP - 2, lastColP).getValues();

  const pIdxTimestamp = prodHeaders.indexOf("Production Date");
  const pIdxItemName = prodHeaders.indexOf("Item Name");
  const pIdxItemCode = prodHeaders.indexOf("Item Code");
  const pIdxDrum = prodHeaders.indexOf("Drum");
  const pIdxLength = prodHeaders.indexOf("Party assign Length");
  const pIdxDept = prodHeaders.indexOf("Department");
  const pIdxColor = prodHeaders.indexOf("Code");
  const pIdxStatus = prodHeaders.indexOf("Status");
  const pIdxPartyName = prodHeaders.indexOf("Party Name");

  prodData.forEach(r => {
    const rowDateStr = Utilities.formatDate(new Date(r[pIdxTimestamp]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    if (rowDateStr > filterDateStr) return;

    const partyName = String(r[pIdxPartyName]).trim().toUpperCase();
    if (!partyName || partyName === "GODOWN") return;

    const key = normalizeKey(r[pIdxItemName], r[pIdxItemCode], r[pIdxDept], r[pIdxColor], partyName);
    if (!customerStock[key]) customerStock[key] = { drums: {} };

    const drumKey = String(r[pIdxDrum] || "").trim();
    const length = Number(r[pIdxLength]) || 0;
    const status = String(r[pIdxStatus] || "").trim().toUpperCase();

    if (status === "IN") {
      customerStock[key].drums[drumKey] = (customerStock[key].drums[drumKey] || 0) + length;
    } else if (status === "OUT" || status === "SCRAP") {
      const existing = customerStock[key].drums[drumKey] || 0;
      const newLength = existing - length;
      if (newLength > 0) {
        customerStock[key].drums[drumKey] = newLength;
      } else {
        delete customerStock[key].drums[drumKey];
      }
    }
  });

  // === Clear old FG content ===
  fgSheet.getRange("A4:Z1000").clearContent().clearFormat();

  // --- Party-wise grouping ---
  let stockArray = Object.entries(customerStock).filter(([_, grp]) =>
    Object.values(grp.drums).some(len => len > 0)
  );

  stockArray.sort((a, b) => {
    const partyA = a[0].split("|")[4] || "";
    const partyB = b[0].split("|")[4] || "";
    if (partyA < partyB) return -1;
    if (partyA > partyB) return 1;
    const itemNameA = a[0].split("|")[0];
    const itemNameB = b[0].split("|")[0];
    return itemNameA.localeCompare(itemNameB, undefined, { numeric: true });
  });

  let grouped = {};
  stockArray.forEach(([key, group]) => {
    const [itemName, itemCode, dept, color, partyName] = key.split("|");
    const itemKey = `${itemName}-${itemCode}(${color})(${dept})`.replace(/\bCORE\b/gi, "C")
      .replace(/\s*X\s*/gi, "X")
      .replace(/\s*SQ\s*MM\s*/gi, "")
      .replace(/\s+/g, "");
    if (!grouped[partyName]) grouped[partyName] = [];
    grouped[partyName].push({
      itemKey,
      drums: Object.entries(group.drums).filter(([_, len]) => len > 0)
    });
  });

  // --- Write output dynamically with column tracking ---
  const startRow = 4;
  const startCol = 1;
  const maxCol = 26; // column Z
  const colPerParty = 3; // 2 for drum+length + 1 spacing
  const colWidth = 80;
  let colUsage = {}; // last used row per starting column
  let currentCol = startCol;

  for (const partyName of Object.keys(grouped)) {
    if (currentCol + colPerParty - 1 > maxCol) {
      currentCol = startCol;
    }

    const currentRow = colUsage[currentCol] ? colUsage[currentCol] + 1 : startRow;

    // Set column widths
    for (let c = currentCol; c < currentCol + 2; c++) fgSheet.setColumnWidth(c, colWidth);

    // Party Header
const displayPartyName = partyName.split(" ")[0]; // Take only first word
const partyRange = fgSheet.getRange(currentRow, currentCol, 1, 2).merge();
partyRange.setValue(displayPartyName)
  .setFontSize(14)
  .setHorizontalAlignment("center")
  .setVerticalAlignment("middle")
  .setBackground("#F2ECEB")
  .setWrap(true)
  .setBorder(true, true, true, true, true, true);

    let localRow = currentRow + 1;

    grouped[partyName].forEach(item => {
      // Item Name
      const itemRange = fgSheet.getRange(localRow, currentCol, 1, 2).merge();
      itemRange.setValue(item.itemKey)
        .setBackground("#CCC4C4")
        .setBorder(true, true, true, true, true, true);
      fgSheet.autoResizeRows(localRow, 1);
      localRow++;

      // Drum Header
      const drumHeader = fgSheet.getRange(localRow, currentCol, 1, 2);
      drumHeader.setValues([["Drum", "Length"]])
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true);
      fgSheet.autoResizeRows(localRow, 1);

      let total = 0;
      item.drums.forEach(([drum, len]) => {
        const drumRange = fgSheet.getRange(localRow, currentCol, 1, 2);
        drumRange.setValues([[drum, len]])
          .setHorizontalAlignment("center")
          .setBorder(true, true, true, true, true, true);
        fgSheet.autoResizeRows(localRow, 1);
        total += len;
        localRow++;
      });

      // Total row
      const totalRange = fgSheet.getRange(localRow, currentCol, 1, 2);
      totalRange.setValues([["Total", total]])
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setBackground("#F2ECEB")
        .setBorder(true, true, true, true, true, true);
      localRow++;
    });

    colUsage[currentCol] = localRow - 1;
    currentCol += colPerParty;
  }
}

// Utility: normalize item+code+dept+color+party
function normalizeKey(itemName, itemCode, dept, color, party) {
  return [
    String(itemName).trim().toUpperCase(),
    String(itemCode).trim().toUpperCase(),
    String(dept).trim().toUpperCase(),
    String(color).trim().toUpperCase(),
    String(party).trim().toUpperCase()
  ].join("|");
}




function getProductionSummaryIN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const summarySheet = ss.getSheetByName("Stock Summary");

  // Get from & to dates
  const fromDate = new Date(summarySheet.getRange("J2").getValue());
  const toDate = new Date(summarySheet.getRange("L2").getValue());
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);

  const lastRow = prodSheet.getLastRow();
  if (lastRow <= 1) return;

  const lastCol = prodSheet.getLastColumn();
  const headers = prodSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const data = prodSheet.getRange(3, 1, lastRow - 1, lastCol).getValues();

  const idxItemName = headers.indexOf("Item Name");
  const idxItemCode = headers.indexOf("Item Code");
  const idxCategory = headers.indexOf("Category");
  const idxLength = headers.indexOf("Production Length");
  const idxStatus = headers.indexOf("Status");
  const idxDate = headers.indexOf("Production Date");

  let stockMap = {}; // key = ItemName|ItemCode|Category, value = total IN

  data.forEach(row => {
    const prodDate = new Date(row[idxDate]);
    prodDate.setHours(0, 0, 0, 0);
    if (prodDate < fromDate || prodDate > toDate) return;

    const status = String(row[idxStatus]).trim().toUpperCase();
    if (status !== "IN") return; // only IN

    const length = Number(row[idxLength]) || 0;
    const key = [row[idxItemName], row[idxItemCode], row[idxCategory]].join("|");

    if (!stockMap[key]) stockMap[key] = 0;
    stockMap[key] += length;
  });

  // Convert to array
  const output = Object.entries(stockMap)
    .map(([key, qty]) => {
      const [itemName, itemCode, category] = key.split("|");
      return [itemName, itemCode, category, qty];
    });
    const grandTotal = output.reduce((sum, row) => sum + Number(row[3]), 0);
    if (output.length > 0) {
    output.push(["Grand Total", "", "", grandTotal]); // Add total row
    }
  // Clear old data
  summarySheet.getRange("I4:L").clearContent();

  if (output.length > 0) {
    summarySheet.getRange(4, 9, output.length, 4).setValues(output);
  }
   const totalRowIndex = 4 + output.length - 1; // last row
    summarySheet.getRange(totalRowIndex, 9, 1, 4)
      .setFontWeight("bold")
      .setFontSize(16);    
}


function getAvailableStockWithDeptColor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const openingSheet = ss.getSheetByName("Opening Stock Details");
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const outputSheet = ss.getSheetByName("Stock By Drum") || ss.insertSheet("Stock By Drum");

  // Clear previous content
  outputSheet.clearContents().clearFormats();

  // --- Get data ---
  const openingData = openingSheet.getRange(3,1,openingSheet.getLastRow()-2,10).getValues();
  const prodData = prodSheet.getRange(3,1,prodSheet.getLastRow()-2,12).getValues();

  // key = itemName||itemCode||drum||party
  // value = {stock, dept, color, category}
  const stockMap = {};

  // --- Opening Stock ---
  openingData.forEach(row => {
    const itemName = String(row[0]).trim();
    const itemCode = String(row[1]).trim();
    const category = String(row[2] || "").trim().toUpperCase();
    const drum = String(row[3]).trim();
    const party = String(row[5]).trim();
    const dept = String(row[7]).trim(); // H column
    const color = String(row[8]).trim(); // I column

    if (!itemName || !itemCode || !drum || !party) return;

    const length = party === "Godown" ? Number(row[4]) || 0 : Number(row[6]) || 0;
    const key = itemName + "||" + itemCode + "||" + drum + "||" + party;

    if (!stockMap[key]) stockMap[key] = {stock: 0, dept: dept, color: color, category: category};
    stockMap[key].stock += length;
  });

  // --- Production/Dispatch ---
  prodData.forEach(row => {
    const itemName = String(row[1] || "").trim();
    const itemCode = String(row[2] || "").trim();
    const category = String(row[3] || "").trim().toUpperCase();
    const drum = String(row[4] || "").trim();
    const party = String(row[6] || "").trim();
    const dept = String(row[8] || "").trim(); // J column
    const color = String(row[9] || "").trim(); // K column
    const status = String(row[10] || "").trim().toUpperCase();

    if (!itemName || !itemCode || !drum || !party) return;

    const length = party === "Godown" ? Number(row[5]) || 0 : Number(row[7]) || 0;
    const key = itemName + "||" + itemCode + "||" + drum + "||" + party;

    if (!stockMap[key]) stockMap[key] = {stock: 0, dept: dept, color: color, category: category};

    if (status === "IN") stockMap[key].stock += length;
    else if (status === "OUT" || status === "SCRAP") stockMap[key].stock -= length;
  });

  // --- Write Output (skip 0 length drums) ---
  const output = [["Item Name","Item Code","Drum","Party/Godown","Department","Color Code","Available Stock","Category"]];
  for (const k in stockMap) {
    const item = stockMap[k];
    const available = Math.max(0, item.stock);
    if (available > 0) { // Only positive stock
      const [itemName,itemCode,drum,party] = k.split("||");
      output.push([itemName,itemCode,drum,party,item.dept,item.color,available,item.category]);
    }
  }

  outputSheet.getRange(1,1,output.length,output[0].length).setValues(output);
 // SpreadsheetApp.getUi().alert("Available stock per drum/party calculated (Department, Color Code & Category included)!");
}

