function getUniquePONumbersWithBuyer() {
 const ss = SpreadsheetApp.getActive();
  const poDataSheet = ss.getSheetByName("PO_Data");
  const summarySheet = ss.getSheetByName("PO Summary");

  const lastRowPO = poDataSheet.getLastRow();
  const lastRowSummary = summarySheet.getLastRow();

  // --- Build PO → Status map from "PO Summary" ---
  const summaryData = summarySheet.getRange(2, 3, lastRowSummary - 1, 11).getValues(); // C:M (11 cols)
  const statusMap = {};
  summaryData.forEach(row => {
    const poNumber = row[0];   // Col C
    const status   = row[9];  // Col M (11th col from C)
    if (poNumber) {
      statusMap[poNumber.toString().trim()] = status ? status.toString().trim().toLowerCase() : "";
    }
  });

  // --- Read data from PO_Data ---
  const poData = poDataSheet.getRange(2, 2, lastRowPO - 1, 38).getValues(); 
  // B..AL (38 cols)

  const seen = new Set();
  const results = [];

  poData.forEach(row => {
    const poNumber  = row[0];   // Col B
    const buyerName = row[3];   // Col E
    const unit      = row[7] ? row[7].toString().trim().toLowerCase() : "";
    const status    = statusMap[poNumber] || "";  // take from PO Summary

    if (poNumber && buyerName && unit === "unit 1" && status === "pending") {
      const key = `${poNumber}-${buyerName}`;
      if (!seen.has(key)) {
        results.push({ po: poNumber, label: `${poNumber} - ${buyerName}` });
        seen.add(key);
      }
    }
  });

  Logger.log(results);
  return results;
}

function getUniquePONumbersWithBuyerUnit2() {
  const ss = SpreadsheetApp.getActive();
  const poDataSheet = ss.getSheetByName("PO_Data");
  const summarySheet = ss.getSheetByName("PO Summary");

  const lastRowPO = poDataSheet.getLastRow();
  const lastRowSummary = summarySheet.getLastRow();

  // --- Build PO → Status map from "PO Summary" ---
  const summaryData = summarySheet.getRange(2, 3, lastRowSummary - 1, 11).getValues(); // C:M (11 cols)
  const statusMap = {};
  summaryData.forEach(row => {
    const poNumber = row[0];   // Col C
    const status   = row[9];  // Col M (11th col from C)
    if (poNumber) {
      statusMap[poNumber.toString().trim()] = status ? status.toString().trim().toLowerCase() : "";
    }
  });

  // --- Read data from PO_Data ---
  const poData = poDataSheet.getRange(2, 2, lastRowPO - 1, 38).getValues(); 
  // B..AL (38 cols)

  const seen = new Set();
  const results = [];

  poData.forEach(row => {
    const poNumber  = row[0];   // Col B
    const buyerName = row[3];   // Col E
    const unit      = row[7] ? row[7].toString().trim().toLowerCase() : "";
    const status    = statusMap[poNumber] || "";  // take from PO Summary

    if (poNumber && buyerName && unit === "unit 2" && status === "pending") {
      const key = `${poNumber}-${buyerName}`;
      if (!seen.has(key)) {
        results.push({ po: poNumber, label: `${poNumber} - ${buyerName}` });
        seen.add(key);
      }
    }
  });

  Logger.log(results);
  return results;
}

function testpo()
{
Logger.log(getUniquePONumbersWithBuyer())

}

function getItemsByPONumber(poNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++){
    if (data[i][1] === poNumber) { // Column B = index 1
      const itemName = data[i][5]; // Column F = index
      if (itemName && !items.includes(itemName)) {
        items.push(itemName);
      }
    }
  }
  return items;
}
function getItemTypeByPOAndItem(poNumber, itemName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowPO = data[i][1];    // Column B = PO number
    const rowItem = data[i][5];  // Column F = Item
    const type = data[i][30];    // Column AE = Type
    if (rowPO === poNumber && rowItem === itemName) {
      return type;
    }
  }
  return ""; // Return blank if not found
}
Logger.log(getDrumsByItem("3.5 Core X 120 Sq mm","A2XFY", "Tesla Transformers (India) Ltd."))

function getDrumsByItem(itemName, itemCode, buyer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Production-Unit I");
  const data = sheet.getDataRange().getValues();

  let assignedDrums = [];
  let godownDrums = [];

  for (let i = 1; i < data.length; i++) { // skip header
    const rowItemName = data[i][0].toString().trim().toLowerCase();
    const rowItemCode = data[i][1].toString().trim().toLowerCase();
    const drum = data[i][2].toString().trim();
    const party = data[i][3].toString().trim().toLowerCase();

    if (rowItemName === itemName.trim().toLowerCase() && rowItemCode === itemCode.trim().toLowerCase()) {
      if (!drum) continue; // skip empty drums
      if (party === buyer.trim().toLowerCase()) {
        assignedDrums.push(drum);
      } else if (party === "godown") {
        godownDrums.push(drum);
      }
    }
  }
  if (assignedDrums.length > 0) {
    Logger.log(`Drums assigned to ${buyer}: ${assignedDrums.join(", ")}`);
    return assignedDrums;
  } else {
    Logger.log(`No drums assigned to ${buyer}. Returning Godown drums: ${godownDrums.join(", ")}`);
    return godownDrums;
  }
}

function getDrumsByItemUnitII(itemName, itemCode, buyer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Production-Unit II");
  const data = sheet.getDataRange().getValues();

  let assignedDrums = [];
  let godownDrums = [];

  for (let i = 1; i < data.length; i++) { // skip header
    const rowItemName = data[i][0].toString().trim().toLowerCase();
    const rowItemCode = data[i][1].toString().trim().toLowerCase();
    const drum = data[i][2].toString().trim();
    const party = data[i][3].toString().trim().toLowerCase();

    if (rowItemName === itemName.trim().toLowerCase() && rowItemCode === itemCode.trim().toLowerCase()) {
      if (!drum) continue; // skip empty drums
      if (party === buyer.trim().toLowerCase()) {
        assignedDrums.push(drum);
      } else if (party === "godown") {
        godownDrums.push(drum);
      }
    }
  }
  if (assignedDrums.length > 0) {
    Logger.log(`Drums assigned to ${buyer}: ${assignedDrums.join(", ")}`);
    return assignedDrums;
  } else {
    Logger.log(`No drums assigned to ${buyer}. Returning Godown drums: ${godownDrums.join(", ")}`);
    return godownDrums;
  }
}

function getQtyByDrum(drumNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Production-Unit I");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim().toLowerCase() === drumNo.trim().toLowerCase()) {
      return data[i][6]; // Available Qty (G column, index 6)
    }
  }
  return "";
}
function getQtyByDrumUnitII(drumNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Production-Unit II");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim().toLowerCase() === drumNo.trim().toLowerCase()) {
      return data[i][6]; // Available Qty (G column, index 6)
    }
  }
  return "";
}

function testDrums()
{
 let result= getDrumsByItem("7 Core X 2.5 Sq mm","YWY")
 Logger.log("Drum "+result);
}
Logger.log(getConsigneeDetails());
function getConsigneeDetails(poNumber) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("PO_Data");
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // First row
  const PO_INDEX = headers.indexOf("PO Number");
  const BUYER_NAME_INDEX = headers.indexOf("Buyer Name");
  const CONSIGNEE_NAME_INDEX = headers.indexOf("Consignee Name");
  const CONSIGNEE_ADDRESS_INDEX = headers.indexOf("Consignee Address");
  const CONSIGNEE_MOBILE_INDEX = headers.indexOf("Consignee Mobile");
  const BUYER_ADDRESS_INDEX = headers.indexOf("Buyer Address");
  const BUYER_MOBILE_INDEX = headers.indexOf("Buyer Mobile");
  const PO_DATE_INDEX = headers.findIndex(h => h.trim().toLowerCase() === "po date");

  for (let i = 1; i < data.length; i++) {
    if (data[i][PO_INDEX] == poNumber) {
      const poDateRaw = data[i][PO_DATE_INDEX];
      const poDateFormatted = poDateRaw
        ? Utilities.formatDate(new Date(poDateRaw), Session.getScriptTimeZone(), "dd-MM-yyyy")
        : "";
      return {
        buyerName: data[i][BUYER_NAME_INDEX] || "",
        consigneeName: data[i][CONSIGNEE_NAME_INDEX] || "",
        consigneeAddress: data[i][CONSIGNEE_ADDRESS_INDEX] || "",
        consigneeMobile: data[i][CONSIGNEE_MOBILE_INDEX] || "",
        buyerAddress: data[i][BUYER_ADDRESS_INDEX] || "",
        buyerMobile: data[i][BUYER_MOBILE_INDEX] || "",
        poDate: poDateFormatted
      };
    }
  }
  return {}; // return empty if not found
}
function savePackingDetails(formData) {
  const packSS = SpreadsheetApp.getActive().getSheetByName("Packing Details");
  const timestamp = new Date();
  const packingType = formData.packingType;

  // Prepare rows to insert/update
  const rows = formData.items.map(item => [
    timestamp,
    formData.poNumber,
    formData.buyerName,
    formData.material,
    formData.consignee,
    formData.consigneeAddress,
    formData.consigneeMobile1,
    formData.consigneeMobile2,
    item.itemName,
    item.type,
    item.drum,
    item.length,
    packingType,       // Column M: Packing Type
    formData.poDate,
    formData.inv1,
    formData.invdate1,
    formData.inv2,
    formData.invdate2,
    packingType,
    formData.unit
  ]);

  const data = packSS.getDataRange().getValues();
  const poIndex = 1; // Column B = PO Number
  const packingTypeIndex = 12; // Column M = Packing Type

  // -------- NEW → Preview --------
  if (formData.mode === "NEW" && packingType === "Preview") {
    packSS.getRange(packSS.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  // -------- DRAFT → Preview (update existing Draft) --------
  else if (formData.mode === "DRAFT" && packingType === "Preview") {
    const deleteRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][poIndex] === formData.poNumber && data[i][packingTypeIndex] === "Preview") {
        deleteRows.push(i + 1);
      }
    }
    for (let i = deleteRows.length - 1; i >= 0; i--) {
      packSS.deleteRow(deleteRows[i]);
    }
    packSS.getRange(packSS.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  // -------- FINAL --------
  else if (packingType === "Final") {
    const draftRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][poIndex] === formData.poNumber && data[i][packingTypeIndex] === "Preview") {
        draftRows.push(i + 1); // Draft rows to update
      }
    }

    if (draftRows.length > 0) {
      // DRAFT → Final: update existing Draft rows
      draftRows.forEach((rowNum, idx) => {
        const rowData = rows[idx] || rows[0]; // Map items
        packSS.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      });
    } else {
      // NEW → Final: append as new Final rows
      packSS.getRange(packSS.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  }

  // -------- Generate packing sheet URL --------
  let url = "";
  if (formData.unit === "I") {
    url = generatePackingSheet(formData);
  } else {
    url = generatePackingSheetUnit2(formData);
  }

  return url;
}

function getDraftPONumbersWithBuyerUnitI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Packing Details");
  const data = sheet.getDataRange().getValues();

  const poMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const status = data[i][12]; // Column M = index 12
    const poNumber = data[i][1]; // PO Number = column B = index 1
    const buyerName = data[i][2]; // Buyer Name = column C = index 2
    const unit = data[i][19];
    if (status === "Preview" && poNumber && unit==="I" ) {
      if (!poMap.has(poNumber)) {
        poMap.set(poNumber, buyerName);
      }
    }
  }

  const poList = [];
  poMap.forEach((buyer, po) => {
    poList.push({
      po: po,
      label: `${po} - ${buyer}`
    });
  });
  Logger.log(poList);
  return poList;
}

function getDraftPONumbersWithBuyerUnitII() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Packing Details");
  const data = sheet.getDataRange().getValues();

  const poMap = new Map();

  for (let i = 1; i < data.length; i++) {
    const status = data[i][12]; // Column M = index 12
    const poNumber = data[i][1]; // PO Number = column B = index 1
    const buyerName = data[i][2]; // Buyer Name = column C = index 2
    const unit = data[i][19];
    if (status === "Preview" && poNumber && unit==="II" ) {
      if (!poMap.has(poNumber)) {
        poMap.set(poNumber, buyerName);
      }
    }
  }

  const poList = [];
  poMap.forEach((buyer, po) => {
    poList.push({
      po: po,
      label: `${po} - ${buyer}`
    });
  });
  Logger.log(poList);
  return poList;
}

function testDraftData()
{
  Logger.log(getDraftPOData("TTIL/VISHAL/2025-26/400436/PO/018"));
}

function getDraftPOData(poNumber,unit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Packing Details");
  const data = sheet.getDataRange().getValues();

  const rows = data.slice(1).filter(r => r[1] === poNumber && r[19] === unit); // PO Number is in column B (index 1)

  if (rows.length === 0) return null;

  const headerRow = rows[0];

  // Format PO Date (column N → index 13)
  const rawPoDate = headerRow[13];
  const poDateFormatted = rawPoDate
    ? Utilities.formatDate(new Date(rawPoDate), Session.getScriptTimeZone(), "yyyy-MM-dd") // required for <input type="date">
    : "";
    const rawinvDate1=headerRow[15];
      const invDate1Formatted = rawinvDate1
    ? Utilities.formatDate(new Date(rawinvDate1), Session.getScriptTimeZone(), "yyyy-MM-dd") // required for <input type="date">
    : "";

     const rawinvDate2=headerRow[17];
      const invDate2Formatted = rawinvDate2
    ? Utilities.formatDate(new Date(rawinvDate2), Session.getScriptTimeZone(), "yyyy-MM-dd") // required for <input type="date">
    : "";

  // Map header fields
  const header = {
    "PO Number": headerRow[1],
    "Buyer Name": headerRow[2],
    "Material": headerRow[3],
    "Consignee Name": headerRow[4],
    "Consignee Address": headerRow[5],
    "Mobile 1": headerRow[6],
    "Mobile 2": headerRow[7],
    "Po Date": poDateFormatted,
    "Inv 1":headerRow[14],
    "Inv 2":headerRow[16],
    "invDate1": invDate1Formatted,
    "invDate2": invDate2Formatted
  };
  // Item-level fields
  const items = rows.map(r => ({
    "Items": r[8],
    "Item Type": r[9],
    "Drum No": r[10],
    "Length": r[11]
  }));

  return {
    header,
    items
  };
}

function test()
{
    let y=getDraftPOData("VMS/CABLE/PO/11")
    Logger.log(y);
}


function generatePackingSheet(data) {
  const timestamp = new Date();
  stockSummary = SpreadsheetApp.getActive().getSheetByName("Production-Unit I");
  const stockData = stockSummary.getRange(2, 1, stockSummary.getLastRow() - 1, 8).getValues();
  // Stock Summary Columns: [Item Name, Item Code, Category, Drum, Department, Color Code, Available Qty]

  const inOutSS = SpreadsheetApp
    .openById("1n_sIZucW6Ha2_WXkDuXpVU9hMoyVLPdI18Dc-NHW2Hc")
    .getSheetByName("In-Out");
  let inOutRows = [];
  data.items.forEach(item => {
    // find matching drum in stock summary
    const match = stockData.find(row =>
      row[0].toString().trim().toLowerCase() === item.itemName.toLowerCase().trim() &&
      row[1].toString().trim().toLowerCase() === item.type.toLowerCase().trim() &&
      row[2].toString().trim().toLowerCase() === item.drum.toLowerCase().trim()
    );
    const today = new Date();
const dd = String(today.getDate()).padStart(2, '0');
const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months are 0-based
const yyyy = today.getFullYear();
const formattedDate = dd + '-' + mm + '-' + yyyy; // e.g., 21-09-2025

    if (match) {
      const department = match[4] || "";
      const colorCode = match[5] || "";
      const category = match[7] || "";
      inOutRows.push([
        timestamp,
        item.itemName,
        item.type,
        category,
        item.drum,
        item.length,
        data.buyerName,
        item.length,
        department,
        colorCode,
        "OUT",
        formattedDate,
      ]);
    } else {
      // fallback if not found
      inOutRows.push([
        timestamp,
        item.itemName,
        item.type,
        category,
        item.drum,
        item.length,
        data.buyerName,
        item.length,
        "",
        "",
        "OUT",
        formattedDate,
      ]);
    }
  });
  if (inOutRows.length > 0) {
    console.log("Save in production");
    if(data.packingType==="Final")
    {
    inOutSS.getRange(inOutSS.getLastRow() + 1, 1, inOutRows.length, inOutRows[0].length)
      .setValues(inOutRows);
    }
  }

  const masterFileId = '1ocr3vP60u3dbExOWJv4aRUpHcmPcr4846-1M8gFd0P4';
  const templateFileId = '1TUdj7Tx-CHp8OIXAUqjdXa2Y3zWCJMxp7ZhTN400Ovc';
  const masterSS = SpreadsheetApp.openById(masterFileId);
  const templateSS = SpreadsheetApp.openById(templateFileId);
  const templateSheet = templateSS.getSheets()[0]; // assumes first sheet is the template

  if (!templateSheet) {
    throw new Error("Template sheet not found");
  }
  // Delete old packing sheet with same PO number
  const oldSheet = masterSS.getSheetByName(data.poNumber);
  if (oldSheet) masterSS.deleteSheet(oldSheet);

  // Copy template to master
  const sheet = templateSheet.copyTo(masterSS).setName(data.poNumber);

  // Replace placeholders
  replacePackingPlaceholders(sheet, data);

  // Insert item rows
  insertPackingItemsAtPlaceholder(sheet, "[PACKING_TABLE]", data.items);

  const sheetId = sheet.getSheetId();
  return masterSS.getUrl() + "#gid=" + sheetId;
}

function generatePackingSheetUnit2(data) {
  const timestamp = new Date();
  stockSummary = SpreadsheetApp.getActive().getSheetByName("Production-Unit II");
  const stockData = stockSummary.getRange(2, 1, stockSummary.getLastRow() - 1, 8).getValues();
  // Stock Summary Columns: [Item Name, Item Code, Category, Drum, Department, Color Code, Available Qty]

  const inOutSS = SpreadsheetApp
    .openById("1CqWc5fVKQKOhUJvz3zy9I7382Yi1zBGEqMwK4DM25O0")
    .getSheetByName("In-Out");
  let inOutRows = [];
  data.items.forEach(item => {
    // find matching drum in stock summary
    const match = stockData.find(row =>
      row[0].toString().trim().toLowerCase() === item.itemName.toLowerCase().trim() &&
      row[1].toString().trim().toLowerCase() === item.type.toLowerCase().trim() &&
      row[2].toString().trim().toLowerCase() === item.drum.toLowerCase().trim()
    );
    const today = new Date();
const dd = String(today.getDate()).padStart(2, '0');
const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months are 0-based
const yyyy = today.getFullYear();
const formattedDate = dd + '-' + mm + '-' + yyyy; // e.g., 21-09-2025

    if (match) {
      const department = match[4] || "";
      const colorCode = match[5] || "";
      const category = match[7] || "";
      inOutRows.push([
        timestamp,
        item.itemName,
        item.type,
        category,
        item.drum,
        item.length,
        data.buyerName,
        item.length,
        department,
        colorCode,
        "OUT",
        formattedDate,
      ]);
    } else {
      // fallback if not found
      inOutRows.push([
        timestamp,
        item.itemName,
        item.type,
        category,
        item.drum,
        item.length,
        data.buyerName,
        item.length,
        "",
        "",
        "OUT",
        formattedDate,
      ]);
    }
  });
  if (inOutRows.length > 0) {
    console.log("Save in production");
    if(data.packingType==="Final")
    {
    inOutSS.getRange(inOutSS.getLastRow() + 1, 1, inOutRows.length, inOutRows[0].length)
      .setValues(inOutRows);
    }
  }

  const masterFileId = '1ocr3vP60u3dbExOWJv4aRUpHcmPcr4846-1M8gFd0P4';
  const templateFileId = '1TUdj7Tx-CHp8OIXAUqjdXa2Y3zWCJMxp7ZhTN400Ovc';
  const masterSS = SpreadsheetApp.openById(masterFileId);
  const templateSS = SpreadsheetApp.openById(templateFileId);
  const templateSheet = templateSS.getSheets()[0]; // assumes first sheet is the template

  if (!templateSheet) {
    throw new Error("Template sheet not found");
  }
  // Delete old packing sheet with same PO number
  const oldSheet = masterSS.getSheetByName(data.poNumber);
  if (oldSheet) masterSS.deleteSheet(oldSheet);

  // Copy template to master
  const sheet = templateSheet.copyTo(masterSS).setName(data.poNumber);

  // Replace placeholders
  replacePackingPlaceholders(sheet, data);

  // Insert item rows
  insertPackingItemsAtPlaceholder(sheet, "[PACKING_TABLE]", data.items);

  const sheetId = sheet.getSheetId();
  return masterSS.getUrl() + "#gid=" + sheetId;
}

function replacePackingPlaceholders(sheet, data) {
  const range = sheet.getDataRange();
  const values = range.getValues();

  const replacements = {
    '[PO]': data.poNumber,
    '[PO_DATE]': formatDateStructure(data.poDate),
    '[CUSTOMER]': data.buyerName,
    '[CONSIGNEE]': data.consignee,
    '[CONSIGNEE_ADD]': data.consigneeAddress,
    '[CONTACT_PERSON1]': data.consigneeMobile1,
    '[CONTACT_PERSON2]': data.consigneeMobile2,
    '[MATERIAL]': data.material,
    '[TODAY]': formatDateStructure(new Date()),
    '[INV]': formatDateStructure(data.inv1),
    '[INV_DATE]': formatDateStructure(data.invdate1),
   
    // add more if your template has more placeholders
  };

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      let cell = values[i][j];
      if (typeof cell === 'string') {
        Object.keys(replacements).forEach(key => {
          cell = cell.replace(key, replacements[key]);
        });
        values[i][j] = cell;
      }
    }
  }
  range.setValues(values);
}
Logger.log(formatDateStructure(new Date()));
function insertPackingItemsAtPlaceholder(sheet, marker, items) {
  const markerCell = sheet.createTextFinder(marker).findNext();
  if (!markerCell) return;

  const startRow = markerCell.getRow();
  const startCol = markerCell.getColumn();
  markerCell.clearContent();

  const columnHeaders = [['Sr.No', 'Item', 'Item Type', 'Drum No.', 'Length']];
  sheet.insertRowsAfter(startRow, items.length + 2);

  // Header row
  const headerRange = sheet.getRange(startRow, startCol, 1, columnHeaders[0].length);
  headerRange.setValues(columnHeaders);
  headerRange.setFontWeight("bold").setBackground("#f2f2f2").setHorizontalAlignment("center");

  // Item rows
  const itemData = items.map((item, idx) => [
    idx + 1,
    item.itemName,
    item.type,
    item.drum,
    item.length
  ]);

  const itemRange = sheet.getRange(startRow + 1, startCol, itemData.length, columnHeaders[0].length);
  itemRange.setValues(itemData).setHorizontalAlignment("center");

  // Optional: Apply border or styling
  const fullRange = sheet.getRange(startRow, startCol, itemData.length + 1, columnHeaders[0].length);
  fullRange.setBorder(true, true, true, true, true, true);
}

function formatDateStructure(dateStr) {
  if (!dateStr) return "";

  if (Object.prototype.toString.call(dateStr) === "[object Date]") {
    const day = String(dateStr.getDate()).padStart(2, '0');
    const month = String(dateStr.getMonth() + 1).padStart(2, '0');
    const year = dateStr.getFullYear();
    return `${day}/${month}/${year}`;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [yyyy, mm, dd] = dateStr.split("-");
    return `${dd}/${mm}/${yyyy}`;
  }

  if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) {
    const [dd, mm, yyyy] = dateStr.split("-");
    return `${dd}/${mm}/${yyyy}`;
  }

  return String(dateStr);
}

function getPOItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PI Generation");
  const poDataSheet = ss.getSheetByName("PO_Data");

  const poNo = sheet.getRange("H7").getValue().toString().trim();
  if (!poNo) {
    SpreadsheetApp.getUi().alert("Please enter a PO Number in H7.");
    return;
  }

  // Read PO_Data columns B:AH
  const data = poDataSheet.getRange("B2:AH" + poDataSheet.getLastRow()).getValues().filter(r => r[0]);

  // Filter rows by PO number
  const filteredRows = data.filter(r => String(r[0]).trim() === poNo);
  if (filteredRows.length === 0) {
    SpreadsheetApp.getUi().alert("No matching data found for PO Number: " + poNo);
    return;
  }

  // Extract item data (Cols F, G, H)
  const items = filteredRows.map(r => [r[4], r[5], r[6]]);
  const freight = filteredRows[0][32] || 0; // AH column = Freight

  // Prepare display table
  const tableData = items.map((item, i) => [i + 1, item[0], item[1], item[2]]);
  tableData.push(["", "Freight", "", ""]);
  tableData.push(["", "Total Basic Value", "", ""]);
  tableData.push(["", "GST 18%", "", ""]);
  tableData.push(["", "Total", "", ""]);
  tableData.push(["", "Advance (%)", "", ""]);
  tableData.push(["", "Net Amount", "", ""]);

  // Clear old data
  sheet.getRange("E16:J").clearContent().clearDataValidations().setFontWeight("normal").setBackground(null);

  // Paste table
  sheet.getRange(16, 5, tableData.length, 4).setValues(tableData);

  const itemCount = items.length;
  const freightRow = 16 + itemCount;
  const totalRow = freightRow + 1;
  const gstRow = totalRow + 1;
  const grandTotalRow = gstRow + 1;
  const advanceRow = grandTotalRow + 1;
  const netRow = advanceRow + 1;

  // --- Formulas ---
  for (let i = 0; i < itemCount; i++) {
    const row = 16 + i;
    sheet.getRange(row, 9).setFormula(`=G${row}*H${row}`);
  }

  // Freight amount
  sheet.getRange(`I${freightRow}`).setFormula(`=${freight}`);

   // ✅ Total Basic Value = SUM of checked items + freight
  // Using FILTER to include only checked rows
  sheet.getRange(`I${totalRow}`).setFormula(`=SUM(FILTER(I16:I${freightRow-1}, J16:J${freightRow-1}=TRUE)) + I${freightRow}`);

  // GST 18% = 18% of Total Basic Value
  sheet.getRange(`I${gstRow}`).setFormula(`=I${totalRow}*0.18`);

  // Grand Total = Total Basic Value + GST
  sheet.getRange(`I${grandTotalRow}`).setFormula(`=I${totalRow}+I${gstRow}`);

  // Advance Amount = (Advance% * Total) / 100
  sheet.getRange(`I${advanceRow}`).setFormula(`=IF(G${advanceRow}>0, (G${advanceRow}/100)*I${grandTotalRow}, 0)`);

  // Net Amount = Advance Amount
  sheet.getRange(`I${netRow}`).setFormula(`=I${advanceRow}`);

  // --- Formatting ---
  // Bold summary rows
  sheet.getRange(`E${freightRow}:J${netRow}`).setFontWeight("bold");
  // Light background for summary section
  sheet.getRange(`E${freightRow}:J${netRow}`).setBackground("#f0f0f0");
  // Highlight Advance % cell (G column)
  sheet.getRange(`G${advanceRow}`).setBackground("#fff475"); // light yellow
  sheet.getRange(`G${advanceRow}`).setValue(""); // clear any old value
  // Checkboxes only for item rows
  sheet.getRange(16, 10, itemCount, 1).insertCheckboxes();
  // Remove checkbox on freight row
  sheet.getRange(`J${freightRow}`).clearDataValidations().clearContent();
  SpreadsheetApp.getUi().alert("Items, Freight, GST, Total, and Advance details loaded successfully for PO: " + poNo);
}

function copyCheckedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 16;
  const startCol = 5;
  const numCols = 6;

  // Detect last data row based on column F
  const values = sheet.getRange("F16:F" + sheet.getLastRow()).getValues();
  let lastDataRow = startRow;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0]) lastDataRow = startRow + i;
  }

  const data = sheet.getRange(startRow, startCol, lastDataRow - startRow + 1, numCols).getValues();

  // Find key rows
  const freightIndex = data.findIndex(r => String(r[1]).trim().toLowerCase().includes("freight"));
  const advanceRow = data.findIndex(r => String(r[1]).trim().toLowerCase().includes("advance"));
  if (freightIndex === -1) {
    SpreadsheetApp.getUi().alert("Freight row not found — please generate the table first.");
    return;
  }
  let advancePercent = "";
  if (advanceRow !== -1) {
    advancePercent = data[advanceRow][2] || data[advanceRow][3] || data[advanceRow][1];
  }
  const checkedItems = data
    .filter((r, i) => r[5] === true && i < freightIndex)
    .map(r => r.slice(0, 5));

  const summaryRows = data.slice(freightIndex).map(r => r.slice(0, 5));
  const advRowIndex = summaryRows.findIndex(r => String(r[1]).trim().toLowerCase().includes("advance"));
  if (advRowIndex !== -1 && advancePercent) {
    summaryRows[advRowIndex][1] = `${advancePercent}% Advance`;
  }
  const finalData = checkedItems.concat(summaryRows);
  const targetStartRow = 21;
  const targetStartCol = 12;

  // Clear old area
  const clearRange = sheet.getRange(`L${targetStartRow}:P${targetStartRow + 60}`);
  clearRange.clear({ contentsOnly: false });
  clearRange.breakApart();

  if (finalData.length > 0) {
    const pasteRange = sheet.getRange(targetStartRow, targetStartCol, finalData.length, 5);
    pasteRange.setValues(finalData);
    for (let i = 0; i < checkedItems.length; i++) {
      const row = targetStartRow + i;
      sheet.getRange(`P${row}`).setFormula(`=N${row}*O${row}`);
    }
    const freightPasteRow = targetStartRow + checkedItems.length;
    const netAmountRow = targetStartRow + finalData.length - 1;
    sheet.getRange(`L${freightPasteRow}:P${netAmountRow}`)
      .setFontWeight("bold")
      .setBackground("#ffffff");

    sheet.getRange(`L${targetStartRow}:P${netAmountRow}`)
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    const netAmountValue = sheet.getRange(`P${netAmountRow}`).getValue();
    const amountInWords = convertNumberToWords(netAmountValue);

    // Amount in Words row
    const amountWordsRow = netAmountRow + 1;
    sheet.getRange(`L${amountWordsRow}:P${amountWordsRow}`)
      .merge()
      .setValue(amountInWords)
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    // ✅ Add one empty row, then R6 → L column, S6 → M column
    const signatureRow = amountWordsRow + 2; // +1 empty row +1 data row
    const r6Value = sheet.getRange("R6").getValue();
    const s6Value = sheet.getRange("S6").getValue();
    const s7heading = sheet.getRange("S7").getValue();
      const r8name = sheet.getRange("R8").getValue();
      const s8name = sheet.getRange("S8").getValue();
      const r9Bankers = sheet.getRange("R9").getValue();
      const s9Bankers = sheet.getRange("S9").getValue();
      const r10Branch = sheet.getRange("R10").getValue();
      const s10Branch = sheet.getRange("S10").getValue();
      const s10Branch1 = sheet.getRange("S11").getValue();
      const s10Branch2 = sheet.getRange("S12").getValue();
      const r11Acc = sheet.getRange("R13").getValue();
      const s11Acc = sheet.getRange("S13").getValue();
      const r12IFC = sheet.getRange("R14").getValue();
      const s12IFC = sheet.getRange("S14").getValue();

    sheet.getRange(`L${signatureRow}`).setValue(r6Value).setFontWeight("bold");
    sheet.getRange(`M${signatureRow}`).setValue(s6Value).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+2}`).setValue(s7heading).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+3}`).setValue(r8name).setFontWeight("bold");
    sheet.getRange(`M${signatureRow+3}`).setValue(s8name).setFontWeight("bold");
    sheet.getRange(`O${signatureRow+3}`).setValue("For Vishal Cables Pvt.Ltd").setFontWeight("bold");
    sheet.getRange(`L${signatureRow+4}`).setValue(r9Bankers).setFontWeight("bold");
    sheet.getRange(`M${signatureRow+4}`).setValue(s9Bankers).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+5}`).setValue(r10Branch).setFontWeight("bold");
    sheet.getRange(`M${signatureRow+5}`).setValue(s10Branch).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+6}`).setValue(s10Branch1).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+7}`).setValue(s10Branch2).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+8}`).setValue(r11Acc).setFontWeight("bold");
    sheet.getRange(`M${signatureRow+8}`).setValue(s11Acc).setFontWeight("bold");
    sheet.getRange(`O${signatureRow+8}`).setValue("(Authorised Signatory)").setFontWeight("bold");
    sheet.getRange(`L${signatureRow+9}`).setValue(r12IFC).setFontWeight("bold");
    sheet.getRange(`M${signatureRow+9}`).setValue(s12IFC).setFontWeight("bold");
    sheet.getRange(`L${signatureRow+11}`).setValue("COMPUTER GENERATED, HENCE NO SIGNATURE REQUIRED").setFontWeight("bold").setFontColor("Red").setFontSize("16");
    // Optional: border for signature row
   // sheet.getRange(`L${signatureRow}:M${signatureRow}`)
     // .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  SpreadsheetApp.getUi().alert("Checked items + Freight + Summary rows copied successfully with Amount in Words and signature.");
}


/**
 * ✅ Converts number to Indian currency words
 */
function convertNumberToWords(amount) {
  if (isNaN(amount) || amount === 0) return "Zero Only";

  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
                "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen",
                "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

  function numToWords(num) {
    if (num < 20) return ones[num];
    if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? " " + ones[num % 10] : "");
    if (num < 1000) return ones[Math.floor(num / 100)] + " Hundred" + (num % 100 ? " and " + numToWords(num % 100) : "");
    return "";
  }

  const crore = Math.floor(amount / 10000000);
  const lakh = Math.floor((amount % 10000000) / 100000);
  const thousand = Math.floor((amount % 100000) / 1000);
  const hundred = Math.floor((amount % 1000) / 100);
  const remainder = Math.floor(amount % 100);

  let words = "";
  if (crore) words += numToWords(crore) + " Crore ";
  if (lakh) words += numToWords(lakh) + " Lakh ";
  if (thousand) words += numToWords(thousand) + " Thousand ";
  if (hundred) words += ones[hundred] + " Hundred ";
  if (remainder) words += "and " + numToWords(remainder) + " ";

  words = words.trim() + " Only";
  return "Rupees " + words;
}
