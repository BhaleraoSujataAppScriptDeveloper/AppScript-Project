function doGet(e) {
  var page = e.parameter.Page || "IN";  // default = form1
  if (page === "IN") {
    return HtmlService.createHtmlOutputFromFile("IN");
  } else if (page === "StockEntry") {
    return HtmlService.createHtmlOutputFromFile("StockEntry");
  } 
  else {
    return HtmlService.createHtmlOutputFromFile("Invalid form");
  }
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Production System")
    .addItem("Open Transfer Form", "openTransferForm")
    .addToUi();
}
function openTransferForm() {
  const html = HtmlService.createHtmlOutputFromFile("ItemTransfer") // your .html filename
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Transfer Form");
}
function OpenStockEntryForm() {
  const html = HtmlService.createHtmlOutputFromFile("StockEntry")
    .setTitle("Stock Details Form")
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Stock Details Form");
}
function getAllItems() 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item Master");
  const data = sheet.getDataRange().getValues();
  const items = [...new Set(data.slice(1).map(r => r[0]))]; // Column A = Item Name
  return items.sort();
}
function getCodesByItem(itemName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item Master");
  const data = sheet.getDataRange().getValues();
  const codes = data.slice(1)
                    .filter(r => r[0] === itemName)
                    .map(r => r[1]); // Column B = Code
  return codes.sort();
}

function getCategory(itemName, itemCode) 
{
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item Master");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === itemName && data[i][1] === itemCode) {
      return data[i][5]; // Column H = Category
    }
  }
  return "";
}

function testCategory() {
  const category = getCategory("1 Core x 1 Sq mm", "Y");
  Logger.log(category);
}
function testDrums() {
  const category = getDrumNumbers("IS7098");
  Logger.log(category);
}
function checkDrumNoUnique(itemName, itemCode, drumNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item Master");
  const data = sheet.getDataRange().getValues();
  
  // Get category for given item name & code
  let category = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === itemName && data[i][1] === itemCode) {
      category = data[i][5]; // Assuming Category is column F (index 5)
      break;
    }
  }
  console.log("Category "+category);
  if (!category) return false; // No category found → fail validation

  // Check if any row has same category and same drum number
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === category && String(data[i][2]) === String(drumNo)) {
      return false; // Drum number already exists for this category
    }
  }
  return true; // Allowed
}
function submitProduction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("In-Out");
  const itemSheet = ss.getSheetByName("Item Master");

  const lastRow = itemSheet.getLastRow();
  const imData = itemSheet.getRange(3, 1, lastRow - 2, 11).getValues(); // A–K

  let found = false;
  let updatedStock = 0;
  let blankDrumRow = -1;
  let category = "";

  // 1️⃣ Find category for this item+code
  for (let i = 0; i < imData.length; i++) {
    if (imData[i][0] === data.itemName && imData[i][1] === data.itemCode) {
      category = imData[i][5]; // Col F = Category
      break;
    }
  }
  if (!category) throw new Error("❌ Category not found for this Item.");

  // 2️⃣ Loop through rows to check/update
  for (let i = 0; i < imData.length; i++) {
    const [itemName, itemCode, drumNo, stock, , rowCategory] = imData[i];

    // ❌ Prevent same drum in same category for same item+code
    if (
      rowCategory === category &&
      drumNo &&
      drumNo.toString() === data.drumNo.toString() &&
      !(itemName === data.itemName && itemCode === data.itemCode && rowCategory === category)
    ) {
      throw new Error(`❌ Drum No ${data.drumNo} is already used for this item & category (${category}).`);
    }

    // ✅ Exact match: same item+code+category+drum
    if (
      itemName === data.itemName &&
      itemCode === data.itemCode &&
      rowCategory === category &&
      drumNo &&
      drumNo.toString() === data.drumNo.toString()
    ) {
      let currentStock = stock ? Number(stock) : 0;

      if (currentStock === 0 && data.status !== "IN") {
        throw new Error(`❌ Drum No ${data.drumNo} is finished and cannot be used again.`);
      }
      if (data.status === "IN") {
        currentStock += Number(data.length);
      } else if (data.status === "OUT") {
        currentStock -= Number(data.length);
      }
      if (currentStock < 0) throw new Error(`❌ Not enough stock. Available: ${stock || 0}`);
      updatedStock = Math.max(currentStock, 0);
      itemSheet.getRange(i + 3, 4).setValue(updatedStock); // Col D = Stock
      found = true;
      break;
    }

    // 📝 Save blank drum row for same item+code+category
    if (
      itemName === data.itemName &&
      itemCode === data.itemCode &&
      rowCategory === category &&
      (!drumNo || drumNo.toString().trim() === "")
    ) {
      blankDrumRow = i + 3;
    }
  }

  // 3️⃣ Assign to blank drum row if available
  if (!found && blankDrumRow > -1) {
    if (data.status !== "IN") throw new Error("❌ Cannot do OUT entry. Drum not found.");

    const currentStock = Number(itemSheet.getRange(blankDrumRow, 4).getValue()) || 0;
    const newStock = currentStock + Number(data.length);

    itemSheet.getRange(blankDrumRow, 3).setValue(data.drumNo); // Col C = Drum No
    itemSheet.getRange(blankDrumRow, 4).setValue(newStock);     // Col D = Stock
    updatedStock = newStock;

    logTransaction(prodSheet, data);
    return `✅ Assigned Drum No ${data.drumNo} to blank row and updated stock.`;
  }

  // 4️⃣ If still not found, append new row
  if (!found) {
    if (data.status !== "IN") throw new Error("❌ Cannot do OUT entry. Drum not found.");
    itemSheet.appendRow([
      data.itemName,
      data.itemCode,
      data.drumNo,
      Number(data.length),
      "Meter",
      category,
      "",
      "",
      "",
      "",
      ""
    ]);
    updatedStock = Number(data.length);
  }

  // 5️⃣ Log the transaction
  logTransaction(prodSheet, data);
  return `✅ Stock updated for Drum No ${data.drumNo}. Current Stock: ${updatedStock}`;
}
// Helper: log production entry
function logTransaction(sheet, data) {
  
  const nextRow = sheet.getLastRow() + 1;
 sheet.getRange(nextRow, 1, 1, 10).setValues([[
  new Date(),
  data.itemName,
  data.itemCode,
  data.drumNo,
  data.length,
  data.status,
  data.color1,
  data.dept,
  data.party,
  data.prodtype
]]);
}

function updateItemSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName("Item Master");
  const summarySheet = ss.getSheetByName("Summary");

  const data = itemSheet.getRange(3, 1, itemSheet.getLastRow() - 2, 6).getValues();
  const summary = {};

  data.forEach(row => {
    const [itemName, itemCode, unit, , , stock] = row;
    if (!itemName || !itemCode || !unit) return;

    const key = itemName + "|" + itemCode + "|" + unit;
    summary[key] = (summary[key] || 0) + (Number(stock) || 0);
  });

  // Clear and set headers
  summarySheet.clear();
  summarySheet.appendRow(["Item Name", "Item Code", "Unit", "Total Stock"]);

  // Append summarized data
  for (let key in summary) {
    const [itemName, itemCode, unit] = key.split("|");
    summarySheet.appendRow([itemName, itemCode, unit, summary[key]]);
  }
}
function getAllUniqueCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item Master"); // Change sheet name if different
  const data = sheet.getRange(3, 2, sheet.getLastRow() - 2, 1).getValues(); // Assuming codes are in column B
  const uniqueCodes = [...new Set(data.flat().filter(String))];
  return uniqueCodes.sort();
}
function getPendingBuyers123() {
  const ss = SpreadsheetApp.openById("1ZCjxuaee-VOTYVIv83iACqV2XO3ZKLtbQup_m2MJpJw"); // Change to actual ID
  const sheet = ss.getSheetByName("PO Summary");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // No data

  const data = sheet.getRange(3, 2, lastRow - 2, 12).getValues(); 
  // Col B = Buyer Name, Col M = Order Status (M is column 13, but we got 12 columns starting from B so M becomes index 11)

  const buyers = data
    .filter(row => row[10] && row[10].toString().toLowerCase() === "pending" ||row[11].toString().toLowerCase() === "completed") // Status is completed
    .map(row => row[0]) // Buyer Name
    .filter(name => name && name.toString().trim() !== "");

  // Remove duplicates
  return [...new Set(buyers)].sort();
}
function getPendingBuyers() {
  const ss = SpreadsheetApp.openById("1ZCjxuaee-VOTYVIv83iACqV2XO3ZKLtbQup_m2MJpJw");
  const poSummary = ss.getSheetByName("PO Summary");
  const poData = ss.getSheetByName("PO_Data");

  // --- Get PO_Data ---
  const lastRowData = poData.getLastRow();
  const poDataValues = poData.getRange(2, 2, lastRowData - 1, 8).getValues();
  // Col B = PO Number (index 0), E = Buyer Name (index 3), I = Unit (index 7)

  // Collect all PO numbers belonging to Unit II
  const unit2POs = new Set(
    poDataValues
      .filter(r => r[7] && r[7].toString().trim().toLowerCase() === "unit 2")
      .map(r => r[0]) // PO Number
  );

  if (unit2POs.size === 0) return [];

  // --- Get PO_Summary ---
  const lastRowSummary = poSummary.getLastRow();
  if (lastRowSummary < 3) return [];

  const summaryValues = poSummary.getRange(3, 2, lastRowSummary - 2, 12).getValues();
  // Col B = Buyer Name (index 0), Col C = PO Number (index 1), Col M = Order Status (index 11)

  const buyers = summaryValues
    .filter(r =>
      r[1] && unit2POs.has(r[1]) && // PO is in Unit II
      r[10] && ["pending", "completed"].includes(r[10].toString().toLowerCase())
    )
    .map(r => r[0]) // Buyer Name
    .filter(name => name && name.toString().trim() !== "");
Logger.log([...new Set(buyers)].sort());
  return [...new Set(buyers)].sort();
}

function getAvailableStockByCategoryFormatted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const openingSheet = ss.getSheetByName("Opening Stock Details");
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const outputSheet = ss.getSheetByName("Stock By Category") || ss.insertSheet("Stock By Category");
  const selectedDate = outputSheet.getRange("B1").getValue(); // Only one date
  if (!selectedDate) {
    SpreadsheetApp.getUi().alert("Please enter the Date in B1 of Stock By Category sheet.");
    return;
  }
  // Clear content & formatting from second row onwards
  outputSheet.getRange(2, 1, outputSheet.getMaxRows() - 1, outputSheet.getMaxColumns()).clearContent().clearFormat();
  // Normalize date object
  const endDate = new Date(selectedDate);
  endDate.setHours(23, 59, 59, 999);

  // --- Get data ---
  const openingData = openingSheet.getRange(3, 1, openingSheet.getLastRow() - 2, 10).getValues();
  const prodData = prodSheet.getRange(3, 1, prodSheet.getLastRow() - 2, 12).getValues();

  const stockMap = {}; // category -> itemName|itemCode -> stock

  // --- Opening Stock ---
  openingData.forEach(row => {
    const category = row[2];
    const itemName = row[0];
    const itemCode = row[1];
    const length = Number(row[4]) || 0;
    const prodDate = row[9] ? new Date(row[9]) : null; // J column
    if (!category || !itemName || !itemCode || !prodDate) return;

    prodDate.setHours(0,0,0,0);
    if (prodDate > endDate) return; // Filter only till selected date

    if (!stockMap[category]) stockMap[category] = {};
    const key = itemName + "||" + itemCode;
    if (!stockMap[category][key]) stockMap[category][key] = 0;
    stockMap[category][key] += length;
  });

  // --- Production/Dispatch ---
  prodData.forEach(row => {
    const status = String(row[10]).trim().toUpperCase(); // Status
    const category = row[3];
    const itemName = row[1];
    const itemCode = row[2];
    const length = Number(row[5]) || 0;
    const prodDate = row[11] ? new Date(row[11]) : null; // L column
    if (!category || !itemName || !itemCode || !prodDate) return;

    prodDate.setHours(0,0,0,0);
    if (prodDate > endDate) return; // Filter only till selected date

    if (!stockMap[category]) stockMap[category] = {};
    const key = itemName + "||" + itemCode;
    if (!stockMap[category][key]) stockMap[category][key] = 0;

    if (status === "IN") stockMap[category][key] += length;
    else if (status === "OUT" || status==="SCRAP") {
      stockMap[category][key] -= length;
      if (stockMap[category][key] < 0) stockMap[category][key] = 0;
    }
  });

  // --- Write to output sheet ---
  let startCol = 1;
  const maxCol = 26; // Z
  let grandTotalAllCategories = 0; // ✅ keep track of total across all categories

  for (const category in stockMap) {
    const items = stockMap[category];
    const itemKeys = Object.keys(items);

    outputSheet.getRange(2, startCol, 1, 3).merge()
      .setValue(category)
      .setFontWeight("bold")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setBackground("#d9ead3")
      .setFontColor("#1155cc");

    outputSheet.getRange(3, startCol, 1, 3)
      .setValues([["Item Name", "Item Code", "Available Stock"]])
      .setFontWeight("bold")
      .setFontSize(11)
      .setHorizontalAlignment("center")
      .setBackground("#cfe2f3");

    let currentRow = 4;
    let categoryTotal = 0;

    itemKeys.forEach(k => {
      const [itemName, itemCode] = k.split("||");
      const stock = items[k];
      outputSheet.getRange(currentRow, startCol).setValue(itemName);
      outputSheet.getRange(currentRow, startCol + 1).setValue(itemCode);
      const stockCell = outputSheet.getRange(currentRow, startCol + 2);
      stockCell.setValue(stock);
      if (stock < 0) stockCell.setFontColor("red").setFontWeight("bold");
      categoryTotal += stock;
      currentRow++;
    });

    outputSheet.getRange(currentRow, startCol, 1, 2).merge()
      .setValue("Grand Total")
      .setFontWeight("bold")
      .setFontSize(11)
      .setHorizontalAlignment("center")
      .setBackground("#ffe599");
    outputSheet.getRange(currentRow, startCol + 2)
      .setValue(categoryTotal)
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground("#ffe599");

    grandTotalAllCategories += categoryTotal; // ✅ add to final grand total

    startCol += 4;
    if (startCol + 2 > maxCol) startCol = 1;
  }

  // ✅ Write Final Grand Total at the bottom-left
  outputSheet.getRange("F1").setValue(grandTotalAllCategories);
  
  SpreadsheetApp.getUi().alert("Stock by category generated up to selected date with grand totals (per category + final total).");
}


