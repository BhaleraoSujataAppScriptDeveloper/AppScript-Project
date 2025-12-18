function showPOItemsWithDrumsUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Packing List Unit I");
  const poDataSheet = ss.getSheetByName("PO_Data");
  const prodSheet = ss.getSheetByName("Production-Unit I");

  const poNumber = mainSheet.getRange("B4").getValue();
  const partyName = mainSheet.getRange("B5").getValue().toString().trim().toUpperCase();

  if (!poNumber || !partyName) {
    SpreadsheetApp.getUi().alert("Please enter PO number in B4 and Party name in B5.");
    return;
  }

  // --- Get PO Data ---
  const poData = poDataSheet.getDataRange().getValues();
  const poHeaders = poData.shift(); // remove headers
  const idxPONumber = poHeaders.indexOf("PO Number");
  const idxUnit = poHeaders.indexOf("Unit");
  const idxItem = poHeaders.indexOf("Items");
  const idxItemType = poHeaders.indexOf("Item Type");
  const idxItemCode = poHeaders.indexOf("Item Code");

  // Get only items for this PO and Unit 2
  const poItems = poData.filter(r => r[idxPONumber] == poNumber && r[idxUnit] == "Unit 1")
                        .map(r => ({
                          itemName: r[idxItem],
                          itemCode: r[idxItemCode],
                          type: r[idxItemType]
                        }));

  const startRow = 16;
  const lastRow = mainSheet.getLastRow();

  // --- Clear old data but keep formatting, remove old checkboxes ---
  if (lastRow >= startRow) {
    const rangeToClear = mainSheet.getRange(startRow, 1, lastRow - startRow + 1, 5);
    rangeToClear.clearContent(); // clears only values
    rangeToClear.offset(0, 4, lastRow - startRow + 1, 1).clearDataValidations(); // remove checkboxes
  }

  if (poItems.length === 0) {
    SpreadsheetApp.getUi().alert("No items found in PO_Data for this PO and Unit II.");
    return;
  }

  // --- Get Production Data ---
  const prodData = prodSheet.getDataRange().getValues();
  const prodHeaders = prodData.shift();
  const idxItemName = prodHeaders.indexOf("Item Name");
  const idxItemCodeProd = prodHeaders.indexOf("Item Code");
  const idxDrum = prodHeaders.indexOf("Drum");
  const idxParty = prodHeaders.indexOf("Party Name");
  const idxQty = prodHeaders.indexOf("Available Qty");

  let output = [];

  poItems.forEach(poItem => {
    // --- First try matching party for this item ---
    let matchingDrums = prodData.filter(r => 
      r[idxItemCodeProd] == poItem.itemCode &&
      r[idxItemName].toString().trim().toUpperCase() === poItem.itemName.toString().trim().toUpperCase() &&
      r[idxParty].toString().trim().toUpperCase() === partyName
    );

    // --- If no rows for this item with party, fallback to Godown ---
    if (matchingDrums.length === 0) {
      matchingDrums = prodData.filter(r => 
        r[idxItemCodeProd] == poItem.itemCode &&
        r[idxItemName].toString().trim().toUpperCase() === poItem.itemName.toString().trim().toUpperCase() &&
        r[idxParty].toString().trim().toUpperCase() == "GODOWN"
      );
    }

    // --- Add output for this item ---
    if (matchingDrums.length > 0) {
      matchingDrums.forEach(d => {
        output.push([
          d[idxItemName],       // Item Name
          d[idxItemCodeProd],   // Item Code
          d[idxDrum],           // Drum
          d[idxQty],            // Available Qty
          false                 // Checkbox placeholder
        ]);
      });
    } else {
      // no data at all for this item
      output.push([
        poItem.itemName,
        poItem.itemCode,
        "",
        "",
        false
      ]);
    }
  });

  // --- Write new output ---
  if (output.length > 0) {
    const outputRange = mainSheet.getRange(startRow, 1, output.length, 5);
    outputRange.setValues(output);
    outputRange.offset(0, 4, output.length, 1).insertCheckboxes(); // insert fresh checkboxes
  }
}

function generatePackingListUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Packing List Unit I");
  const tz = ss.getSpreadsheetTimeZone();

  // Source range
  const startRow = 16;
  const dataRange = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, 5);
  const data = dataRange.getValues();

  // Filter rows where checkbox (5th col) = TRUE
  const checkedRows = data
    .filter(row => row[4] === true)
    .map(row => [row[0], row[1], row[2], row[3]]);

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No rows selected.");
    return;
  }

  // Group rows by Item + Item Type
  let grouped = {};
  checkedRows.forEach(r => {
    const key = r[0] + "||" + r[1];
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(r);
  });

  // Build final output with totals
  let output = [];
  let serials = [];
  // Target start (H23:L)
  const targetStartRow = 23;
  const targetStartCol = 9; // I column
  const totalCols = 5; // H:L (5 columns including serial)
for (let key in grouped) {
  let rows = grouped[key];
  let serialCounter = 1; // restart serial for each item group
  let startRowIndex = output.length + targetStartRow; // H23 + rows already in output

  rows.forEach(r => {
    output.push(r);
    serials.push(serialCounter++);
  });

  // Add total row with SUM formula for Length (D column)
  // The formula will sum all D cells of this group dynamically
  const totalRowIndex = output.length + targetStartRow; // current last row + 1
  output.push(["Total", "", "", `=SUM(L${startRowIndex}:L${totalRowIndex-1})`]);
  serials.push(""); // no serial for total row
}

  // --- Clear old data including formatting ---
  sheet.getRange(targetStartRow, 8, sheet.getMaxRows() - targetStartRow + 1, totalCols).clear();

  // Write serials in H, data in I:L
  const serialRange = sheet.getRange(targetStartRow, 8, serials.length, 1);
  serialRange.setValues(serials.map(s => [s]));

  const dataRangeOutput = sheet.getRange(targetStartRow, targetStartCol, output.length, 4);
  dataRangeOutput.setValues(output);

  // Apply thick border to the full table (H:L)
  sheet.getRange(targetStartRow, 8, output.length, totalCols).setBorder(
    true, true, true, true, true, true,
    "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
 
  // Set font: Cambria, size 15, bold for the entire table
  sheet.getRange(targetStartRow, 8, output.length, totalCols)
    .setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold");

  // --- Add Invoice details text ---
  const invoiceNo = sheet.getRange("B12").getValue();
  const invoiceDateRaw = sheet.getRange("D12").getValue();
  const invoiceDate = invoiceDateRaw
    ? Utilities.formatDate(new Date(invoiceDateRaw), tz, "DD-MM-yyyy")
    : "";

const invoiceRow = targetStartRow + output.length + 1;

const invoiceLine = "Invoice No. " ;
const invoiceLineData =  invoiceNo 

const invoiceDateRow = targetStartRow + output.length + 1+1;
const invoiceDateLine="Dated: ";
const invoiceDateData =  invoiceDate 

// H column → full line
sheet.getRange(invoiceRow, 8)
  .setValue(invoiceLine)
  .setFontFamily("Cambria")
  .setFontSize(15)
  .setFontWeight("bold")
  .setWrap(false)             // Disable wrapping
  .setHorizontalAlignment("left");
sheet.getRange(invoiceRow, 9)
  .setValue(invoiceLineData)
  .setFontFamily("Cambria")
  .setFontSize(15)
  .setFontWeight("bold")
  .setWrap(false)             // Disable wrapping
  .setHorizontalAlignment("left");

sheet.getRange(invoiceDateRow, 8)
  .setValue(invoiceDateLine)
  .setFontFamily("Cambria")
  .setFontSize(15)
  .setFontWeight("bold")
  .setWrap(false)             // Disable wrapping
  .setHorizontalAlignment("left");


  // I → Dated
 // sheet.getRange(invoiceRow, 9).setValue("Dated: " + invoiceDate);

  // K → For Vishal Cables Pvt.Ltd.
  sheet.getRange(invoiceRow, 11).setValue("For Vishal Cables Pvt.Ltd.");

  // Leave 2 blank rows, then add Authorised Signatory in K column
  const signRow = invoiceRow + 3; 
  sheet.getRange(signRow, 11).setValue(" Authorised Signatory").setFontFamily("Cambria").setFontSize(15).setFontWeight("bold");

  // Set font for invoice lines
  sheet.getRange(invoiceRow, 8, 1, 3).setFontFamily("Cambria").setFontSize(15).setFontWeight("bold").setWrap(false);
  sheet.getRange(invoiceRow, 11).setFontFamily("Cambria").setFontSize(15).setFontWeight("bold");
  sheet.getRange(signRow, 12).setFontFamily("Cambria").setFontSize(15).setFontWeight("bold");
  sheet.getRange(invoiceRow-1, 8, 6, 5).setBorder(
  true,  // top
  true,  // left
  true,  // bottom
  true,  // right
  false, // no vertical inside
  false, // no horizontal inside
  "black",
  SpreadsheetApp.BorderStyle.SOLID_MEDIUM
);
}


function getPackingListDataFromSheetUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Packing List Unit I");
  
  const buyer = sheet.getRange("I10").getValue(); 
  const podate = sheet.getRange("L14").getValue();
  const poNo = sheet.getRange("I14").getValue();
  const consignee = sheet.getRange("B7").getValue();
  const ConsigneeAddress = sheet.getRange("B8").getValue();
  const contactPerson = sheet.getRange("B11").getValue();
  const material = sheet.getRange("I12").getValue();
  const invoiceno = sheet.getRange("B12").getValue();
  const IncoiceDate = sheet.getRange("D12").getValue();
  const actualBuyer = sheet.getRange("H2").getValue();
  
  const startRow = 23;
  const startCol = 8; // Column I
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  
  const itemsData = sheet.getRange(startRow, startCol, numRows, 5).getValues();
  
  // Filter rows where column I contains "Core"
  const filteredItems = itemsData.filter(row => row[1] && row[1].toString().toLowerCase().includes("core"));
  
  // Log length for info
  //Logger.log("Number of filtered items: " + filteredItems.length);
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  return {
    timestamp,
    poNo,
    podate,
    buyer,
    itemDetails: filteredItems,  // return array, not JSON string
    consignee,
    ConsigneeAddress,
    contactPerson,
    material,
    invoiceno,
    IncoiceDate,
    actualBuyer
  };
}


function getFinalDataUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Packing List Unit I");
  
  const buyer = sheet.getRange("O10").getValue(); 
  const podate = sheet.getRange("R14").getValue();
  const poNo = sheet.getRange("O14").getValue();
  const consignee = sheet.getRange("O16").getValue();
  const ConsigneeAddress = sheet.getRange("O17").getValue();
  const contactPerson = sheet.getRange("Q17").getValue();
  const material = sheet.getRange("O12").getValue();
  const invoiceno = sheet.getRange("T1").getValue();
  const Invoice_Date = sheet.getRange("T2").getValue();
  const actualBuyer = sheet.getRange("N2").getValue();
  
if (!invoiceno && !Invoice_Date) {
  SpreadsheetApp.getUi().alert("Please enter Invoice No and Invoice Date");
  return;
} else if (!invoiceno) {
  SpreadsheetApp.getUi().alert("Please enter Invoice No");
  return;
} else if (!Invoice_Date) {
  SpreadsheetApp.getUi().alert("Please enter Invoice Date");
  return;
}
  const startRow = 23;
  const startCol = 14; // Column I
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - startRow + 1;
  
  const itemsData = sheet.getRange(startRow, startCol, numRows, 5).getValues();
  
  // Filter rows where column I contains "Core"
  const filteredItems = itemsData.filter(row => row[1] && row[1].toString().toLowerCase().includes("core"));
  
  // Log length for info
  //Logger.log("Number of filtered items: " + filteredItems.length);
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  return {
    timestamp,
    poNo,
    podate,
    buyer,
    itemDetails: filteredItems,  // return array, not JSON string
    consignee,
    ConsigneeAddress,
    contactPerson,
    material,
    invoiceno,
    Invoice_Date,
    actualBuyer
  };
}


function SavePackinglistDraftUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName("Packing History Unit I");
  const packingListData = getPackingListDataFromSheetUnitI();
  
  packingListData.itemDetails.forEach(item => {
    draftSheet.appendRow([
      packingListData.timestamp,
      packingListData.poNo,
      packingListData.podate,
      packingListData.buyer,
      packingListData.consignee,
      packingListData.ConsigneeAddress,
      item[0],  // Item no
      item[1],  // Item name
      item[2],  // Itme code
      item[3],  // drum
      item[4],  // Length
      packingListData.material,
      packingListData.contactPerson,
      packingListData.invoiceno,
      packingListData.IncoiceDate,
     "Draft",
     packingListData.actualBuyer

    ]);
  });
}

function SavePackinglistFinalUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName("Packing History Unit I");
  const packingListData = getFinalDataUnitI();
  
  packingListData.itemDetails.forEach(item => {
    draftSheet.appendRow([
      packingListData.timestamp,
      packingListData.poNo,
      packingListData.podate,
      packingListData.buyer,
      packingListData.consignee,
      packingListData.ConsigneeAddress,
      item[0],  // Item no
      item[1],  // Item name
      item[2],  // Itme code
      item[3],  // drum
      item[4],  // Length
      packingListData.material,
      packingListData.contactPerson,
      packingListData.invoiceno,
      packingListData.Invoice_Date,
     "Final",
     packingListData.actualBuyer
    ]);
  });
}

function retrivePackinglistHistoryUnitI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName("Packing History Unit I");
  const formSheet = ss.getSheetByName("Packing List Unit I");
  const poNo = formSheet.getRange("P2").getValue();
  const buyerName = formSheet.getRange("N2").getValue();

  if (!poNo || !buyerName) {
    SpreadsheetApp.getUi().alert("Please select PO number and buyer name.");
    return;
  }

  const data = draftSheet.getDataRange().getValues();
  const dataWithoutHeader = data.slice(1);

  const poIndex = 1,
        buyerIndex = 16,
        statusIndex = 15;

  const matchedRows = dataWithoutHeader.filter(row =>
    String(row[poIndex]).trim().toLowerCase() === String(poNo).trim().toLowerCase() &&
    String(row[buyerIndex]).trim().toLowerCase() === String(buyerName).trim().toLowerCase() &&
    String(row[statusIndex]).trim().toLowerCase() === "draft"
  );

  if (matchedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No matching packing list found.");
    return;
  }
  // Clear old data at N23:R
  const targetStartRow = 23;
  formSheet.getRange(targetStartRow, 14, formSheet.getMaxRows() - targetStartRow + 1, 5).clear();

  // Fill metadata fields from first matched row
  const firstRow = matchedRows[0];
  formSheet.getRange("O10").setValue(firstRow[3]); // Buyer
  formSheet.getRange("O14").setValue(firstRow[1]); // PO No
  formSheet.getRange("R14").setValue(firstRow[2]); // PO Date
  formSheet.getRange("O12").setValue(firstRow[11]); // Material
  formSheet.getRange("O16").setValue(firstRow[4]); // Consignee
  formSheet.getRange("O17").setValue(firstRow[5]); // Consignee Address
  formSheet.getRange("Q17").setValue(firstRow[12]); // Contact Person
  let invoice_no=(firstRow[13]); // Invoice No from column N
  let invoice_date=(firstRow[14]); 
  if (invoice_date) {
  const tz = ss.getSpreadsheetTimeZone();
  invoice_date = Utilities.formatDate(new Date(invoice_date), tz, "dd-MMM-yyyy");
}

  // Group rows by item name + item code (columns H=7, I=8)
  let grouped = {};
  matchedRows.forEach(row => {
    const key = row[7] + "||" + row[8];
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(row);
  });

  let output = [];
  let serials = [];

  const serialCol = 14; // column N
  const dataStartCol = 15; // column O for item data
  const totalCols = 4; // O, P, Q, R

  for (let key in grouped) {
    const groupRows = grouped[key];
    let serialCounter = 1;
    let startRowIndex = output.length + targetStartRow;

    groupRows.forEach(r => {
      serials.push([serialCounter++]);
      output.push([r[7], r[8], r[9], r[10]]); // item name, code, drum, length
    });

    // Add total row with SUM formula for Length column (column Q, which is 4th in output range)
    const totalRowIndex = output.length + targetStartRow;
    output.push(["Total", "", "", `=SUM(R${startRowIndex}:R${totalRowIndex - 1})`]);
    serials.push([""]); // no serial for total row
  }

  // Write serial numbers into column N (14)
  formSheet.getRange(targetStartRow, serialCol, serials.length, 1).setValues(serials);

  // Write item data into columns O to R
  formSheet.getRange(targetStartRow, dataStartCol, output.length, totalCols).setValues(output);

  // Apply thick border (N:R)
  formSheet.getRange(targetStartRow, 14, output.length, 5).setBorder(
    true, true, true, true, true, true,
    "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // Set font: Cambria, 15, bold for the whole table
  formSheet.getRange(targetStartRow, 14, output.length, 5).setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold");

  // Invoice footer
  const invoiceRow = targetStartRow + output.length + 1;
  const invoiceDateRow=targetStartRow+output.length+2;
 // const invoiceNo = formSheet.getRange("B12").getValue();
 // const invoiceDateRaw = formSheet.getRange("D12").getValue();
 // const tz = ss.getSpreadsheetTimeZone();
 // const invoiceDate = invoiceDateRaw ? Utilities.formatDate(new Date(invoiceDateRaw), tz, "dd-MMM-yyyy") : "";

  const invoiceLine = "Invoice No. ";
  const invoiceDateLine="Dated ";

  // Invoice No line in column N
  formSheet.getRange(invoiceRow, 14)
    .setValue(invoiceLine)
    .setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold")
    .setWrap(false)
    .setHorizontalAlignment("left");

    formSheet.getRange(invoiceDateRow, 14)
    .setValue(invoiceDateLine)
    .setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold")
    .setWrap(false)
    .setHorizontalAlignment("left");

  // "For Vishal Cables Pvt.Ltd." in column P
  formSheet.getRange(invoiceRow, 17)
    .setValue("For Vishal Cables Pvt.Ltd.")
    .setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold");

  // Authorised Signatory in column P, 3 rows below invoice
  formSheet.getRange(invoiceRow + 3, 17)
    .setValue("Authorised Signatory")
    .setFontFamily("Cambria")
    .setFontSize(15)
    .setFontWeight("bold");


    formSheet.getRange(invoiceRow-1, 14, 6, 5).setBorder(
  true,  // top
  true,  // left
  true,  // bottom
  true,  // right
  false, // no vertical inside
  false, // no horizontal inside
  "black",
  SpreadsheetApp.BorderStyle.SOLID_MEDIUM
);
}
