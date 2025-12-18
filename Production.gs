function getDrumDetails(itemName, itemCode, category, drumNo) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Production/Dispatch");

  const data = sheet.getRange(3, 1, sheet.getLastRow() - 1, 12).getValues();

  const row = data.find(r =>
    r[1] == itemName &&
    r[2] == itemCode &&
    r[3] == category &&
    r[4] == drumNo
  );
Logger.log("Row"+row);
  if (!row) return null;

  return {
    itemName: row[1],
    itemCode: row[2],
    category: row[3],
    drumNo: row[4],
    prodLength: row[5],
    partyName: row[6],
    partyLength: row[7],
    department: row[8],
    colorCode: row[9]
  };
}
function test1()
{
  getDrumDetails("4 Core x 10 Sq mm","A2XFY","IS7098","595")
}


function getAvailableDrumsForOut(itemName, itemCode, category) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Production/Dispatch");

  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 12).getValues();

  // STEP 1: Get all IN drums
  const inDrums = data
    .filter(r =>
      r[1] === itemName &&      // Item Name
      r[2] === itemCode &&      // Item Code
      r[3] === category &&      // Category
      r[10] === "IN"            // Status = IN
    )
    .map(r => r[4]);            // Drum No

  // STEP 2: Get all OUT drums
  const outDrums = data
    .filter(r =>
      r[1] === itemName &&
      r[2] === itemCode &&
      r[3] === category &&
      r[10] === "OUT"
    )
    .map(r => r[4]);

  // STEP 3: Remove OUT drums from IN list
  const availableDrums = inDrums.filter(drum => !outDrums.includes(drum));

  //Logger.log("Available Drums (IN but not OUT): " + availableDrums);

  return availableDrums;
}

function test()
{
  getAvailableDrumsForOut("4 Core x 35 Sq mm","A2XFY", "IS7098")
}

function getUniqueItemNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Master");
  const data = sheet.getRange("A3:A" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(Boolean))];
}

function getUniqueItemCodes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Master");
  const data = sheet.getRange("B3:B" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(Boolean))];
}
function getCategoryByItemNameAndCode(itemName, itemCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item Master");
  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 4).getValues(); // Columns A-D

  for (let i = 0; i < data.length; i++) {
    let sheetItemName = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    let sheetItemCode = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
    let inputItemName = itemName ? itemName.toString().trim().toLowerCase() : "";
    let inputItemCode = itemCode ? itemCode.toString().trim().toLowerCase() : "";

    Logger.log(`Checking row ${i + 3}: Name='${sheetItemName}', Code='${sheetItemCode}'`);

    if (sheetItemName === inputItemName && sheetItemCode === inputItemCode) {
      Logger.log("Match found. Category: " + data[i][3]);
      return data[i][3]; // Category in D
    }
  }
  Logger.log("No match found.");
  return "";
}
function saveStockEntry(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Opening Stock Details");

  if (!sheet) {
    throw new Error('Sheet "Opening Stock Details" not found!');
  }
 if (!data.category || data.category.toString().trim() === "") {
    throw new Error("Category is mandatory.");
  }
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    lastRow = 1; // Only header present
  }

  // Get existing data: Drum No is in column 4 (index 3), Category is column 3 (index 2)
  const existingData = (lastRow > 1)
    ? sheet.getRange(2, 1, lastRow - 1, 7).getValues()
    : [];

  const existingKeys = new Set();
  existingData.forEach(row => {
    const key = (row[2] + "|" + row[3]).toLowerCase();
    existingKeys.add(key);
  });

  const rowsToAppend = [];
  for (const drum of data.drums) {
    const key = (data.category + "|" + drum.drumNo).toLowerCase();

    if (existingKeys.has(key)) {
      throw new Error(`Drum No "${drum.drumNo}" has already been assigned to the category "${data.category}".`);
    }
    rowsToAppend.push([
      data.itemName,
      data.itemCode,
      data.category,
      drum.drumNo,
      drum.prodLength,
      drum.partyName,
      drum.partyLength,
      drum.dept,
      drum.code
    ]);
    existingKeys.add(key);
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
         .setValues(rowsToAppend);
  }
}

function saveProductionEntry(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inOutSheet = ss.getSheetByName("In-Out");
  const openingStockSheet = ss.getSheetByName("Opening Stock Details");

  if (!inOutSheet) throw new Error('Sheet "In-Out" not found!');
  if (!openingStockSheet) throw new Error('Sheet "Opening Stock Details" not found!');

  // --- Map for existing IN stock: key → { dept, code } ---
  const existingMap = new Map();

  // --- Collect from In-Out sheet ---
  let lastRowInOut = inOutSheet.getLastRow();
  if (lastRowInOut > 1) {
    const inOutData = inOutSheet.getRange(2, 1, lastRowInOut - 1, 12).getValues(); // A-L
    inOutData.forEach(row => {
      const status   = row[10]; // Col K → Status (IN/OUT)
      const itemName = row[1];  // Col B
      const itemCode = row[2];  // Col C
      const drumNo   = row[4];  // Col E
      const dept     = row[8];  // Col I
      const code     = row[9];  // Col J

      if (status === "IN" && itemName && itemCode && drumNo) {
        const key = (itemName + "|" + itemCode + "|" + drumNo).toLowerCase();
        existingMap.set(key, { dept, code });
      }
    });
  }

  // --- Collect from Opening Stock sheet ---
  let lastRowOpening = openingStockSheet.getLastRow();
  if (lastRowOpening > 1) {
    const openingData = openingStockSheet.getRange(2, 1, lastRowOpening - 1, 9).getValues(); // A-I
    openingData.forEach(row => {
      const itemName = row[0]; // Col A
      const itemCode = row[1]; // Col B
      const drumNo   = row[3]; // Col D
      const dept     = row[7]; // adjust if needed
      const code     = row[8]; // adjust if needed

      if (itemName && itemCode && drumNo) {
        const key = (itemName + "|" + itemCode + "|" + drumNo).toLowerCase();
        if (!existingMap.has(key)) {
          existingMap.set(key, { dept, code });
        }
      }
    });
  }

  const rowsToAppend = [];

  for (const drum of data.drums) {
    const key = (data.itemName + "|" + data.itemCode + "|" + drum.drumNo).toLowerCase();

    if (data.status === "IN") {
      // Prevent duplicate insertion for IN
      if (existingMap.has(key)) {
        throw new Error(`Drum "${drum.drumNo}" for item "${data.itemName}" (code ${data.itemCode}) already exists!`);
      }
    } else if (data.status === "OUT" || data.status==="SCRAP") {
      // Ensure drum exists before OUT
      if (!existingMap.has(key)) {
        throw new Error(`Drum "${drum.drumNo}" for item "${data.itemName}" (code ${data.itemCode}) not found in stock!`);
      }
      const stored = existingMap.get(key);

      // Auto-fill dept/code if user left blank
      if (!drum.dept) drum.dept = stored.dept;
      if (!drum.code) drum.code = stored.code;

      // Validate if user typed wrong values
      if (drum.dept !== stored.dept || drum.code !== stored.code) {
        throw new Error(
          `Mismatch for Drum "${drum.drumNo}" of item "${data.itemName}" (code ${data.itemCode}): 
           Expected Dept "${stored.dept}", Code "${stored.code}", but got Dept "${drum.dept}", Code "${drum.code}".`
        );
      }
    }

    // Prepare row to append
    rowsToAppend.push([
      Utilities.formatDate(new Date(), "Asia/Kolkata", "dd/MM/yyyy HH:mm:ss"),
      data.itemName,
      data.itemCode,
      data.category,
      drum.drumNo,
      drum.prodLength,
      drum.partyName,
      drum.partyLength,
      drum.dept,
      drum.code,
      data.status,
      data.productionDate
    ]);

    // Add new IN keys
    if (data.status === "IN") {
      existingMap.set(key, { dept: drum.dept, code: drum.code });
    }
  }

  // Append rows to In-Out sheet
  if (rowsToAppend.length > 0) {
    inOutSheet.getRange(inOutSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
              .setValues(rowsToAppend);
  }
}

function generateStockSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const openingSheet = ss.getSheetByName("Opening Stock Details");
  const prodSheet = ss.getSheetByName("Production/Dispatch");

  // Safely load Opening Stock
  let openingData = [];
  if (openingSheet.getLastRow() > 1) {
    openingData = openingSheet.getRange(2, 1, openingSheet.getLastRow() - 1, 9).getValues();
  }

  // Safely load Production/Dispatch
  let prodData = [];
  if (prodSheet.getLastRow() > 1) {
    prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 11).getValues();
  }

  const stockMap = new Map();

  // Helper to create unique key
  const makeKey = (itemName, itemCode, category, drum, dept, color) =>
    [itemName, itemCode, category, drum, dept, color].join("|");

  // Load opening stock
  openingData.forEach(row => {
    const [itemName, itemCode, category, drum, length, , , dept, color] = row;
    const key = makeKey(itemName, itemCode, category, drum, dept, color);
    stockMap.set(key, (stockMap.get(key) || 0) + (Number(length) || 0));
  });

  // Apply production IN/OUT
  prodData.forEach(row => {
    const [, itemName, itemCode, category, drum, prodLength, , , dept, color, status] = row;
    const key = makeKey(itemName, itemCode, category, drum, dept, color);
    let qty = Number(prodLength) || 0;
    if (status === "OUT" || status=== "SCRAP") qty = -qty;
    stockMap.set(key, (stockMap.get(key) || 0) + qty);
  });

  // Prepare summary data
  const summary = [];
  for (let [key, qty] of stockMap.entries()) {
    if (qty !== 0) {
      const parts = key.split("|");
      summary.push([...parts, qty]);
    }
  }

  // Write to "Stock Summary"
  let summarySheet = ss.getSheetByName("Stock Summary");
  if (!summarySheet) summarySheet = ss.insertSheet("Stock Summary");

  // Clear old summary (A4:G)
  const lastRow = summarySheet.getLastRow();
  if (lastRow > 3) {
    summarySheet.getRange(4, 1, lastRow - 3, 7).clearContent();
  }
  // Write data starting from A4
  if (summary.length > 0) {
    summarySheet.getRange(4, 1, summary.length, 7).setValues(summary);
  }
}

function getProductionSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const summarySheet = ss.getSheetByName("Stock Summary");

  // Get from & to dates from Stock Summary sheet
  const fromDate = new Date(summarySheet.getRange("J2").getValue());
  const toDate = new Date(summarySheet.getRange("L2").getValue());

  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);

  const lastRow = prodSheet.getLastRow();
  if (lastRow <= 1) return;

  // Now includes Production Date (col 12 → index 11)
  const allData = prodSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Filter rows using Production Date instead of Timestamp
  const filtered = allData
    .filter(row => {
      const prodDate = new Date(row[11]); // Production Date column
      prodDate.setHours(0, 0, 0, 0);
      const status = String(row[10]).trim().toUpperCase(); // Status column
      return prodDate >= fromDate && prodDate <= toDate && status === "IN";
    })
    .map(row => {
      // Exclude Timestamp (row[0]) → take Item Name to Code (1–5), Drum, Dept, Status, and Production Date
      return [row[11],row[1], row[2], row[3], row[4], row[5], row[8], row[9], row[10]];
    });

  // Clear old data in Stock Summary (I4:R, 9 columns now)
  summarySheet.getRange("I4:R").clearContent();

  // Paste filtered results starting from I4
  if (filtered.length > 0) {
    summarySheet.getRange(4, 9, filtered.length, 9).setValues(filtered);
  }
}

function generateGodownFGDetail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("Production/Dispatch");
  const fgSheet = ss.getSheetByName("Godown FG");
  const openingSheet = ss.getSheetByName("Opening Stock Details");

  const filterDate = fgSheet.getRange("B2").getValue();
  if (!filterDate) {
    SpreadsheetApp.getUi().alert("Please enter a date in cell B2 of Godown FG sheet.");
    return;
  }
  const filterDateStr = Utilities.formatDate(new Date(filterDate), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  let godownStock = {}; // key = ItemName|ItemCode|Dept|Category

  // --- From Opening Stock ---
  const lastCol = openingSheet.getLastColumn();
  const openHeaders = openingSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const lastRow = openingSheet.getLastRow();
  const openData = openingSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const oIdxItemName = openHeaders.indexOf("Item Name");
  const oIdxItemCode = openHeaders.indexOf("Item Code");
  const oIdxCategory = openHeaders.indexOf("Category");
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
    if (rowDateStr <= filterDateStr && partyName === "GODOWN") {
      const key = r[oIdxItemName] + "|" + r[oIdxItemCode] + "|" + r[oIdxDept] + "|" + r[oIdxCategory];
      if (!godownStock[key]) godownStock[key] = { category: r[oIdxCategory], drums: [] };
      godownStock[key].drums.push({
        drum: r[oIdxDrum],
        length: Number(r[oIdxLength]) || 0,
        color: r[oIdxColor]
      });
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
  const pIdxCategory = prodHeaders.indexOf("Category");
  const pIdxDrum = prodHeaders.indexOf("Drum");
  const pIdxLength = prodHeaders.indexOf("Production Length");
  const pIdxDept = prodHeaders.indexOf("Department");
  const pIdxColor = prodHeaders.indexOf("Code");
  const pIdxStatus = prodHeaders.indexOf("Status");
  const pIdxPartyName = prodHeaders.indexOf("Party Name");

  prodData.forEach(r => {
    if (!r[pIdxTimestamp]) return;
    const rowDateStr = Utilities.formatDate(new Date(r[pIdxTimestamp]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    if (rowDateStr > filterDateStr) return;

    const key = r[pIdxItemName] + "|" + r[pIdxItemCode] + "|" + r[pIdxDept] + "|" + r[pIdxCategory];
    if (!godownStock[key]) godownStock[key] = { category: r[pIdxCategory], drums: [] };

    const status = String(r[pIdxStatus]).trim().toUpperCase();
    const partyName = String(r[pIdxPartyName]).trim().toUpperCase();
    let lengthVal = Number(r[pIdxLength]) || 0;

    if ((status === "IN" && partyName === "GODOWN") || ((status === "OUT" || status === "SCRAP"))) {
      const drumLen = status === "IN" ? lengthVal : -lengthVal;
      godownStock[key].drums.push({
        drum: r[pIdxDrum],
        length: drumLen,
        color: r[pIdxColor]
      });
    }
  });

  // === Clear old FG content ===
  fgSheet.getRange("A4:Z1000").clearContent().clearFormat();

  // === Group by Category ===
  let categoryMap = {};
  Object.entries(godownStock).forEach(([key, val]) => {
    const category = val.category || "Uncategorized";
    if (!categoryMap[category]) categoryMap[category] = {};
    const [itemName] = key.split("|");
    if (!categoryMap[category][itemName]) categoryMap[category][itemName] = [];
    categoryMap[category][itemName].push({ key, drums: val.drums });
  });

  const maxCol = 26;
  const blockWidth = 2;
  let row = 4;
  Object.entries(categoryMap).forEach(([category, itemGroups]) => {
    fgSheet.getRange(row, 1, 1, maxCol).merge()
      .setValue(category)
      .setFontWeight("bold")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#b6d7a8")
      .setBorder(true, true, true, true, null, null);
    row++;

    let colRowTracker = {};
    let startCol = 1;

    Object.entries(itemGroups).forEach(([itemName, variations]) => {
      if (!colRowTracker[startCol]) colRowTracker[startCol] = row;
      let localRow = colRowTracker[startCol];

      variations.forEach(({ key, drums }) => {
        const [iName, iCode, dept] = key.split("|");
        const itemText = `${iName}-${iCode}(${dept})`
          .replace(/Core/gi, "C")
          .replace(/ x /gi, "X")
          .replace(/ Sq\s*mm/gi, "")
          .replace(/\s+/g, "");

        // --- Group drums by color ---
        let colorGroups = {};
        drums.forEach(d => {
          const colorKey = d.color || "";
          if (!colorGroups[colorKey]) colorGroups[colorKey] = [];
          colorGroups[colorKey].push(d);
        });

        Object.entries(colorGroups).forEach(([color, drumList]) => {
          let drumNet = {};
          drumList.forEach(d => {
            const k = d.drum + "|" + d.color;
            drumNet[k] = (drumNet[k] || 0) + (Number(d.length) || 0);
          });

          const positiveDrums = Object.entries(drumNet).filter(([_, len]) => len > 0);
          const totalLength = positiveDrums.reduce((sum, [_, len]) => sum + len, 0);
          if (totalLength === 0) return;

          // Sort positiveDrums by numeric part after 'X'
          const getNumericPart = (str) => {
            const match = str.match(/X(\d+)/i);
            return match ? parseInt(match[1], 10) : Number.MAX_SAFE_INTEGER;
          };
          positiveDrums.sort(([a], [b]) => {
            const [drumA] = a.split("|");
            const [drumB] = b.split("|");
            return getNumericPart(drumA) - getNumericPart(drumB);
          });

          // Show item with color
          fgSheet.getRange(localRow, startCol, 1, 2).merge()
            .setValue(`${itemText} [${color}]`)
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground("#d9d9d9");
          localRow++;

          fgSheet.getRange(localRow, startCol, 1, 2).setValues([["Drum", "Length"]])
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setBackground("#cfe2f3");
          localRow++;

          let total = 0;
          positiveDrums.forEach(([k, len]) => {
            const [drumName, drumColor] = k.split("|");
            fgSheet.getRange(localRow, startCol).setValue(drumName);
            const lengthCell = fgSheet.getRange(localRow, startCol + 1);
            lengthCell.setValue(len || 0);
            if (drumColor) {
              let bgColor = drumColor.trim().toLowerCase();
              const colorMap = {
                "red": "#f4cccc",
                "yellow": "#fff2cc",
                "green": "#d9ead3",
                "blue": "#c9daf8",
                "black": "#d0cece",
                "white": "#ffffff"
              };
              lengthCell.setBackground(colorMap[bgColor] || bgColor);
            } else {
              lengthCell.setBackground(null);
            }
            total += len || 0;
            localRow++;
          });

          fgSheet.getRange(localRow, startCol).setValue("Total").setFontWeight("bold").setBackground("#CAE8E7");
          fgSheet.getRange(localRow, startCol + 1).setValue(total).setFontWeight("bold").setBackground("#CAE8E7");

          fgSheet.getRange(localRow - positiveDrums.length - 2, startCol, positiveDrums.length + 3, 2)
            .setBorder(true, true, true, true, null, null);

          localRow++;
        });
      });

      fgSheet.setColumnWidth(startCol, 70);
      fgSheet.setColumnWidth(startCol + 1, 120);

      colRowTracker[startCol] = localRow;
      if (startCol + blockWidth + 1 > maxCol) {
        startCol = 1;
      } else {
        startCol += blockWidth + 1;
      }
    });
    row = Math.max(...Object.values(colRowTracker)) + 2;
  });
}



