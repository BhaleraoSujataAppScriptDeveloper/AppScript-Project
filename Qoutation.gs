//function doGet() {
 // return HtmlService.createHtmlOutputFromFile("Quotation Form");
//}
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Quotation Form');
  template.inqNo = e.parameter.inqNo || "";  // Pass to HTML
  return template.evaluate();
}

// Get all unique Descriptions (Headers)
function getHeaderDescriptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const data = sheet.getRange("B2:B" + sheet.getLastRow()).getValues().flat();
  return [...new Set(data.filter(Boolean))];
}

// Get all Codes (Types) for a given Description
Logger.log(getTypesByDescription("1.1 KV LT XLPE Aluminium Armoured Cable (As Per IS :7098)"))

function getTypesByDescription(description) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Master");
  const data = sheet.getDataRange().getValues();

  const cleanedDesc = (description || "").toLowerCase().replace(/\s+/g, " ").trim();
  const types = new Set();

  for (let i = 1; i < data.length; i++) {
    const rowDesc = (data[i][1] || "").toLowerCase().replace(/\s+/g, " ").trim(); // col B = desc
    if (rowDesc === cleanedDesc) {
      types.add(data[i][0]); // col A = type
    }
  }

  return Array.from(types);
}

function getAllItems() {
  const sheet = SpreadsheetApp.openById('1ZCjxuaee-VOTYVIv83iACqV2XO3ZKLtbQup_m2MJpJw').getSheetByName('Master');
  const data = sheet.getRange('A2:A').getValues().flat().filter(String);
  return [...new Set(data)]; // unique list
}

function getItemSuggestions(query) {
  const sheet = SpreadsheetApp.openById('1ZCjxuaee-VOTYVIv83iACqV2XO3ZKLtbQup_m2MJpJw').getSheetByName('Master');
  const data = sheet.getRange('A2:A').getValues().flat().filter(String); // Descriptions in column A

  // Convert to lowercase and filter only those starting with the query
  const filtered = data.filter(item => item.toLowerCase().startsWith(query.toLowerCase()));

  return [...new Set(filtered)]; // unique filtered list
}
function getStates(query) {
  const sheet = SpreadsheetApp.openById("1CznHf7LLuOelhwO1lk9bvT6W1AnqEU447znD_15wqpQ").getSheetByName("States-Cities");
  const states = sheet.getRange("A2:A").getValues().flat().filter(Boolean);
  query = query.toLowerCase();
  return [...new Set(states)].filter(s => s.toLowerCase().startsWith(query)).slice(0, 20);
}

function getCitiesByState(cityQuery, stateName) {
  const sheet = SpreadsheetApp.openById("1CznHf7LLuOelhwO1lk9bvT6W1AnqEU447znD_15wqpQ").getSheetByName("States-Cities");
  const data = sheet.getRange("A2:B").getValues(); // A: State, B: City
  cityQuery = cityQuery.toLowerCase();
  stateName = stateName.toLowerCase();

  const filteredCities = data
    .filter(([state, city]) => state && city && state.toLowerCase() === stateName)
    .map(([_, city]) => city)
    .filter(city => city.toLowerCase().startsWith(cityQuery));

  return [...new Set(filteredCities)].slice(0, 20);
}

function submitOrUpdateQuotation(data) {
  var docURL = "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Quotation Format");

  const rows = [];

  data.headers.forEach(header => {
    const cleanHeader = header.headerLabel
      ? header.headerLabel.replace(/<[^>]*>/g, '').replace(/[➕✖]/g, '').trim()
      : '';

    let totalQty = 0;
    let totalUnitPrice = 0;
    let totalBasicValue = 0;

    header.items.forEach(item => {
      totalQty += item.qty;
      totalUnitPrice += item.unitPrice;
      totalBasicValue += item.basicValue;
    });

    header.items.forEach((item, index) => {
      rows.push([
        data.quotationDate,
        data.quotationNo,
        data.quotationType,
        data.cableType,
        data.companyName,
        data.address,
        data.state,
        data.city,
        data.contact,
        data.email,
        (cleanHeader + " " + (header.description || "")).trim(),
        index + 1,
        item.item,
        item.type,
        item.uom,
        item.qty,
        item.unitPrice,
        item.basicValue,
        totalQty,
        totalUnitPrice,
        totalBasicValue,
        data.salesExecutive,
        data.status || "Draft",
        data.kname,
        data.kmobile,
        data.project,
        data.quotTitle,
        data.finalQuotationNumber
      ]);
    });
  });

  // 🔹 Delete old draft if exists (same Quotation No + Draft)
  if (data.status  === "Draft") {
    const lastRow = mainSheet.getLastRow();
    if (lastRow > 1) {
      const values = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn()).getValues();
      const rowsToDelete = [];
      values.forEach((row, i) => {
        const qNo = row[1];   // column B = Quotation No
        const status = row[22]; // column V = Quotation Status (index 21 → 22nd col)
        if (String(qNo).trim() === String(data.quotationNo).trim() && String(status).toLowerCase() === "draft") {
          rowsToDelete.push(i + 2); // +2 bcoz header row + 1-based index
        }
      });

      // delete from bottom to top (to avoid shifting issue)
      rowsToDelete.reverse().forEach(r => mainSheet.deleteRow(r));
    }
  }

  // 🔹 Insert all rows at once
  if (rows.length > 0) {
    mainSheet.getRange(mainSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    docURL = generateQuotationSheetFromTemplate_Excel(data);
  }

  // 🔹 Save terms to QuotationTerms sheet
  let termsSheet = ss.getSheetByName("QuotationTerms");
  if (!termsSheet) {
    termsSheet = ss.insertSheet("QuotationTerms");
    termsSheet.appendRow(["Quotation No", "Quotation Type", "Terms & Conditions"]);
  }

  termsSheet.appendRow([
    data.quotationNo,
    data.quotationType,
    data.termsText || ""
  ]);
  return docURL;
}



function stripHtml(html) {
  return html
    .replace(/\r\n/g, '\n')                     // Normalize line endings
    .replace(/<br\s*\/?>/gi, '\n')              // <br> -> newline
    .replace(/<\/(div|p)>/gi, '\n')             // </div>/<p> -> newline
    .replace(/<[^>]+>/g, '')                    // Remove HTML tags
    .replace(/&nbsp;/gi, ' ')                   // &nbsp; -> space
    .replace(/\u00A0/g, ' ')                    // Non-breaking spaces
    .replace(/(?<!\n)([0-9]+\.\s)/g, '\n$1')    // Ensure each number starts on new line
    .split('\n')                                // ✅ Split into lines
    .map(line => line.trim())                   // ✅ Trim spaces on each line
    .filter(line => line.length > 0)            // ✅ Remove empty lines completely
    .join('\n');                                // ✅ Join back
}
function getNextFinalNumber(type) {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 22).getValues() : [];

  const prefix = "VCPL/QTN/";
  const fy = getFinancialYear();
  const isTender = type === "Tender";

  // Collect all draft numbers of this type
  const numbers = data
    .filter(row => row[21] === "Final" && row[0]) // row[0] = Quotation Number, row[21] = Status
    .map(row => row[0])
    .filter(q => isTender ? q.includes("/T-") : !q.includes("/T-"))
    .map(q => {
      const parts = q.split("/");
      return parseInt(parts[2].replace("T-", ""), 10);
    })
    .filter(n => !isNaN(n));

  const nextNum = numbers.length ? Math.max(...numbers) + 1 : 1; // start from 1
  const numberStr = String(nextNum).padStart(4, "0");

  return isTender
    ? `${prefix}T-${numberStr}/${fy}`
    : `${prefix}${numberStr}/${fy}`;
}

function generateQuotationNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  
  let maxNum = 0;

  for (let i = 0; i < values.length; i++) {
    const val = values[i][0];
    let isNumeric = false;
    let numValue = null;

    // Check if it's a number
    if (typeof val === "number" && !isNaN(val)) {
      isNumeric = true;
      numValue = val;
    }
    // Check if it's a string of only digits
    else if (typeof val === "string" && /^[0-9]+$/.test(val.trim())) {
      isNumeric = true;
      numValue = parseInt(val.trim(), 10);
    }

    // Log every row for debugging
    Logger.log(`Row ${i+2}: value="${val}", isNumeric=${isNumeric}, numValue=${numValue}`);

    // Update max
    if (isNumeric && numValue > maxNum) {
      maxNum = numValue;
    }
  }

  const nextNum = maxNum + 1;
  Logger.log("Next Quotation Number: " + nextNum);
  return nextNum.toString();
}


function getFinancialYear() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const fyStart = month >= 4 ? year : year - 1;
  const fyEnd = fyStart + 1;
  return `${fyStart}-${String(fyEnd).slice(2)}`;
}


function getMasterNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  const data = sheet.getRange("E2:E" + sheet.getLastRow()).getValues();
  const names = data.flat().filter(name => name); // remove empty
  return [...new Set(names)]; // return unique names
}
function onOpen() {

}

function getQuotationTerms(quotationNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QuotationTerms");
  if (!sheet) return "";

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues(); // [Quotation No, Type, Terms]

  for (const [qNo, , terms] of data) {
    if (qNo === quotationNo) return terms;
  }

  return ""; // Not found
}

function getDraftQuotationNumbers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const result = [];
  const seen = new Set();

  const quotationNoCol = headers.indexOf("Quotation Number");
  const companyNameCol = headers.indexOf("Company Name");
  const statusCol = headers.indexOf("Quotation Status");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusCol];
    if (status && status.toString().toLowerCase() === "draft") {
      const qNo = row[quotationNoCol];
      const company = row[companyNameCol];
      const key = `${qNo} - ${company}`;
      if (qNo && company && !seen.has(key)) {
        seen.add(key);
        result.push({ value: qNo, label: key });
      }
    }
  }
  return result;
}
function getFinalQuotations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return []; // no data

  const headers = data[0];
  const quotationNoCol = headers.indexOf("Final Quotation Number");
  const companyNameCol = headers.indexOf("Company Name");
  const statusCol = headers.indexOf("Quotation Status");

  if (quotationNoCol === -1 || companyNameCol === -1 || statusCol === -1) {
    throw new Error("One or more column headers not found");
  }

  const result = [];
  const seen = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusCol] ? row[statusCol].toString().trim().toLowerCase() : "";
    if (status === "final") {
      const qNo = row[quotationNoCol] ? row[quotationNoCol].toString().trim() : "";
      const company = row[companyNameCol] ? row[companyNameCol].toString().trim() : "";
      const key = `${qNo} - ${company}`;
      if (qNo && company && !seen.has(key)) {
        seen.add(key);
        result.push({ value: qNo, label: key });
      }
    }
  }

  Logger.log("Final Quotations: " + JSON.stringify(result));
  return result;
}
function loadQuotationByNumber(qNo,searchType) {
  let matchedRows="";
  const sheet = SpreadsheetApp.getActive().getSheetByName("Quotation Format");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 28).getValues(); // A2:W
if(searchType==="EDIT")
{
   matchedRows = data.filter(row => String(row[1]).trim() === String(qNo).trim()&&String(row[22]).trim().toLowerCase()==="draft" );
}
if(searchType==="REVISED")
{
   matchedRows = data.filter(row => String(row[27]).trim() === String(qNo).trim()&&String(row[22]).trim().toLowerCase()==="final");
}

  if (!matchedRows.length) return null;

  const firstRow = matchedRows[0];

  const quotation = {
   // quotationDate:  Utilities.formatDate(new Date(firstRow[0]), Session.getScriptTimeZone(), "dd-MM-yyyy"), // A
    quotationDate: Utilities.formatDate(new Date(firstRow[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"),

    quotationNo: firstRow[1],          // B
    quotationType: firstRow[2],        // C
    cableType: firstRow[3],            // D
    companyName: firstRow[4],          // E
    address: firstRow[5],              // F
    state: firstRow[6],                // G
    city: firstRow[7],                 // H
    contact: firstRow[8],              // I
    email: firstRow[9],                // J
    salesExecutive: firstRow[21], 
    project:firstRow[25] ,    // Z
    kname:firstRow[23],
    kmobile:firstRow[24],
    quotTitle:firstRow[26],
    finalQuotationNumber:firstRow[27],
    termsText: "",                     // placeholder (we’ll load this later if needed)
    headers: []
  };

  const headerMap = {};

  matchedRows.forEach(row => {
    const headerLabel = row[10]; // K
    if (!headerLabel) return;

    if (!headerMap[headerLabel]) {
      headerMap[headerLabel] = [];
    }

    headerMap[headerLabel].push({
      item: row[12],       // M
      type: row[13],       // N
      uom: row[14],        // O
      qty: row[15],        // P
      unitPrice: row[16],  // Q
      basicValue: row[17]  // R
    });
  });

  for (const [headerLabel, items] of Object.entries(headerMap)) {
    quotation.headers.push({ headerLabel, items });
  }
  const termsSheet = SpreadsheetApp.getActive().getSheetByName("QuotationTerms");
  if (termsSheet) {
    const termsData = termsSheet.getRange(2, 1, termsSheet.getLastRow() - 1, 3).getValues(); // A2:C
    for (let i = 0; i < termsData.length; i++) {
      if (String(termsData[i][0]).trim() === String(qNo).trim()) {
        quotation.termsText = termsData[i][2] || ""; // Column C: Terms
        break;
      }
    }
  }
  return quotation;
}
function test1236(){
Logger.log(getSalesExecutiveNumber("Manmohansingh.D.A "))
}

function getSalesExecutiveNumber(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Master");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const execName = (data[i][4] || "").toString().trim(); // column E (index 4)
    if (execName.toLowerCase() === name.toLowerCase().trim()) {
      return (data[i][5] || "").toString().trim(); // column F (index 5)
    }
  }
  return "";
}
function generateQuotationDocFromTemplate(data) {
  const templateId = '1trpcBN-Ohe-WW-9skaamccKxPy1HuLAyO_G_UUKZtTk'; // Google Docs template ID
  const outputFolderId = '1eL99HAdUyNHKJBAX9P5UhX5ekM2MDLOq'; // Output folder ID
  const templateFile = DriveApp.getFileById(templateId);
  const outputFolder = DriveApp.getFolderById(outputFolderId);
 const tempCopy = templateFile.makeCopy(`Quotation_${data.quotationNo}`, outputFolder)
  const doc = DocumentApp.openById(tempCopy.getId());
  const body = doc.getBody();

  const safeReplace = (key, val) => {
body.replaceText(`\\[${key}\\]`, val || '');
  };
  // === Replace placeholders ===
  safeReplace('QTN_NO', data.quotationNo);
  safeReplace('DATE', data.quotationDate);
  safeReplace('COMPANY_NAME', data.companyName);
  safeReplace('COMPANY_ADD', data.address);
  safeReplace('STATE', data.state);
  safeReplace('CITY', data.city);
  safeReplace('CONTACT', data.contact);
  safeReplace('EMAIL', data.email);
  safeReplace('QUOT_TYPE', data.quotationType);
  const cleanTerms = (data.termsText || "").replace(/<br\s*\/?>/gi, "");
  safeReplace('TERMS', cleanTerms);
  const salesExecName = data.salesExecutive || "";
  const salesExecNumber = getSalesExecutiveNumber(salesExecName);
  safeReplace('SALES_EXC', salesExecName);
  safeReplace('SALES_NUMBER', salesExecNumber);
  if(data.project)
  {
    safeReplace('PROJECT', "For " + data.project + " Project");
  }
  safeReplace('KNAME', data.kname);
  safeReplace('KMOBILE', data.kmobile);
  safeReplace('CABLE_TYPE', data.cableType);

  // === Insert item table ===
  const found = body.findText('\\[ITEM_TABLE\\]');
  if (!found) return doc.getUrl();
  const element = found.getElement();
  const parent = element.getParent();
  let insertIndex = body.getChildIndex(parent);
  parent.asParagraph().clear();

  let totalQty = 0, totalValue = 0, totalUOMPrice=0;
  let table;


  data.headers.forEach(header => {
    const cleanHeader = header.headerLabel.replace(/\s*X$/, "").trim();
    const headerText = cleanHeader + (header.description ? ' ' + header.description : '');

    // Section header
    const para = body.insertParagraph(++insertIndex, headerText.trim());
    para.setBold(true)
        .setFontSize(11)
        .setSpacingBefore(0)
        .setSpacingAfter(0)
        .setLineSpacing(1.0)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .editAsText().setBackgroundColor(null);
    // Table below header
    table = body.insertTable(++insertIndex);
    // Table headers
    const headerRow = table.appendTableRow();
    ['Sr.No.', 'Description', 'Type', 'UOM', 'Qty', 'Unit Price', 'Basic Value'].forEach(text => {
      const cell = headerRow.appendTableCell(text);
      cell.setBackgroundColor('#cfe2f3');
      cell.editAsText().setBold(true).setFontSize(10);
    });

    // Table items
    header.items.forEach((item, i) => {
      const row = table.appendTableRow();
      [
        String(i + 1),
        item.item,
        item.type,
        item.uom,
        Number(item.qty).toFixed(2),
        Number(item.unitPrice).toFixed(2),
        Number(item.basicValue).toFixed(2)
      ].forEach(val => {
        const cell = row.appendTableCell(val);
        cell.setBackgroundColor("#ffffff");
        cell.editAsText().setFontSize(10);
      });
      totalUOMPrice +=Number(item.unitPrice);
      totalQty += Number(item.qty);
      totalValue += Number(item.basicValue);
    });
    setColumnWidths(table);
  });

  // === Total row ===
  if (table) {
    const totalRow = table.appendTableRow();
    const totalLabel = totalRow.appendTableCell("TOTAL");
    totalLabel.editAsText().setBold(true).setFontSize(10);

    for (let i = 1; i < 4; i++) {
      const emptyCell = totalRow.appendTableCell('');
    }

    const qtyCell = totalRow.appendTableCell(totalQty.toFixed(2));
    qtyCell.editAsText().setBold(true).setFontSize(10);

    const uomPriceCell = totalRow.appendTableCell(totalUOMPrice.toFixed(2));
    uomPriceCell.editAsText().setBold(true).setFontSize(10);

    const valueCell = totalRow.appendTableCell(totalValue.toFixed(2));
    valueCell.editAsText().setBold(true).setFontSize(10);
  }
  doc.saveAndClose();
  // ✅ Return the Google Doc URL for editing
  return doc.getUrl();
}
function cleanTermsAndConditions(rawTerms) {
  return rawTerms
    .replace(/<br\s*\/?>/gi, '\n')     // Convert <br> or <br/> to line breaks
    .replace(/^\s+/gm, '')             // Remove leading spaces from each line
    .replace(/&amp;/g, '&')            // Decode common HTML entities
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}
function generateQuotationSheetFromTemplate_Excel(data) {
  const templateFileId = '1tFYQ19VgGhJTKzuiXVcHcUc1sihlHHHIpt12gJBAwAU'; // Template Spreadsheet ID
  const parentFolderId = '1eL99HAdUyNHKJBAX9P5UhX5ekM2MDLOq'; // Folder where new files will go

  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const templateFile = DriveApp.getFileById(templateFileId);

  // 🔹 Create a fresh copy of the template as a new standalone file
  const newFile = templateFile.makeCopy(`${data.companyName}_${data.quotationNo}`, parentFolder);
  const newSS = SpreadsheetApp.openById(newFile.getId());
  const sheet = newSS.getSheets()[0]; // First sheet in the copied file

  // 🔹 Replace placeholders
  replaceSheetPlaceholders(sheet, data);

  // 🔹 Insert item table
  insertItemsAtPlaceholder(sheet, "[ITEM_TABLE]", data.headers);

  // ✅ Return the new file URL (not just a tab inside master)
  return newSS.getUrl();
}

const formatDateStructure = (dateStr) => {
  const date = new Date(dateStr);
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
};

function formatQuotationNumber(number) {
  // Determine current financial year
  const today = new Date();
  let fyStartYear = today.getFullYear();
  let fyEndYear = fyStartYear + 1;

  // If month is Jan, Feb, Mar => belong to previous FY
  if (today.getMonth() < 3) {
    fyStartYear -= 1;
    fyEndYear -= 1;
  }
  const financialYear = fyStartYear + "-" + fyEndYear.toString().slice(-2);
  const qNumber = `VCPL/QTN/${number}/${financialYear}`;
  Logger.log(qNumber);
  return qNumber;
}

// Example usage:


function replaceSheetPlaceholders(sheet, data) {
  let FinalQNumber="";
  Logger.log("Replace Placeholder");
  const range = sheet.getDataRange();
  const values = range.getValues();
  let ContactValue_ = data.contact ? "Contact - " + data.contact : "";
    if(data.status==="Final" || data.status==="REVISED")
    {
      FinalQNumber= formatQuotationNumber(data.finalQuotationNumber)
      data.quotationNo=FinalQNumber;
    }
    else
    {
      data.quotationNo= formatQuotationNumber(data.quotationNo)
    }
  const replacements = {
    '[Quotation_Number]': data.quotationNo,
    '[DATE]':formatDateStructure(data.quotationDate),
    '[COMPANY_NAME]': data.companyName.replace(/,+$/, "").trim(),
    '[COMPANY_ADD]': data.address,
    '[STATE]': data.state,
    '[CITY]': data.city,
    '[CONTACT]': ContactValue_,
    '[EMAIL]': data.email,
    '[QUOT_TYPE]': data.quotationType,
    '[CABLE_TYPE]': data.cableType,
    '[PROJECT]': "For " +data.project + " Project" || '',
    '[KNAME]': data.kname,
    '[KMOBILE]': data.kmobile,
    '[TERMS]': stripHtml(data.termsText || ""),
    '[SALES_EXC]': data.salesExecutive,
    '[SALES_NUMBER]': getSalesExecutiveNumber(data.salesExecutive),
    '[QUOTTITLE]': data.quotTitle || ''
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

function insertItemsAtPlaceholder(sheet, marker, headers) {
  Logger.log("Inserting items at placeholder...");

  const markerCell = sheet.createTextFinder(marker).findNext();
  if (!markerCell) {
    Logger.log("Marker not found: " + marker);
    return;
  }
  const startRow = markerCell.getRow();
  const startCol = markerCell.getColumn();
  markerCell.clearContent(); // Remove the placeholder

  const columnWidths = [40, 250, 80, 80, 90, 90, 100];
  columnWidths.forEach((width, i) => {
  sheet.setColumnWidth(startCol + i, width);
});

  // We'll collect all item rows and calculate totals after loop
  let allItemRows = [];
  let srNo = 1;

  // First, calculate how many rows we need (for insertion)
  let totalNewRows = 0;
  headers.forEach(header => {
    if (!header.items || header.items.length === 0) return;
    totalNewRows += 1  + header.items.length; // header + col headers + items
  });

  totalNewRows += 1; // for total row

  if (totalNewRows > 0) {
    sheet.insertRowsAfter(startRow, totalNewRows);
  }

  let currentRow = startRow;
  const fullTableStart = currentRow;
    // --- Column Headers ---
    const colHeaders = [['Sr.No.', 'Description', 'Type', 'UOM', 'Qty', 'Unit Price', 'Basic Value']];
    const colRange = sheet.getRange(currentRow, startCol, 1, 7);
    colRange.setValues(colHeaders)
            .setFontWeight("bold")
            .setBackground("#f2f2f2")
            .setHorizontalAlignment("center");
    currentRow++;

  headers.forEach((header,i) => {
    const items = header.items;
    if (!items || items.length === 0) return;
    let srNo = 1;
    // --- Header Label ---
    const alphaLabel = String.fromCharCode(65 + i); // 65 = 'A'
    const fullHeader = header.headerLabel + (header.description ? ' - ' + header.description : '');
    sheet.getRange(currentRow, startCol).setValue(alphaLabel + '.')
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

    const headerRange = sheet.getRange(currentRow, startCol + 1, 1, 6);
    headerRange.merge();
    headerRange.setValue(fullHeader.replace(/\s*X$/, "").replace(/<[^>]*>/g, '').replace(/[➕✖]/g, '').trim())
               .setFontWeight("bold")
               .setFontSize(11)
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle");
    currentRow++;
    // --- Items ---
    const itemRows = items.map(item => [
      srNo++,
      item.item,
      item.type,
      item.uom,
      Number(item.qty),
      (Number(item.unitPrice) === 0) ? "" : Number(item.unitPrice), 
      Number(item.basicValue)
    ]);

    allItemRows.push(...itemRows);

    const itemRange = sheet.getRange(currentRow, startCol, itemRows.length, 7);
    itemRange.setValues(itemRows)
             .setHorizontalAlignment("center");
             sheet.getRange(currentRow, startCol + 6, itemRows.length, 1)
     .setHorizontalAlignment("right");
    currentRow += itemRows.length;
  });

  // --- Total Row (One Time Only) ---
  const totalQty = allItemRows.reduce((sum, row) => sum + Number(row[4]), 0);
  const totalUnitPrice = allItemRows.reduce((sum, row) => sum + Number(row[5]), 0);
  const totalBasic = allItemRows.reduce((sum, row) => sum + Number(row[6]), 0);

// Merged cell "Total" (B to F)
sheet.getRange(currentRow, 3, 1, 4)
     .merge()
     .setValue("Total")
     .setFontWeight("bold")
     .setHorizontalAlignment("center")
     .setBorder(true, true, true, true, true, true);  // <-- Border here

// totalQty in G
sheet.getRange(currentRow, 7)
     .setValue(totalQty)
     .setFontWeight("bold")
     .setHorizontalAlignment("center")
     .setBorder(true, true, true, true, true, true);

// totalBasic in I
sheet.getRange(currentRow, 9)
     .setValue(totalBasic)
     .setFontWeight("bold")
     .setHorizontalAlignment("right")
     .setBorder(true, true, true, true, true, true);
// Clear column H if needed
sheet.getRange(currentRow, 8).clearContent();

  const fullTableEnd = currentRow;
  const fullRange = sheet.getRange(fullTableStart, startCol, fullTableEnd - fullTableStart + 1, 7);
  fullRange.setBorder(true, true, true, true, true, true);
}

function setColumnWidths(table) {
  const widths = [40, 150, 50, 50, 60, 60, 70]; // Adjust widths as needed
  const numRows = table.getNumRows();
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    for (let c = 0; c < widths.length; c++) {
      row.getCell(c).setWidth(widths[c]);
    }
  }
}

function exportSheetAsExcel(sheetId, fileName, folderId) {
  Logger.log("Export");
  const url = `https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
  const token = ScriptApp.getOAuthToken();
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  const blob = response.getBlob().setName(fileName + '.xlsx');
  const folder = DriveApp.getFolderById(folderId);
  const excelFile = folder.createFile(blob);
  return excelFile.getUrl();
}


function testQuotationExcel() {
  const data = {
  quotationNo: "Q123",
  quotationDate: "2025-07-22",
  companyName: "ABC Ltd.",
  address: "123 Street, Industrial Area",
  state: "Maharashtra",
  city: "Mumbai",
  contact: "9876543210",
  email: "abc@company.com",
  quotationType: "Electrical",
  salesExecutive: "John",
  project: "Metro Line",
  kname: "Mr. Kumar",
  kmobile: "9876543210",
  cableType: "Copper",
  terms: "Payment within 30 days",
  headers: [
    {
      headerLabel: "Main Cable",
      description: "Underground",
      items: [
        { item: "3.5 Core 300 Sqmm", type: "Power", uom: "MTR", qty: 100, unitPrice: 200, basicValue: 20000 }
      ]
    }
  ]
};
  generateQuotationSheetFromTemplate_Excel(data);
}

function getDistinctParties() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const data = sheet.getRange("E2:E" + sheet.getLastRow()).getValues(); // Column E
  const uniqueParties = [...new Set(data.flat().filter(name => name))]; // Remove duplicates and blanks
  return uniqueParties.sort(); // Optional: sort alphabetically
}
function getPartyDetails(partyName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation Format");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; // no data

  // Get header row
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Normalize headers for matching (lowercase trimmed)
  const norm = headers.map(h => String(h).trim().toLowerCase());

  // Helper to find column index by header name (case-insensitive)
  const col = name => norm.indexOf(name.trim().toLowerCase());

  // ***** ADJUST THESE TO MATCH YOUR SHEET HEADER TEXT EXACTLY *****
  const idxCompanyName   = col("Company Name");       // Party select uses this
  const idxAddress       = col("Address");
  const idxState         = col("State");
  const idxCity          = col("City");
  const idxMobile        = col("Mobile");             // Contact phone
  const idxEmail         = col("Email");
  const idxProject       = col("Project");
  const idxSalesExec     = col("Sales Executive");
  const idxKindAttName   = col("Kind Att Name");
  const idxKindAttMobile = col("Kind att mobile");
  const idxQuotationType = col("Quotation Type");     // include if exists
  const idxCableType     = col("Cable Type");         // include if exists

  if (idxCompanyName === -1) return {};
  // Get all data rows
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  // Find the row that matches the selected party
  const row = data.find(r => (r[idxCompanyName] || "").toString().trim() === partyName.trim());
  if (!row) return {};

  // Build result with camelCase keys expected by your client JS
  return {
    companyName:    row[idxCompanyName]   || "",
    address:        row[idxAddress]       || "",
    state:          row[idxState]         || "",
    city:           row[idxCity]          || "",
    contact:        row[idxMobile]        || "",
    email:          row[idxEmail]         || "",
    project:        row[idxProject]       || "",
    salesExecutive: row[idxSalesExec]     || "",
    kname:          row[idxKindAttName]   || "",
    kmobile:        row[idxKindAttMobile] || "",
    quotationType:  idxQuotationType > -1 ? row[idxQuotationType] : "",
    cableType:      idxCableType > -1 ? row[idxCableType] : ""
  };
}

function deleteQuotationFromSheet(qNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Quotation Format");
  const data = sheet.getDataRange().getValues();

  let qNoCol = 1; // Adjust index: 0 = Column A, 1 = Column B
  let rowsToDelete = [];

  // Find all rows with matching quotation number
  for (let i = data.length - 1; i >= 1; i--) { // Start from bottom, skip header
    if (data[i][qNoCol] == qNo) {
      sheet.deleteRow(i + 1);
    }
  }

  return "Quotation No: " + qNo + " deleted successfully.";
}




