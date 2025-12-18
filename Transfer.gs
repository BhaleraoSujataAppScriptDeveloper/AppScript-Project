function saveTransferEntry(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName("In-Out");
  const historySheet = ss.getSheetByName("Transfer History");
  const prodData = prodSheet.getDataRange().getValues();
  // Column indexes based on header
  const col = {
    timeStamp: 0,
    itemName: 1,
    itemCode: 2,
    category: 3,
    drum: 4,
    productionLength: 5,
    partyName: 6,
    partyAssignLength: 7,
    department: 8,
    code: 9,
    status: 10,
    productionDate: 11
  };

  let foundRow = -1;
  let availableLength = 0;

  // Find the row in Production/Dispatch that matches item, code, category, drum, and fromParty
  for (let i = 1; i < prodData.length; i++) {
    if (
      prodData[i][col.itemName] === data.itemName &&
      prodData[i][col.itemCode] === data.itemCode &&
      prodData[i][col.category] === data.category &&
      prodData[i][col.drum] === data.drum &&
      prodData[i][col.partyName] === data.fromParty
    ) {
      foundRow = i + 1; // Spreadsheet row index
      availableLength = prodData[i][col.partyAssignLength];
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error("No matching row found in Production/Dispatch for selected item/drum/party.");
  }

  // Update Production/Dispatch sheet
  if (data.transferLength === availableLength) {
    // Entire length → update the row
    prodSheet.getRange(foundRow, col.partyName + 1).setValue(data.toParty);
    prodSheet.getRange(foundRow, col.partyAssignLength + 1).setValue(availableLength);
  } else {
    // Partial length → split row
    // 1. Update original row with remaining length
    prodSheet.getRange(foundRow, col.partyAssignLength + 1).setValue(availableLength - data.transferLength);

    // 2. Append new row for transferred part
    const newRow = [
      new Date(),              // Timestamp
      data.itemName,
      data.itemCode,
      data.category,
      data.drum,
      data.transferLength,     // Production Length for transferred part
      data.toParty,            // Party Name
      data.transferLength,     // Party assign Length
      prodData[foundRow - 1][col.department],
      prodData[foundRow - 1][col.code],
      "IN",                    // Status
      data.transferDate
    ];
    prodSheet.appendRow(newRow);
  }

  // Save in Transfer History (always append)
  const historyRow = [
    new Date(),
    data.itemName,
    data.itemCode,
    data.category,
    data.drum,
    data.transferLength,
    data.fromParty,
    data.toParty,
    data.transferDate
  ];
  historySheet.appendRow(historyRow);

  return "Transfer saved successfully!";
}
