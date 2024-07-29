
//CONSTANTS
const SPREADSHEETID = "1RODkzgSVCs4KkTENnmw5FzHxOgaWXk0UhWEJlCskqZM";
const DATARANGE = "Sheet1!A2:Q";       // A-Q
const DATASHEET = "Sheet1";
const DATASHEETID = "0";
const LASTCOL = "Q";
const IDRANGE = "Sheet1!A2:A";

/*
//Display HTML page
function doGet(e) {
  let page = e.parameter.mode || "Index";
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Replace {{NAVBAR}} with the Navbar content
  htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}", getNavbar(page)));
  return htmlOutput;
}
*/

function processForm(formObject) {
  // Find the row index or range based on the unique identifier (e.g., invoice)
  
  if (formObject.recId && checkId(formObject.recId)) {
    // If the record exists, prepare the updated values
    const values = [[
      formObject.invoice,
      formObject.customerName,
      formObject.address,
      formObject.orderDate,
      formObject.shipMode,
      formObject.productName,
      formObject.productID,
      formObject.subCategory,
      formObject.category,
      formObject.quantity,
      formObject.unitCost,
      formObject.subtotal,
      formObject.discountPercentage,
      formObject.discountAmount,
      formObject.shippingFee,
      formObject.totalAmount,
      formObject.orderId,
    ]];
    const updateRange = getRangeById(formObject.recId);
    // Update the existing record
    updateRecord(values, updateRange);
  } else {
    // If the record does not exist, prepare to add a new row
    let values = [[
      formObject.invoice,
      formObject.customerName,
      formObject.address,
      formObject.orderDate,
      formObject.shipMode,
      formObject.productName,
      formObject.productID,
      formObject.subCategory,
      formObject.category,
      formObject.quantity,
      formObject.unitCost,
      formObject.subtotal,
      formObject.discountPercentage,
      formObject.discountAmount,
      formObject.shippingFee,
      formObject.totalAmount,
      formObject.orderId,
    ]];
    // Create a new record
    createRecord(values);
  }
  // Return the last 10 records
  return getLastTenRecords();
}

/*
function createRecord(values) {
  try {
    let valueRange = Sheets.newRowData();
    valueRange.values = values;

    let appendRequest = Sheets.newAppendCellsRequest();
    appendRequest.sheetId = SPREADSHEETID;
    appendRequest.rows = valueRange;

    Sheets.Spreadsheets.Values.append(valueRange, SPREADSHEETID, DATARANGE, { valueInputOption: "RAW" });
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}
*/

function createRecord(values) {
  try {
    console.log("Creating new record. Values:", JSON.stringify(values));
    let valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.append(valueRange, SPREADSHEETID, DATARANGE, { valueInputOption: "RAW" });
    console.log("New record created successfully");
  } catch (err) {
    console.error('Failed to create record:', err);
    throw err;
  }
}


function readRecord(range) {
  try {
    let result = Sheets.Spreadsheets.Values.get(SPREADSHEETID, range);
    return result.values;
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}

function updateRecord(values, updateRange) {
  try {
    console.log("Updating record. Range:", updateRange, "Values:", JSON.stringify(values));
    let valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.update(valueRange, SPREADSHEETID, updateRange, { valueInputOption: "RAW" });
    console.log("Record updated successfully");
  } catch (err) {
    console.error('Failed to update record:', err);
    throw err;
  }
}


function deleteRecord(id) {
  const rowToDelete = getRowIndexById(id);
  const deleteRequest = {
    "deleteDimension": {
      "range": {
        "sheetId": DATASHEETID,
        "dimension": "ROWS",
        "startIndex": rowToDelete,
        "endIndex": rowToDelete + 1
      }
    }
  };
  Sheets.Spreadsheets.batchUpdate({ "requests": [deleteRequest] }, SPREADSHEETID);
  return getLastTenRecords();
}


function getLastTenRecords() {
  let lastRow = readRecord(DATARANGE).length + 1;
  let startRow = lastRow - 9;
  if (startRow < 2) { //If less than 10 records, eleminate the header row and start from second row
    startRow = 2;
  }
  let range = DATASHEET + "!A" + startRow + ":" + LASTCOL + lastRow;
  let lastTenRecords = readRecord(range);
  Logger.log(lastTenRecords);
  return lastTenRecords;
}

//GET ALL RECORDS
function getAllRecords() {
  const allRecords = readRecord(DATARANGE);
  console.log(allRecords);
  return allRecords;

}

//GET RECORD FOR THE GIVEN ID
function getRecordById(id) {
  if (!id || !checkId(id)) {
    return null;
  }
  const range = getRangeById(id);
  if (!range) {
    return null;
  }
  const result = readRecord(range);
  Logger.log(result);
  return result;
}

function testRecord(){
  getRecordById('# 13242')
}

function getRowIndexById(id) {
  if (!id) {
    throw new Error('Invalid ID');
  }

  const idList = readRecord(IDRANGE);
  for (var i = 0; i < idList.length; i++) {
    if (id == idList[i][0]) {
      var rowIndex = parseInt(i + 1);
      console.log(rowIndex);
      return rowIndex;
    }
  }
}

function testRowIndex(){
  getRowIndexById('# 13242')
}


//VALIDATE ID
function checkId(id) {
  const idList = readRecord(IDRANGE).flat();
  console.log(idList.includes(id));
  return idList.includes(id);
}


function testId(){
  checkId('# 15115');
}


//GET DATA RANGE IN A1 NOTATION FOR GIVEN ID
function getRangeById(id) {
  if (!id) {
    console.log("Not found")
    return null;
  }
  const idList = readRecord(IDRANGE);
  const rowIndex = idList.findIndex(item => item[0] === id);
  console.log(rowIndex);
  if (rowIndex === -1) {
    console.log("not found")
    return null;
  }
  const range = `Sheet1!A${rowIndex + 2}:${LASTCOL}${rowIndex + 2}`;
  console.log(range);
  return range;
}

function testRange(){
  getRangeById('# 15115');

}


//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}


//SEARCH RECORDS
function searchRecords(formObject) {
  let result = [];
  try {
    if (formObject.searchText) {//Execute if form passes search text
      const data = readRecord(DATARANGE);
      const searchText = formObject.searchText;

      // Loop through each row and column to search for matches
      for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
          const cellValue = data[i][j];
          if (cellValue.toLowerCase().includes(searchText.toLowerCase())) {
            result.push(data[i]);
            break; // Stop searching for other matches in this row
          }
        }
      }
    }
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
  return result;
}