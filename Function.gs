const FOLDER_NAME_PDFS = "Inv_pdf"; // The source folder name of your PDF files

// 1. PDF -> Google Doc
/**
 * Convert a PDF file into a Google Doc
 * @param {string} id The id of the PDF
 * @returns {DocumentApp.Document} The Google Doc object
 */
function convertPdfToDoc(id, ocrLanguage="en"){
  const pdf = DriveApp.getFileById(id);
  const resource = {
    title: pdf.getName(),
    mimeType: pdf.getMimeType()
  };
  const mediaData = pdf.getBlob();
  const options = {
    convert: true,
    ocr: true,
    ocrLanguage
  };
  const newFile = Drive.Files.insert(resource, mediaData, options);
  return DocumentApp.openById(newFile.id);
}

// 2. Read text from Google Doc
/**
 * Read text content from the PDF file
 * @param {string} id The id of the PDF file
 * @param {boolean} trashDocFile Delete or keep the Google Doc
 * @returns {string} The text content on the PDF file
 */
function readTextFromPdf(id, trashDocFile=true){
  const doc = convertPdfToDoc(id);
  const text = doc.getBody().getText();
  DriveApp.getFileById(doc.getId()).setTrashed(trashDocFile);
  return text;
}

// 3. Get PDF Files
function getPdfFiles(folderName=FOLDER_NAME_PDFS){
  const ss = SpreadsheetApp.getActive();
  const currentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const folders = currentFolder.getFoldersByName(folderName);
  if (!folders.hasNext()) return [];
  const folder = folders.next();
  const files = folder.getFilesByType(MimeType.PDF);
  const ids = [];
  while(files.hasNext()){
    ids.push(files.next().getId());
  }
  return ids;
}

// 4. Parse Invoice Data
function parseInvoiceData(itemLines) {
  const regex = /(.+?) (\d+) \$([\d,.]+) \$([\d,.]+) (.+?), (.+?), (.+)$/;
  const items = [];
  
  itemLines.forEach(itemLine => {
    const match = itemLine.match(regex);
    if (match) {
      const productName = match[1].trim();
      const quantity = parseInt(match[2], 10);
      const unitPrice = `$${match[3]}`;
      const subtotal = parseFloat(match[4].replace(/[^0-9.-]+/g, ""));
      const subCategory = match[5].trim();
      const category = match[6].trim();
      const productId = match[7].trim();
      items.push({
        productName,
        quantity,
        unitPrice,
        subtotal: `$${subtotal.toFixed(2)}`,
        subCategory,
        category,
        productId
      });
    }
  });
  return items;
}

// 5. Get Invoice Data From Text
function getInvoiceDataFromText(text) {
  const lines = text.split("\n");
  console.log(lines);
  
  let customerName = "", address = "", orderDate = "", shipMode = "", invoice = "", discountPercent = "", discountAmount = "", shippingFee = "", totalAmount = "", orderId = "";
  const items = [];
  
  lines.forEach((line, index) => {
    if (line.startsWith("Discount (")) {
      discountPercent = line.match(/\d+/)[0] + '%';
      const nextLine = lines[index + 7].trim();
      discountAmount = nextLine.split(" ")[2].trim();
    }
    if (line.startsWith("Bill To: ")) {
      customerName = lines[index + 1].trim();
      address = lines[index + 3].trim();
      if (!lines[index + 4].trim().startsWith("Date:")) {
        address += lines[index + 4].trim();
      }
    }
    if (line.startsWith("Date: ")) {
//      orderDate = formatDateToMMDDYYYY(lines[index + 4].trim());
        orderDate = lines[index + 4].trim();
    }
    if (line.startsWith("Ship Mode: ")) {
      shipMode = lines[index + 4];
    }
    if (line.startsWith("INVOICE #")) {
      invoice = line.split(" ")[2];
    }
    if (line.includes("Item Quantity Rate Amount")) {
      for (let i = index + 1; i < lines.length; i++) {
        const itemLine = lines[i];
        const itemData = parseInvoiceData([itemLine]);
        if (itemData.length > 0) {
          items.push(itemData[0]);
        }
      }
    }
    if (line.startsWith("Shipping: ")) {
      if (discountPercent == ""){
        targetLine = lines[index + 6];
        shippingFee = targetLine.split(" ")[1].trim();
      } else {
        targetLine = lines[index + 6];
        shippingFee = targetLine.split(" ")[2].trim();
      }
    }
    if (line.startsWith("Balance Due:")) {
      totalAmount = lines[index + 4];
    }
    if (line.startsWith("Order ID : ")) {
      orderId = line.split(" ")[3];
    }
  });

  return {
    customerName,
    address,
    orderDate,
    shipMode,
    invoice,
    items,
    discountPercent,
    discountAmount,
    shippingFee,
    totalAmount,
    orderId
  };
}

// 6. Export to Google Sheet
function exportToSheet() {
  const values = [
    ["Invoice", "Customer Name", "Address", "Order Date", "Ship Mode", "Product Name", "ProductID", "Sub-Category", "Category", "Quantity", "Unit Cost", "Subtotal", "Discount Percentage", "Discount Amount", "Shipping Fee", "Total Amount", "Order ID"]
  ];
  const ids = getPdfFiles();
  ids.forEach(id => {
    const text = readTextFromPdf(id);
    const { invoice, customerName, address, orderDate, shipMode, items, discountPercent, discountAmount, shippingFee, totalAmount, orderId } = getInvoiceDataFromText(text);
    items.forEach(item => {
      values.push([
        invoice,
        customerName,
        address,
        orderDate,
        shipMode,
        item.productName,
        item.productId,
        item.subCategory,
        item.category,
        item.quantity,
        item.unitPrice,
        item.subtotal,
        discountPercent,
        discountAmount,
        shippingFee,
        totalAmount,
        orderId
      ]);
    });
  });
  const outputSheet = SpreadsheetApp.getActive().getActiveSheet();
  outputSheet.clear();
  outputSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

// Analysis

function addImportantColumnsToNewSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  var data = sourceSheet.getDataRange().getValues();

  // Create a new sheet or get it if it already exists
  var newSheetName = "Important Data";
  var newSheet = ss.getSheetByName(newSheetName);
  if (!newSheet) {
    newSheet = ss.insertSheet(newSheetName);
  } else {
    newSheet.clear();  // Clear existing content if the sheet already exists
    newSheet.getRange(1, 1, newSheet.getMaxRows(), newSheet.getMaxColumns()).clearContent(); // Clear existing data
  }

  // Headers for new columns
  var headers = ["Order Date", "Order Month", "Category", "Product Name", "Quantity", "Total Amount", "Country"];
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Process each row and add the calculated columns
  var newData = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Extract necessary values from the original data
    var orderDate = new Date(row[3]); // Order Date is in column 4 (index 3)
    var orderMonth = orderDate.getMonth() + 1; // Get month from Order Date
    var category = row[8]; // Category is in column 9 (index 8)
    var productName = row[5]; // Product Name is in column 6 (index 5)
    var quantity = parseFloat(row[9]); // Quantity is in column 10 (index 9)
    var totalAmount = parseFloat(row[15].replace(/[^0-9.]/g, '')); // Total Amount is in column 16 (index 15)
    var address = row[2]; // Address is in column 3 (index 2)
    var country = address.split(', ').pop(); // Extract country from Address

    // Append the relevant columns to the new data array
    newData.push([orderDate, orderMonth, category, productName, quantity, totalAmount, country]);
  }

  // Write the processed data to the new sheet
  newSheet.getRange(2, 1, newData.length, headers.length).setValues(newData);
}

function createTotalSalesByMonthChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Important Data");

  // Get the range for Order Month and Total Amount columns
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange('B1:E' + lastRow);  // Assuming 'B' is Order Month and 'E' is Total Amount

  // Generate the chart
  var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(range)
      .setPosition(5, 5, 0, 0)
      .setOption('title', 'Total Sales by Month')
      .setOption('hAxis', {title: 'Month'})
      .setOption('vAxis', {title: 'Total Sales'})
      .build();
  
  sheet.insertChart(chart);
}

// Main function to process data and create chart
function main() {
  addImportantColumnsToNewSheet();
  createTotalSalesByMonthChart();
}


// Generate Graphs
function getData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return data;
}

function getOrCreateSheet(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    // Clear the existing sheet
    sheet.clear();
  } else {
    // Create a new sheet
    sheet = spreadsheet.insertSheet(sheetName);
  }

  return sheet;
}

// Total Amount by Category
function totalAmountByCategory() {
  var data = getData();
  var categoryTotals = {};

  for (var i = 1; i < data.length; i++) {
    var category = data[i][2];
    var amount = parseFloat(data[i][5]);
    if (!categoryTotals[category]) {
      categoryTotals[category] = 0;
    }
    categoryTotals[category] += amount;
  }

  var chartData = [];
  for (var category in categoryTotals) {
    chartData.push([category, categoryTotals[category]]);
  }

  var sheet = getOrCreateSheet('Total Amount by Category');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// Quantity Sold by Product
function quantitySoldByProduct() {
  var data = getData();
  var productQuantities = {};

  for (var i = 1; i < data.length; i++) {
    var product = data[i][3];
    var quantity = parseInt(data[i][4]);
    if (!productQuantities[product]) {
      productQuantities[product] = 0;
    }
    productQuantities[product] += quantity;
  }

  var chartData = [];
  for (var product in productQuantities) {
    chartData.push([product, productQuantities[product]]);
  }

  var sheet = getOrCreateSheet('Quantity Sold by Product');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// Total Amount by Month and Year
function totalAmountByMonthYear() {
  var data = getData();
  var monthYearTotals = {};

  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][0]); // Assuming the first column is the date
    var month = date.getMonth() + 1; // getMonth() returns 0-11, so add 1
    var year = date.getFullYear();
    var monthYear = month + '/' + year;
    var amount = parseFloat(data[i][5]);

    if (!monthYearTotals[monthYear]) {
      monthYearTotals[monthYear] = 0;
    }
    monthYearTotals[monthYear] += amount;
  }

  var chartData = [];
  for (var monthYear in monthYearTotals) {
    chartData.push([monthYear, monthYearTotals[monthYear]]);
  }

  var sheet = getOrCreateSheet('Total Amount by Month-Year');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// Total Amount by Country
function totalAmountByCountry() {
  var data = getData();
  var countryTotals = {};

  for (var i = 1; i < data.length; i++) {
    var country = data[i][6];
    var amount = parseFloat(data[i][5]);
    if (!countryTotals[country]) {
      countryTotals[country] = 0;
    }
    countryTotals[country] += amount;
  }

  var chartData = [];
  for (var country in countryTotals) {
    chartData.push([country, countryTotals[country]]);
  }

  var sheet = getOrCreateSheet('Total Amount by Country');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// Quantity Sold by Category
function quantitySoldByCategory() {
  var data = getData();
  var categoryQuantities = {};

  for (var i = 1; i < data.length; i++) {
    var category = data[i][2];
    var quantity = parseInt(data[i][4]);
    if (!categoryQuantities[category]) {
      categoryQuantities[category] = 0;
    }
    categoryQuantities[category] += quantity;
  }

  var chartData = [];
  for (var category in categoryQuantities) {
    chartData.push([category, categoryQuantities[category]]);
  }

  var sheet = getOrCreateSheet('Quantity Sold by Category');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// Sales Trends Over Time
function salesTrendsOverTime() {
  var data = getData();
  var dateAmounts = {};

  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][0]);
    var amount = parseFloat(data[i][5]);
    if (!dateAmounts[date]) {
      dateAmounts[date] = 0;
    }
    dateAmounts[date] += amount;
  }

  var chartData = [];
  for (var date in dateAmounts) {
    chartData.push([new Date(date), dateAmounts[date]]);
  }

  var sheet = getOrCreateSheet('Sales Trends Over Time');
  sheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, chartData.length, 2))
    .setPosition(1, 4, 0, 0)
    .build();

  sheet.insertChart(chart);
}


function onOpen(){
  var ui = SpreadsheetApp.getUi();
    ui.createMenu("DatabaseInput")
      .addItem("Import data", "exportToSheet")
      .addToUi()
    
    ui.createMenu('Analysis')
    .addItem('Run Analysis', 'main')
    .addToUi();

      ui.createMenu('Generate Charts')
    .addItem('Total Amount by Category', 'totalAmountByCategory')
    .addItem('Quantity Sold by Product', 'quantitySoldByProduct')
    .addItem('Total Amount by Month-Year', 'totalAmountByMonthYear')
    .addItem('Total Amount by Country', 'totalAmountByCountry')
    .addItem('Quantity Sold by Category', 'quantitySoldByCategory')
    .addItem('Sales Trends Over Time', 'salesTrendsOverTime')
    .addToUi();
}


