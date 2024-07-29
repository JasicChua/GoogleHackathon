function doGet(e) {
  let page = e.parameter.mode || "Main";
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Replace {{NAVBAR}} with the Navbar content
  htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}", getNavbar(page)));
  return htmlOutput;
}


//Create Navigation Bar
function getNavbar(activePage) {
  var scriptURLHome = getScriptURL();
  var scriptURLPage2 = getScriptURL("mode=Page2");
  var scriptURLChart = getScriptURL("mode=Chart");
  var scriptURLEmail = getScriptURL("mode=Email");

  var navbar = 
    `<nav class="navbar navbar-expand-lg navbar-dark bg-dark">
        <div class="container">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNavAltMarkup">
          <div class="navbar-nav">
            <a class="navbar-brand" href="${scriptURLHome}">BetterPlan</a>
            <a class="nav-item nav-link ${activePage === 'Main' ? 'active' : ''}" href="${scriptURLHome}">Home</a>
            <a class="nav-item nav-link ${activePage === 'Chart' ? 'active' : ''}" href="${scriptURLChart}">Chart</a>
            <a class="nav-item nav-link ${activePage === 'Email' ? 'active' : ''}" href="${scriptURLEmail}">Email</a>
            <a class="nav-item nav-link" href="http://localhost:8501/" target="_blank">PPT</a>
            <a class="nav-item nav-link" href="https://docs.google.com/spreadsheets/d/1RODkzgSVCs4KkTENnmw5FzHxOgaWXk0UhWEJlCskqZM/edit?gid=181777901#gid=181777901" target="_blank">Sheet</a>
            
          </div>
        </div>
        </div>
            </nav>
      <style>
      .navbar {
      width: 100%;
      background-color: #3E4E42; 
      padding: 0.3rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .navbar-brand {
      color: white;
      text-decoration: none;
      font-size: 1.5rem;
      font-weight: bold;
      margin-left:30px;
    }

    .navbar-nav {
      list-style: none;
      margin: 1rem;
      padding: 0;
    }

    .navbar-nav .nav-item {
      margin: 1rem;
    }

    .navbar-nav .nav-link {
      color: #ffffff;
      text-decoration: none;
      font-size: 1rem;
    }

    .navbar-nav .nav-link:hover {
      color: #00ffcc;
      border-radius: 5px;
    }

    .navbar-nav .nav-link.active {
      color: #00ffcc;
    }
      </style>`;
  return navbar;
}


//returns the URL of the Google Apps Script web app
function getScriptURL(qs = null) {
  var url = ScriptApp.getService().getUrl();
  if(qs){
    if (qs.indexOf("?") === -1) {
      qs = "?" + qs;
    }
    url = url + qs;
  }
  return url;
}

//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



function getChartData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();
  
  const chartData = [['Category', 'Total Amount']];
  const categoryTotals = data.slice(1).reduce((acc, row) => {
    const category = row[2];
    const amount = parseFloat(row[5]);
    if (!acc[category]) acc[category] = 0;
    acc[category] += amount;
    return acc;
  }, {});

  for (const [category, total] of Object.entries(categoryTotals)) {
    chartData.push([category, total]);
  }

  return chartData;
}

function getProductQuantityData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();
  
  const chartData = [['Product', 'Quantity Sold']];
  const productQuantities = data.slice(1).reduce((acc, row) => {
    const product = row[3];
    const quantity = parseInt(row[4]);
    if (!acc[product]) acc[product] = 0;
    acc[product] += quantity;
    return acc;
  }, {});

  for (const [product, quantity] of Object.entries(productQuantities)) {
    chartData.push([product, quantity]);
  }

  return chartData;
}

function getTotalAmountByMonthYear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();
  
  const chartData = [['Month/Year', 'Total Amount']];
  const monthYearTotals = data.slice(1).reduce((acc, row) => {
    const date = new Date(row[0]);
    const monthYear = `${date.getMonth() + 1}/${date.getFullYear()}`;
    const amount = parseFloat(row[5]);
    if (!acc[monthYear]) acc[monthYear] = 0;
    acc[monthYear] += amount;
    return acc;
  }, {});

  for (const [monthYear, total] of Object.entries(monthYearTotals)) {
    chartData.push([monthYear, total]);
  }

  return chartData;
}

function getTotalAmountByCountry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();
  
  const chartData = [['Country', 'Total Amount']];
  const countryTotals = data.slice(1).reduce((acc, row) => {
    const country = row[6];
    const amount = parseFloat(row[5]);
    if (!acc[country]) acc[country] = 0;
    acc[country] += amount;
    return acc;
  }, {});

  for (const [country, total] of Object.entries(countryTotals)) {
    chartData.push([country, total]);
  }

  return chartData;
}

// Quantity Sold by Category
function quantitySoldByCategory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();

  const chartData = [['Category', 'Quantities']];
  const categoryQuantities = {};

  for (var i = 1; i < data.length; i++) {
    var category = data[i][2];
    var quantity = parseInt(data[i][4]);
    if (!categoryQuantities[category]) {
      categoryQuantities[category] = 0;
    }
    categoryQuantities[category] += quantity;
  }

  for (var category in categoryQuantities) {
    chartData.push([category, categoryQuantities[category]]);
  }
  return chartData;
}

function salesTrendsOverTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Important Data');
  const data = ss.getDataRange().getValues();

  // Add headers for the chart data
  const chartData = [['Date', 'Total Amount']];
  const dateAmounts = {}; // Initialize dateAmounts

  // Calculate total amount for each date
  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][0]);
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var amount = parseFloat(data[i][5]);
    if (!dateAmounts[formattedDate]) {
      dateAmounts[formattedDate] = 0;
    }
    dateAmounts[formattedDate] += amount;
  }

  // Convert the date amounts into an array for Google Charts
  for (var date in dateAmounts) {
    chartData.push([date, dateAmounts[date]]);
  }

  return chartData;
}
