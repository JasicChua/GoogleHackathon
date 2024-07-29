const properties = PropertiesService.getScriptProperties().getProperties();
const geminiApiKey = properties['GOOGLE_API_KEY'];
const geminiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro-latest:generateContent?key=${geminiApiKey}`;
const baseUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=';
const geminiProVisionEndpoint = baseUrl + geminiApiKey;

const WORKSPACE_TOOLS = {
 "function_declarations": [
   {
        "name": "draftEmail",
        "description": "Write an email by analyzing data or charts in a Google Sheets file.",
        "parameters": {
          "type": "object",
          "properties": {
            "sheet_name": {
              "type": "string",
              "description": "The name of the sheet to analyze."
            },
            "recipient": {
              "type": "string",
              "description": "The gmail name of the recipient."
            }, 
          },
          "required": [
            "sheet_name",
            "recipient",
            
          ]
        }
      },      
 ]
};

function doGet(e) {
  let page = e.parameter.mode || "Main";
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Replace {{NAVBAR}} with the Navbar content
  htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}", getNavbar(page)));
  return htmlOutput;
}

function callGemini(prompt, temperature=0) {
  const payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          },
        ]
      }
    ], 
    "generationConfig":  {
      "temperature": temperature,
    },
  };

  const options = { 
    'method' : 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(geminiEndpoint, options);
  const data = JSON.parse(response);
  const content = data["candidates"][0]["content"]["parts"][0]["text"];
  return content;
}


function callGeminiProVision(prompt, image, temperature=0) {
  const imageData = Utilities.base64Encode(image.getAs('image/png').getBytes());

  const payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          },
          {
            "inlineData": {
              "mimeType": "image/png",
              "data": imageData
            }
          }          
        ]
      }
    ], 
    "generationConfig":  {
      "temperature": temperature,
    },
  };

  const options = { 
    'method' : 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(geminiProVisionEndpoint, options);
  const data = JSON.parse(response);
  const content = data["candidates"][0]["content"]["parts"][0]["text"];
  return content;
}


function callGeminiWithTools(prompt, tools, temperature=0) {
  const payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          },
        ]
      }
    ], 
    "tools" : tools,
    "generationConfig":  {
      "temperature": temperature,
    },    
  };

  const options = { 
    'method' : 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(geminiEndpoint, options);
  const data = JSON.parse(response);
  const content = data["candidates"][0]["content"]["parts"][0]["functionCall"];
  return content;
}


function testGeminiTools() {
  const prompt = "Tell me how many days there are left in this month.";
  const tools = {
    "function_declarations": [
      {
        "name": "datetime",
        "description": "Returns the current date and time as a formatted string.",
        "parameters": {
          "type": "string"
        }
      }
    ]
  };
  const output = callGeminiWithTools(prompt, tools);
  console.log(prompt, output);
}


function draftEmail(sheet_name, recipient) {
  
  const prompt = `Compose the email body for ${recipient} with your insights for this chart. Use information in this chart only and do not do historical comparisons. Be concise.`;

  var files = DriveApp.getFilesByName(sheet_name);
  var sheet = SpreadsheetApp.openById(files.next().getId()).getSheetByName("Important Data");
  var expenseChart = sheet.getCharts()[0];

  var chartFile = DriveApp.createFile(expenseChart.getBlob().setName("ImportantData.png"));
  var emailBody = callGeminiProVision(prompt, expenseChart);
  GmailApp.createDraft(recipient+"@gmail.com", "Invoice Data", emailBody, {
      attachments: [chartFile.getAs(MimeType.PNG)],
      name: 'myname'
  });
  // Delete the temporary chart file after attachment
  DriveApp.getFileById(chartFile.getId()).setTrashed(true);
}


function do_draft_email(username) {
  //const userQuery = "Set up a meeting at 10AM tomorrow with Helen to discuss the news in the Gemini-blog.txt file.";
  const userQuery = "Draft an email for " + username + " with insights from the chart in the Inv_db sheet.";           // change the name at here

  var tool_use = callGeminiWithTools(userQuery, WORKSPACE_TOOLS);
  Logger.log(tool_use);
  
  if(tool_use['name'] == "draftEmail") {
    draftEmail(tool_use['args']['sheet_name'], tool_use['args']['recipient']);
    Logger.log("Check your Gmail to review the draft");
  }
  else
    Logger.log("no proper tool found");
}
