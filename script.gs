
// Define the range of cells containing JIRA issue keys
var issueKeysRangeStartColumn = 4; // 1-based
var issueKeysRangeStartRow = 2; // 1-based
var issueKeysRangeNumCols = 2; // One for iOS and one for Android

var jiraDomain = "mycompany.atlassian.net";

var authToken = "eW91cl9qaXJhX3VzZXI6eW91cl9qaXJhX3Bhc3N3b3Jk"; // Your JIRA "user:pass" string, base64 encoded

var issuesUpdateInProgressToastTitle = "Processing";
var issuesUpdateInProgressToastMessage = "Updating JIRA issues status...";
var issuesUpdateDoneToastTitle = "Done!";
var issuesUpdateDoneToastMessage = "JIRA issues were updated successfully.";
var onOpenToastTitle = "¡Hey!";
var onOpenToastMessage = "Keep in mind you can now update JIRA issues status from the top menu...  😉";

var menuOptionTitle = "📥 Update from JIRA";
var submenuOptionTitle = "Update issues from JIRA";

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

var issuesStatuses = null; // Lazy initialization

var JIRAColorToLocalColorMap = {
  "green" : "#d9ead3",
  "blue-gray" : "white",
  "yellow" : "#fff2cc"
};

function statusColorForIssue(issueKey) {
  for (i=0; i < issuesStatuses.issues.length; i++) {
    var issue = issuesStatuses.issues[i];
    if (issue.key == issueKey) {
      if (issue.fields.status.name == "Blocked") {
        return "#e6b8af";
      } else {
        var JIRAColor = issue.fields.status.statusCategory.colorName;
        return JIRAColorToLocalColorMap[JIRAColor] || "#e6b8af";
      }
    }
  }
  
  return "#e6b8af";
}

function retrieveIssuesStatuses() {
  var postBody = {
    "jql" : buildJQLQueryWithCurrentIssues(),
    "startAt": 0,
    "maxResults": 10000,
    "fields": [
      "status"
    ]
  };
  
  var options = {
    "method" : "post",
    "contentType" : "application/json",
    "headers" : {
      "Authorization" : "Basic " + authToken
    },
    "payload" : JSON.stringify(postBody)
  };

  var response = UrlFetchApp.fetch("https://" + jiraDomain + "/rest/api/2/search", options);  
  //Logger.log("Response status: " + response.getResponseCode());
  //Logger.log("Response string: " + response.getContentText());

  return JSON.parse(response.getContentText());
}

function loopThroughIssueKeysExecutingFunction(func) {
  var dataRange = sheet.getDataRange();
  //Logger.log("Data range: rows = " + dataRange.getNumRows() + " cols = " + dataRange.getNumColumns());
  
  var issueKeysRange = sheet.getRange(issueKeysRangeStartRow, issueKeysRangeStartColumn, dataRange.getNumRows() - 1, issueKeysRangeNumCols);
  //Logger.log("issueKeys range: rows = " + issueKeysRange.getNumRows() + " cols = " + issueKeysRange.getNumColumns());

  var issueKeysRangeValues = issueKeysRange.getValues();
  //Logger.log("issueKeys range values: rows = " + issueKeysRangeValues.length + " cols = " + issueKeysRangeValues[0].length);
  
  for (var row = 0; row < issueKeysRange.getNumRows(); row++) {
    for (var col = 0; col < issueKeysRange.getNumColumns(); col++) {
      var currentValue = issueKeysRangeValues[row][col];
      //Logger.log("Processing row=" + row + " col=" + col + " --> " + currentValue);
      func(row, col, currentValue);
    }
  }
}

function buildJQLQueryWithCurrentIssues() {
  var issueKeys = "DUMMYKEY-0"; // To make concatenation simpler and prevent empty lists
  loopThroughIssueKeysExecutingFunction(function (row, col, cellValue) {
    if (cellValue.length > 0 && cellValue.indexOf("-") > 0) { // Ignore cells that don't look like a JIRA issue key (don't contain a dash)
      issueKeys = issueKeys + "," + cellValue;
    }
  });
  return "issueKey in (" + issueKeys + ")";
}

function refreshIssuesStatuses() {
  SpreadsheetApp.getActiveSpreadsheet().toast(issuesUpdateInProgressToastMessage, issuesUpdateInProgressToastTitle, 600);
  
  issuesStatuses = retrieveIssuesStatuses();
  
  loopThroughIssueKeysExecutingFunction(function (row, col, cellValue) {
    if (cellValue.length > 0) {
        if (cellValue.indexOf("-") <= 0) { // Mark cells that don't look like a JIRA issue key (don't contain a dash) with another color
          currentColor = "#c9daf8";
        } else {
          currentColor = statusColorForIssue(cellValue);
        }
        var currentCellRange = sheet.getRange(row + issueKeysRangeStartRow, col + issueKeysRangeStartColumn, 1, 1);
        currentCellRange.setBackground(currentColor);
      }
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast(issuesUpdateDoneToastMessage, issuesUpdateDoneToastTitle, 5);
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [{name: submenuOptionTitle, functionName: "refreshIssuesStatuses"}];
  spreadsheet.addMenu(menuOptionTitle, menu);
  SpreadsheetApp.getActiveSpreadsheet().toast(onOpenToastMessage, onOpenToastTitle, 600);
}


