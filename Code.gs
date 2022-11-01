// Add API credentials
var APIKEY = "";
var SEARCHENGINEID = "";

/******************************************************************************************************
 * Connect to Google Image Search via API, return results to be placed in a new sheet.
 * 
 * Instructions
 * 1. Get your Custom Search JSON API key and add it to var APIKEY: https://developers.google.com/custom-search/v1/overview#api_key
 * 2. Create search engine, point it to google.com: https://cse.google.com/all
 * 3. In the settings, tell it to enable Image Search, remove any Sites to search, and Search the Entire Web.
 * 4. Copy the search engine ID and add it to var SEARCHENGINEID.
 * 5. Run the script onOpen() and refresh the Google Sheet.
 * 6. Run the script 'Function: Get Google Image Search Result(s)' from the new Functions menu on the Google Sheet.
 * 
 * Sources
 * https://stackoverflow.com/questions/34035422/google-image-search-says-api-no-longer-available
 * https://webmasters.stackexchange.com/questions/18704/return-first-image-source-from-google-images
 * 
 ******************************************************************************************************/

function startSearch() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("Enter your search term:");
  if (prompt.getSelectedButton() == ui.Button.OK) {
    runQuery(prompt.getResponseText().trim());
  }
}

/******************************************************************************************************
 * Get the results of your search.
 * 
 * @param {String} query The search value we are searching using the Google Custom Search Engine.
 * @param {String} start The search query is limited to 10 results per call. If we're calling it again for the same query, we'll want to bump this # up. (Optional)
 * @return {Array} The responses returned from the custom search.
 * 
 * Sources
 * https://stackoverflow.com/questions/34035422/google-image-search-says-api-no-longer-available
 * https://webmasters.stackexchange.com/questions/18704/return-first-image-source-from-google-images
 * 
 ******************************************************************************************************/

function getGoogleImageSearchResult(query, start) {

  // Declare variables
  var numberOfResults = 10;
  var searchType = "image";
  var start = start || 1;

  // Building call to API: https://developers.google.com/custom-search/v1/reference/rest/v1/cse/list
  var url = "https://www.googleapis.com/customsearch/v1?key=" + APIKEY + "&cx=" + SEARCHENGINEID
    + "&q=" + query + "&num=" + numberOfResults + "&searchType=" + searchType + "&start=" + start;
  console.log(url);

  var params = {
    method: "GET",
    // muteHttpExceptions: true
  };

  // Calling API
  var response = UrlFetchApp.fetch(url, params);

  // Parsing response
  return JSON.parse(response.getContentText());
}
/******************************************************************************************************
 *
 * Return the Google Image Search results for the search item and place them in a new sheet.
 * 
 * @param {String} value The search value we are searching using the Google Custom Search Engine.
 * 
 ******************************************************************************************************/

function runQuery(value) {

  // Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  var value = value || range.getDisplayValue();
  var results = {};
  var returnArray = [];
  var maxResults = 100;

  // Add search results to array
  for (var x = 1; x <= maxResults; x += 10) {
    results = getGoogleImageSearchResult(value, x);
    returnArray = returnArray.concat(results.items);
  }

  // Map results to sheet
  setArraySheet(returnArray, value, spreadsheet, "contextLink", "image", "Reference Page");
}

/******************************************************************************************************
 * 
 * Convert array into sheet.
 * 
 * @param {Array} array The array that we need to map to a sheet
 * @param {String} sheetName The name of the sheet the array is being mapped to
 * @param {Object} spreadsheet The source spreadsheet
 * @param {String} param The name of the parameter we need for the returned API object, optional
 * @param {String} ogColHeader The name of the column header getting replaced for readability, optional
 * @param {String} replacementColHeader The new name of the replaced column header, optional
 * 
 ******************************************************************************************************/

function setArraySheet(array, sheetName, spreadsheet, param, ogColHeader, replacementColHeader) {

  // Declare variables
  var spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var keyArray = [];
  var memberArray = [];
  var sheetRange = "";
  var index = -1;
  var ogColHeader = ogColHeader || "";
  var replacementColHeader = replacementColHeader || "";

  // Define an array of all the returned object's keys to act as the Header Row
  keyArray.length = 0;
  if (param) {
    keyArray = Object.keys(array[0]);

    index = keyArray.indexOf(ogColHeader);

    if (index !== -1) {
      keyArray[index] = replacementColHeader;
    }

  }
  else {
    keyArray = Object.keys(array[0]);
  }
  memberArray.length = 0;
  memberArray.push(keyArray);

  //  Capture members from returned data
  for (var x = 0; x < array.length; x++) {
    memberArray.push(keyArray.map(function (key) {
      if (key == replacementColHeader) {
        return array[x][ogColHeader][param];
      } else {
        return array[x][key];
      }
    }));
  }

  // Select or create the sheet
  try {
    sheet = spreadsheet.insertSheet(sheetName);
  } catch (e) {
    sheet = spreadsheet.getSheetByName(sheetName).clear();
  }

  // Set values  
  sheetRange = sheet.getRange(1, 1, memberArray.length, memberArray[0].length);
  sheetRange.setValues(memberArray);
}

/******************************************************************************************************
* 
* Create a menu option for script functions
*
******************************************************************************************************/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
    .addItem('Function: Get Google Image Search Result(s)', 'startSearch')
    .addToUi();
}
