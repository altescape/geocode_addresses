/* --- HELPER FUNCTIONS: START --- */

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  var dataRange = range ||
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(dataRange.getValues(), normalizeHeaders(headers));
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range;
// Returns an Array of objects.
function getColumnsData(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
  return getObjects(arrayTranspose(range.getValues()), headers);
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader(headers[i]));
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

/* --- HELPER FUNCTIONS: END --- */
  
/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
/*
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};
*/

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Geocode Address",
    functionName : "geoCode"
  }];
  sheet.addMenu("Address", entries);
};

function toaster(title, msg, time) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, time);
}

function geocode_this_cell() {
  // get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get sheet
  var sheet = ss.getSheets()[0];
  // get address column
  //var addressDataRange = ss.getRangeByName("addressDataRange");
  // get row data
  var data = getRowsData(sheet, SpreadsheetApp.getActiveRange());
  // Returns the selected cell
  var cell = sheet.getActiveCell();
  //geoCode(cell.getValue());
  //toaster('row data', data[0], 5);
  //Logger.log(data[1]['address']);
  var rowNum = SpreadsheetApp.getActiveRange().getRowIndex();
  for(var i = 0; i < data.length; i++){
    // there must be a better way than this of getting the row number
    // geoCode(data[i]['address']);
    // Logger.log(rowNum + '::' + SpreadsheetApp.getActiveRange().getRowIndex());
    //Logger.log(geoCode(data[i]['address']));
    // Set lat column value
    var latCol = sheet.getRange(rowNum, 9);
    latCol.setValue(geoCodeLat(data[i]['address']));
    // Set lng column value
    var lngCol = sheet.getRange(rowNum, 10);
    lngCol.setValue(geoCodeLng(data[i]['address']));
    rowNum ++;
    Utilities.sleep(1000);
  }
}

// DRY!!!!
function geoCodeLat(addr){
  var that = this;
  var gc = Maps.newGeocoder();
  var geoJSON = gc.geocode(addr);
  var lat = geoJSON.results[0].geometry.location.lat;
  //var lng = geoJSON.results[0].geometry.location.lng;
  //return lat + ", " + lng;
  //var msg = lat + ", " + lng;
  //toaster('Response', msg, 5);
  Logger.log(lat);
  return lat;
}
// DRY!!!!
function geoCodeLng(addr){
  var that = this;
  var gc = Maps.newGeocoder();
  var geoJSON = gc.geocode(addr);
  //var lat = geoJSON.results[0].geometry.location.lat;
  var lng = geoJSON.results[0].geometry.location.lng;
  //return lat + ", " + lng;
  //var msg = lat + ", " + lng;
  //toaster('Response', msg, 5);
  Logger.log(lng);
  return lng;
}
