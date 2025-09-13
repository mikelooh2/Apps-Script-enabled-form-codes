// Define sheet configurations for identifier columns and formula ranges if needed
var sheetConfigs = {
  "Online Sales Data": {
    idColumn: "Transaction ID" // Unique identifier column 
  }
};

function doGet() {
  // Optional: Add this for testing the web app via GET in a browser
  return ContentService.createTextOutput("Web app is deployed and running.").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    // Log the incoming data for debugging
    Logger.log("Received data: " + JSON.stringify(data));
    
    // Check for secret key
    if (data.secretKey !== "Your Key") {
      throw new Error("Unauthorized request");
    }
    
    // Validate required field: sheetName
    if (!data.sheetName) {
      throw new Error("Missing sheetName in the request");
    }
    
    // Validate action field
    if (!data.action) {
      throw new Error("Missing action in the request");
    }
    
    var sheetName = data.sheetName;
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }
    
    if (data.action === "insert") {
      // Validate values for insert
      if (!data.values || !Array.isArray(data.values) || data.values.length === 0) {
        throw new Error("Missing or invalid values for insert");
      }
      
      // Append the new row
      sheet.appendRow(data.values);
      
      // Get the new row number
      var newRow = sheet.getLastRow();
      
      // Copy formulas if applicable
      var config = sheetConfigs[sheetName];
      if (config && config.formulaRanges && newRow > 1) {
        var lastRow = newRow - 1;
        config.formulaRanges.forEach(function(range) {
          var startCol = range.startCol;
          var endCol = range.endCol;
          var numCols = endCol - startCol + 1;
          var sourceRange = sheet.getRange(lastRow, startCol, 1, numCols);
          var targetRange = sheet.getRange(newRow, startCol, 1, numCols);
          sourceRange.copyTo(targetRange, { contentsOnly: false });
        });
      }
    } else if (data.action === "update") {
      // Validate fields for update
      var config = sheetConfigs[sheetName];
      if (!config || !config.idColumn) {
        throw new Error("Update not supported for this sheet");
      }
      if (!data.identifier || !data.updates || typeof data.updates !== 'object') {
        throw new Error("Missing identifier or updates for update");
      }
      
      var idColumnName = config.idColumn;
      var idColumnIndex = getColumnIndex(sheet, idColumnName);
      var identifier = data.identifier;
      var updates = data.updates;
      
      // Find the row to update
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) { // Start from 1 to skip header
        if (values[i][idColumnIndex - 1] == identifier) {
          rowIndex = i + 1; // 1-based row index
          break;
        }
      }
      if (rowIndex === -1) {
        throw new Error("Identifier not found: " + identifier);
      }
      
      // Update the specified columns
      for (var columnName in updates) {
        var colIndex = getColumnIndex(sheet, columnName);
        sheet.getRange(rowIndex, colIndex).setValue(updates[columnName]);
      }
    } else if (data.action === "delete") {
      // Validate fields for delete
      var config = sheetConfigs[sheetName];
      if (!config || !config.idColumn) {
        throw new Error("Delete not supported for this sheet");
      }
      if (!data.identifier) {
        throw new Error("Missing identifier for delete");
      }
      
      var idColumnName = config.idColumn;
      var idColumnIndex = getColumnIndex(sheet, idColumnName);
      var identifier = data.identifier;
      
      // Find the row to delete
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) { // Start from 1 to skip header
        if (values[i][idColumnIndex - 1] == identifier) {
          rowIndex = i + 1; // 1-based row index
          break;
        }
      }
      if (rowIndex === -1) {
        throw new Error("Identifier not found: " + identifier);
      }
      
      // Delete the row
      sheet.deleteRow(rowIndex);
    } else if (data.action === "callFunction") {
      // Validate functionName for callFunction action
      if (!data.functionName) {
        throw new Error("Missing functionName for callFunction action");
      }
      
      var func = functionMap[data.functionName];
      if (typeof func === 'function') {
        func();
      } else {
        throw new Error("Function not allowed: " + data.functionName);
      }
    } else {
      throw new Error("Invalid action: " + data.action);
    }
    
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    Logger.log("Error: " + error.message);
    return ContentService.createTextOutput("Error: " + error.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

// Helper function to get 1-based column index from column name
function getColumnIndex(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = headers.indexOf(columnName);
  if (index === -1) {
    throw new Error("Column not found: " + columnName);
  }
  return index + 1;
}
