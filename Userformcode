function submitSalesForm() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var inputSheet = ss.getSheetByName('Input Sheet Name');
    
    // Get input values, using getDisplayValue for E5 to preserve exact date format
    var inputFields = [
      inputSheet.getRange("E5").getDisplayValue(), // B (date as displayed)
      inputSheet.getRange("E6").getValue(), // C
      inputSheet.getRange("G6").getValue(), // A
      inputSheet.getRange("E7").getValue(), // D
      inputSheet.getRange("E8").getValue(), // E
      inputSheet.getRange("E9").getValue(), // F
      inputSheet.getRange("E10").getValue(), // G
      inputSheet.getRange("E11").getValue(), // H
      inputSheet.getRange("E12").getValue()  // I
    ];
    
    // Required fields: All
    var requiredFields = [inputFields[0], inputFields[1], inputFields[2], inputFields[3], inputFields[4], inputFields[5], inputFields[6], inputFields[7], inputFields[8]];
    if (requiredFields.some(value => value === "" || value === null)) {
      SpreadsheetApp.getUi().alert("Error: Please fill in all required fields before submitting.");
      return;
    }
    
    // Define destination column mapping (0-based indices for rowData array)
    var columnIndexes = { A:0, B:1, C:2, D:3, E:4, F:5, G:6, H:7, I:8 };
    var rowData = Array(9).fill("");
    
    // Explicitly map inputFields to rowData based on column assignments
    rowData[columnIndexes['A']] = inputFields[2];
    rowData[columnIndexes['B']] = inputFields[0];
    rowData[columnIndexes['C']] = inputFields[1];
    rowData[columnIndexes['D']] = inputFields[3];
    rowData[columnIndexes['E']] = inputFields[4];
    rowData[columnIndexes['F']] = inputFields[5];
    rowData[columnIndexes['G']] = inputFields[6];
    rowData[columnIndexes['H']] = inputFields[7];
    rowData[columnIndexes['I']] = inputFields[8];
    
    // Prepare payload for web app with action set to "insert"
    var payload = {
      action: "insert", // Added to specify insert action
      sheetName: "Destination Sheet Name",
      values: rowData,
      secretKey: "Web App Secret Key"
    };
    
    // Web app URL (replace with your deployed /exec URL)
    var webAppUrl = "Web App Url";  // e.g., https://script.google.com/macros/s/your-id/exec
    
    // Send POST request to web app with OAuth authentication
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      }
    };
    
    var response = UrlFetchApp.fetch(webAppUrl, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode === 200 && responseText === "Success") {
      clearSalesAfterSubmit();
    } else {
      var errorMsg = (responseCode !== 200) ? "HTTP Error " + responseCode + ": Check web app deployment access settings and scopes." : responseText;
      SpreadsheetApp.getUi().alert("Error: " + errorMsg);
    }
  } catch (e) {
    Logger.log("Error in submitSalesForm: " + e.toString());
    SpreadsheetApp.getUi().alert("An error occurred: " + e.message);
  }
}

function clearSalesAfterSubmit() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input');
    var fieldsToClear = ["E6", "E7", "E8", "E9", "E10", "E11", "E12"];
    
    fieldsToClear.forEach(cell => {
      sheet.getRange(cell).clearContent();
    });
  } catch (e) {
    Logger.log("Error in clearSalesAfterSubmit: " + e.toString());
  }
}

/**
 * Clears all fields except G6 when the clear button is clicked.
 */
function clearSalesAllFields() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input');
    var fieldsToClear = ["E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12"];
    
    fieldsToClear.forEach(cell => {
      if (cell !== "G6") {
        sheet.getRange(cell).clearContent();
      }
    });
  } catch (e) {
    Logger.log("Error in clearSalesAllFields: " + e.toString());
  }
}
