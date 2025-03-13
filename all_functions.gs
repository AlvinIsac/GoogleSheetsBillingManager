function myCustomOnEdit(e) {
  console.time("total time")
  processPayment(e);
  console.timeEnd("total time")
}

function processPayment(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();

  if (column !== 4) return;

  let customerName = sheet.getRange(row, 2).getValue();
  let stbNo        = sheet.getRange(row, 5).getValue();
  let amount       = sheet.getRange(row, 3).getValue();
  let subArea      = sheet.getRange(row, 6).getValue();
  let customerArea = sheet.getRange(row, 7).getValue();

  var selectedOption = range.getValue();
  var timestampCell  = sheet.getRange(row, MONTH_TIME_STAMP);
  var currentBgColor = sheet.getRange(row, 2).getBackground();

  let dateTime = new Date().toLocaleString('en-US', {
    month: 'short', 
    day: '2-digit', 
    year: 'numeric', 
    hour: '2-digit', 
    minute: '2-digit', 
    second: '2-digit', 
    hour12: true
  }).replace(
    /(\w{3}) (\d{2}), (\d{4}), (\d{2}:\d{2}:\d{2}) ([AP]M)/, 
    '$2-$1-$3, $4$5'
  );

  if (selectedOption === "PAID" && currentBgColor !== COLOR_PAID) {
    createDocFile(customerName, stbNo, amount, dateTime, subArea, customerArea);
    mainLogger(selectedOption, customerName, stbNo, dateTime, amount, subArea, customerArea);
    timestampCell.setValue(dateTime);
    sheet.getRange(row, 2).setBackground(COLOR_PAID);

  } else if (selectedOption === "NOT_PAID" && currentBgColor === COLOR_PAID) {
    mainLogger(selectedOption, customerName, stbNo, dateTime, amount, subArea, customerArea);
    timestampCell.setValue("");
    sheet.getRange(row, 2).setBackground(null);
    deleteDocFile(row, sheet, customerArea);
    }
}

function createDocFile(customerName, stbNo, amount, dateTime, subArea, customerArea) {
  const docFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  const reciptDocFolder = DriveApp.getFolderById(RECIPT_FOLDER_ID);
  const tempFile = docFile.makeCopy(reciptDocFolder);
  const tempDocFile = DocumentApp.openById(tempFile.getId());

  const body = tempDocFile.getBody();
  body.replaceText("{customerName}", customerName);
  body.replaceText("{stbNo}", stbNo);
  body.replaceText("{amount}", amount);
  body.replaceText("{dateTime}", dateTime);
  body.replaceText("{subArea}", subArea);
  body.replaceText("{customerArea}", customerArea);

  tempFile.setName(customerName + "_" + customerArea + "_" + new Date().toLocaleString('en-US', { month: 'long', day: '2-digit' }).replace(" ", "_"));
  tempDocFile.saveAndClose(); 
}

function deleteDocFile(customerName, customerArea) {
  let fileName = customerName + "_" + customerArea + "_" + new Date().toLocaleString('en-US', 
      { month: 'long', day: '2-digit' }).replace(" ", "_");

  const docFolder = DriveApp.getFolderById(RECIPT_FOLDER_ID);
  let files = docFolder.getFilesByName(fileName);

  while (files.hasNext()) {
    let file = files.next();
    file.setTrashed(true);
    Logger.log("🗑️ Google Doc Deleted: " + fileName);
  }
}

function mainLogger(selectedOption, customerName, stbNumber, dateTime, amount, subArea, customerArea) {
  const sheet = SpreadsheetApp.openById(ADMIN_ONLY_SS).getSheetByName("LOGGER"),
        lastRow = sheet.getLastRow() + 1;
  
  sheet.getRange(lastRow, 1, 1, 7).setValues([[customerName, selectedOption, stbNumber, dateTime, amount, subArea, customerArea]]);
  
  if (selectedOption === "PAID") {
    sheet.getRange(lastRow, 2).setBackground(COLOR_PAID);
  } else if (selectedOption === "NOT_PAID") {
    sheet.getRange(lastRow, 2).setBackground(COLOR_NOT_PAID);
  }
}


function monthlyEntryAllFolderFaster() {
  const folder = DriveApp.getFolderById(ALL_COLLECTION_FOLDER);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (files.hasNext()) {
    const file = files.next();
    const spreadsheet = SpreadsheetApp.open(file);
    const sheets = spreadsheet.getSheets();

    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return; // No data rows (only header or empty)

      // Grab all backgrounds in column B (row 2 down to lastRow)
      const range = sheet.getRange(2, 2, lastRow - 1, 1); // B2 : B
      const backgrounds = range.getBackgrounds(); // 2D array

      // Loop in memory (faster than calling setBackground row-by-row)
      for (let i = 0; i < backgrounds.length; i++) {
        const currentColor = backgrounds[i][0].toUpperCase();

        if (currentColor === "#90EE90") {
          // Green → White
          backgrounds[i][0] = "#FFFFFF";
        } else if (currentColor === "#FFFFFF") {
          // White → Yellow
          backgrounds[i][0] = "#FFFF00";
        }
      }
      range.setBackgrounds(backgrounds);
    });

    Logger.log("Colors updated for: " + spreadsheet.getName());
  }

  Logger.log("Colors updated for ALL spreadsheets in the folder!");
}

function resetDropDownAllFiles() {
    var folderId = ALL_COLLECTION_FOLDER; // Folder ID where spreadsheets are stored
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    while (files.hasNext()) {
        var file = files.next();
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        var sheets = spreadsheet.getSheets(); // Get all sheets in the spreadsheet
        
        sheets.forEach(function(sheet) {
            var lastRow = sheet.getLastRow(); // Get last row with data
            if (lastRow > 1) { // Ensure we don't reset an empty sheet
                sheet.getRange(2, 4, lastRow - 1, 1).setValue(""); // Clear column 4 (D) from row 2 onwards
            }
        });

        Logger.log("Reset Column 4 for all sheets in: " + file.getName());
    }

    Logger.log("All spreadsheets in the folder have been updated!");
}

function monthlyEntryAllFolder() {
    var folder = DriveApp.getFolderById(ALL_COLLECTION_FOLDER);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
        var file = files.next();
        var spreadsheet = SpreadsheetApp.openById(file.getId());
        var sheets = spreadsheet.getSheets();

        sheets.forEach(function(sheet) {
            var lastRow = sheet.getLastRow();

            for (var row = 2; row <= lastRow; row++) { // Skip header row
                var nameCell = sheet.getRange(row, 2); // Column B (Names)
                var currentColor = nameCell.getBackground().toUpperCase();

                if (currentColor === "#90EE90") { // Green → White
                    nameCell.setBackground("#FFFFFF");
                } else if (currentColor === "#FFFFFF") { // White → Yellow
                    nameCell.setBackground("#FFFF00");
                }
                // Yellow → No change
            }
        });

        Logger.log("Colors updated for: " + spreadsheet.getName());
    }

    Logger.log("Colors updated for ALL spreadsheets in the folder!");
}

function copyAllGreenEntriesToFinal() {
  
  const folder = DriveApp.getFolderById(ALL_COLLECTION_FOLDER);
  const files = folder.getFiles();

  // 2) Open the "master" destination spreadsheet & sheet
  const destSpreadsheet = SpreadsheetApp.openById(ADMIN_ONLY_SS);
  const finalSheet = destSpreadsheet.getSheetByName("FINAL_ENTRY");

  if (!finalSheet) {
    SpreadsheetApp.getUi().alert("'FINAL_ENTRY' sheet not found in destination spreadsheet!");
    return;
  }

  // 3) Clear existing rows (except header) on the FINAL_ENTRY sheet
  if (finalSheet.getLastRow() > 1) {
    finalSheet
      .getRange(2, 1, finalSheet.getLastRow() - 1, finalSheet.getLastColumn())
      .clearContent();
  }

  // We'll start inserting new rows at row 2
  let targetRow = 2;
  let totalEntries = 0;

  // The RGB for #90EE90 is (144, 238, 144)
  const targetColor = { red: 144, green: 238, blue: 144 };

  // ----------------------------------------
  // KEY PART: which columns do we want to copy?
  // ----------------------------------------
  // Suppose we pull columns B through K from the sheet
  // (that is 10 columns wide: B..K).
  // Then rowValues is an array of length 10: indexes [0..9].
  // You choose which indexes to copy here:
  // e.g. [0,1,2,7] => means: copy columns B, C, D, *and* column I.
  // If you want to copy them all, you'd do [0,1,2,3,4,5,6,7,8,9].
  const columnsToCopy = [0,1,2,3,4,5,MONTH_TIME_STAMP-2]; 
  
  // 4) Loop through all files in the folder
  while (files.hasNext()) {
    const file = files.next();

    // Only process Google Sheets
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const sourceSpreadsheet = SpreadsheetApp.open(file);
      const allSheets = sourceSpreadsheet.getSheets();

      // 5) For each sheet, check the rows
      allSheets.forEach(function (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return; // no data beyond header

        // ----------------------------------------
        // PULL A WIDER RANGE (B..K = 10 columns)
        // If you only need B..H, you'd do 7 columns, etc.
        // ----------------------------------------
        const numCols = 10; // for columns B..K
        const dataRange = sheet.getRange(2, 2, lastRow - 1, numCols); // (row=2,col=2) => B2
        const data = dataRange.getValues();
        const backgrounds = dataRange.getBackgroundObjects();

        // 6) Iterate each row in the data
        data.forEach(function (rowValues, rowIndex) {
          // Check the background color of Column B in the source
          // (index 0 in rowValues & backgrounds)
          const bgColorObj = backgrounds[rowIndex][0];

          // Check if it’s an RGB color (rather than theme-based or “none”)
          if (bgColorObj && bgColorObj.getColorType() === SpreadsheetApp.ColorType.RGB) {
            const rgb = bgColorObj.asRgbColor();
            // Compare to #90EE90
            if (rgb.getRed() === targetColor.red &&
                rgb.getGreen() === targetColor.green &&
                rgb.getBlue() === targetColor.blue) {
              
              // Build an array of just the columns we want
              // columnsToCopy = e.g. [0,1,2,7]
              const rowDataToCopy = columnsToCopy.map(index => rowValues[index]);

              // Write that row to FINAL_ENTRY
              // Start at column A in FINAL_ENTRY, length is rowDataToCopy.length
              finalSheet
                .getRange(targetRow, 1, 1, rowDataToCopy.length)
                .setValues([rowDataToCopy]);
              
              targetRow++;
              totalEntries++;
            }
          }
        });
      });
    }
  }

  // 8) Final notification with the total number of green rows found
  // SpreadsheetApp.getUi().alert("Finished copying " + totalEntries + " row(s) into FINAL_ENTRY!");
}

function resetAll() {
    resetDropDownAllFiles(); // Call existing function

    const folder = DriveApp.getFolderById(ALL_COLLECTION_FOLDER);
    const files = folder.getFiles();

    while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
            const spreadsheet = SpreadsheetApp.open(file);
            const sheets = spreadsheet.getSheets();

            sheets.forEach(sheet => {
                const lastRow = sheet.getLastRow();
                const lastColumn = sheet.getLastColumn(); // Get the last column with data
                
                if (lastRow > 1) {
                    sheet.getRange(2, 2, lastRow - 1, 1).setBackground(null); // Reset color for Name column (Column B)
                    
                    if (lastColumn >= 8) {
                        sheet.getRange(2, 8, lastRow - 1, lastColumn - 7).clearContent(); // Clear all columns from H onwards
                    }
                }
            });
        }
    }

    // Reset LOGGER sheet in ADMIN_ONLY_SS
    const adminSheet = SpreadsheetApp.openById(ADMIN_ONLY_SS).getSheetByName("LOGGER");
    if (adminSheet) {
        const lastRow = adminSheet.getLastRow();
        const lastColumn = adminSheet.getLastColumn();

        if (lastRow > 1) {
            adminSheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent().setBackground(null);
        }
    }

    // Delete all files in RECIPT_FOLDER_ID
    const receiptFolder = DriveApp.getFolderById(RECIPT_FOLDER_ID);
    const receiptFiles = receiptFolder.getFiles();

    while (receiptFiles.hasNext()) {
        const file = receiptFiles.next();
        file.setTrashed(true); // Move to trash instead of permanently deleting
        Logger.log("Deleted Receipt File: " + file.getName());
    }

    Logger.log("Reset process completed successfully!");
}

function logTemplateFormatting() {
    const templateSpreadsheet = SpreadsheetApp.openById("1JZMtSir0I2hlXI7JRNgn7IU-npPpLTAk1p-fCHGb4r4"); // Replace with your template spreadsheet ID
    const templateSheet = templateSpreadsheet.getSheets()[0]; // First sheet of the template (modify if needed)

    const lastColumn = templateSheet.getLastColumn();
    const lastRow = templateSheet.getLastRow();

    let logDetails = " TEMPLATE FORMATTING DETAILS \n\n";

    //  Log column widths
    logDetails += "Column Widths:\n";
    for (let col = 1; col <= lastColumn; col++) {
        const width = templateSheet.getColumnWidth(col);
        logDetails += `Column ${col}: ${width}px\n`;
    }

    //  Log row heights
    logDetails += "\n Row Heights:\n";
    for (let row = 1; row <= lastRow; row++) {
        const height = templateSheet.getRowHeight(row);
        logDetails += `Row ${row}: ${height}px\n`;
    }

    // Log hidden columns
    logDetails += "\n Hidden Columns:\n";
    let hiddenCols = [];
    for (let col = 1; col <= lastColumn; col++) {
        if (templateSheet.isColumnHiddenByUser(col)) {
            hiddenCols.push(`Column ${col}`);
        }
    }
    logDetails += hiddenCols.length ? hiddenCols.join(", ") + "\n" : "None\n";

    //  Log hidden rows
    logDetails += "\n Hidden Rows:\n";
    let hiddenRows = [];
    for (let row = 1; row <= lastRow; row++) {
        if (templateSheet.isRowHiddenByUser(row)) {
            hiddenRows.push(`Row ${row}`);
        }
    }
    logDetails += hiddenRows.length ? hiddenRows.join(", ") + "\n" : "None\n";

    //  Log frozen rows & columns
    const frozenRows = templateSheet.getFrozenRows();
    const frozenColumns = templateSheet.getFrozenColumns();
    logDetails += `\n Frozen Rows: ${frozenRows}\n`;
    logDetails += ` Frozen Columns: ${frozenColumns}\n`;

    Logger.log(logDetails);
}

function settings(e) {
    var range = e.range;
    var column = range.getColumn();
    var row = range.getRow();

    if (column === 5 && (row === 6 || row === 8 || row === 11 || row === 14) && range.getValue() === true) {

        range.setBackground("#FF0000"); 
        SpreadsheetApp.flush(); //  Forces immediate UI update

        if (row === 6) {
            copyAllGreenEntriesToFinal();
        } else if (row === 8) {
            monthlyEntryAllFolder();
            resetDropDownAllFiles();
        } else if (row === 14) {
            resetAll();
        } else if (row === 11){
            monthlyEntryAllFolderFaster();
            resetDropDownAllFiles();
        }

        range.setBackground(null); 
        range.setValue(false); 
    }
}


