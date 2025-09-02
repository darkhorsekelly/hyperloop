// --- CONFIGURATION ---
// Update these constants to match your specific sheet names and column numbers.
// Make sure to run the createTrigger function once to set up the onEdit trigger.
const INPUT_SHEET_NAME = "";
const REPORT_SHEET_NAME = "";
const URL_COLUMN_NUMBER = 1; // Column where the image URL is pasted.
const TARGET_CELL_COLUMN_NUMBER = 1; // Column which contains the target cell address (e.g., "A10") for the report.

/**
 * Creates an installable 'onEdit' trigger for the handleEdit function.
 * Run this function once from the script editor to set up the automation.
 
function createTrigger() {
  // Check if a trigger for handleEdit already exists to avoid duplicates.
  const allTriggers = ScriptApp.getProjectTriggers();
  const triggerExists = allTriggers.some(trigger => trigger.getHandlerFunction() === 'handleEdit');

  if (!triggerExists) {
    const sheet = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('handleEdit')
      .forSpreadsheet(sheet)
      .onEdit()
      .create();
    console.log("Successfully created the onEdit trigger for handleEdit.");
  } else {
    console.log("The onEdit trigger for handleEdit already exists. No action taken.");
  }
}
*/

/**
 * Main function that runs via an installable "On edit" trigger.
 * It validates the edit, gets image dimensions from a URL, and sets a row height in another sheet.
 * @param {Object} e The event object provided by the trigger.
 */
function handleEdit(e) {
  const functionName = "handleEdit";
  console.log(`--- Execution starting for ${functionName} ---`);

  // Use a try-catch block to handle any unexpected errors gracefully.
  try {
    // 1. Log the event object for detailed debugging.
    if (!e || !e.range) {
      console.warn(`${functionName}: Function was likely run manually or the event object is malformed. Exiting.`);
      return;
    }
    console.log(`${functionName}: Event received. Range: ${e.range.getA1Notation()}, New Value: "${e.value}", Old Value: "${e.oldValue || '(empty)'}"`);

    // 2. Validate the context of the edit to ensure we're on the right sheet and column.
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();

    if (sheetName !== INPUT_SHEET_NAME) {
      console.log(`${functionName}: Edit was on sheet '${sheetName}', not the target sheet '${INPUT_SHEET_NAME}'. Aborting.`);
      return;
    }

    if (range.getColumn() !== URL_COLUMN_NUMBER) {
      console.log(`${functionName}: Edit was in column ${range.getColumn()}, not the target column ${URL_COLUMN_NUMBER}. Aborting.`);
      return;
    }

    // 3. Validate and process the URL from the edited cell.
    const url = e.value;
    if (!url || typeof url !== 'string' || !url.startsWith("http")) {
      console.log(`${functionName}: The new value '${url}' in cell ${range.getA1Notation()} is not a valid URL. Aborting.`);
      return;
    }

    const activeRow = range.getRow();
    console.log(`${functionName}: Valid URL detected in ${range.getA1Notation()}. Processing for row ${activeRow}.`);

    // 4. Get the target cell address from the corresponding column in the same row.
    const targetCellAddress = sheet.getRange(activeRow, TARGET_CELL_COLUMN_NUMBER).getValue();
    if (!targetCellAddress) {
      console.error(`${functionName}: ERROR! No target cell address found in cell G${activeRow}. Cannot proceed. Please check the G${INPUT_SHEET_NAME} sheet.`);
      return;
    }
    console.log(`${functionName}: Target cell address found in G${activeRow}: '${targetCellAddress}'.`);

    // 5. Get the report sheet and validate that the target range is valid.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);
    if (!reportSheet) {
      console.error(`${functionName}: FATAL ERROR! The report sheet named '${REPORT_SHEET_NAME}' could not be found. Aborting.`);
      return;
    }

    let targetRange;
    try {
      targetRange = reportSheet.getRange(targetCellAddress);
    } catch (rangeError) {
      console.error(`${functionName}: ERROR! The address '${targetCellAddress}' (from cell G${activeRow}) is not a valid range in the '${REPORT_SHEET_NAME}' sheet.`);
      console.error(`--> Underlying Error: ${rangeError.message}`);
      return;
    }

    // 6. Perform the core logic: measure the image and set the row height.
    measureAndSetRowHeight(url, reportSheet, targetRange, activeRow);

  } catch (error) {
    // This is a catch-all for any other unexpected errors in the main function.
    console.error(`${functionName}: An unexpected error occurred in the main execution block. Details: ${error.message}`);
    console.error(error.stack); // Log stack for advanced debugging.
  } finally {
    console.log(`--- Execution finished for ${functionName} ---`);
  }
}

/**
 * Fetches an image, measures its height using a temporary Google Slide,
 * and sets the row height in a target sheet.
 * @param {string} url The URL of the image to measure.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet The sheet where the row height will be set.
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange The cell/range that determines the target row.
 * @param {number} sourceRow The row number from the input sheet, used for logging context.
 */
/**
 * Measures an image's true dimensions, sets the target row height,
 * and writes a formula to render the image at its original size.
 * @param {string} url The URL of the image.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet The sheet where the image will be placed.
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange The cell where the formula will be written.
 * @param {number} sourceRow The row number from the input sheet for logging.
 */
function measureAndSetRowHeight(url, reportSheet, targetRange, sourceRow) {
  const functionName = "measureAndSetRowHeight";
  let tempDoc = null; 

  try {
    // 1. Fetch the image blob.
    console.log(`${functionName}: Attempting to fetch image blob from URL: ${url}`);
    const blob = UrlFetchApp.fetch(url).getBlob();
    console.log(`${functionName}: Successfully fetched image blob.`);

    // 2. Use a temporary Google Doc to measure the true, unscaled dimensions.
    console.log(`${functionName}: Creating temporary Google Doc for measurement...`);
    const docName = `temp_image_measuring_doc_${new Date().getTime()}`;
    tempDoc = DocumentApp.create(docName);
    const body = tempDoc.getBody();
    const image = body.appendImage(blob);

    const imageHeight = image.getHeight();
    const imageWidth = image.getWidth(); 
    console.log(`Successfully measured image dimensions: ${imageWidth}w x ${imageHeight}h.`);

    // 3. Set the row height in the report sheet.
    const targetRow = targetRange.getRow();
    const newHeight = Math.round(imageHeight) + 5; // Add 5px padding.
    console.log(`Setting row ${targetRow} to height ${newHeight}px.`);
    reportSheet.setRowHeight(targetRow, newHeight);

    // 4. ✨ IMPORTANT: Write the formula using mode 4 to force the correct size.
    const formula = `=IMAGE("${url}", 4, ${imageHeight}, ${imageWidth})`;
    console.log(`Setting formula in cell ${targetRange.getA1Notation()}: ${formula}`);
    targetRange.setFormula(formula);
    
    console.log(`${functionName}: SUCCESS! Row height and image formula set.`);

  } catch (error) {
    console.error(`${functionName}: SCRIPT FAILED for source row ${sourceRow}. URL: ${url}. Error: ${error.message}`);
    console.error(error.stack);
    // Optional: Write an error message back to the cell.
    targetRange.setValue(`Error measuring image: ${error.message}`);
  } finally {
    // 5. Clean up the temporary file.
    if (tempDoc) {
      try {
        DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
        console.log(`${functionName}: Cleaned up and deleted temporary doc file.`);
      } catch (cleanupError) {
        console.error(`${functionName}: FAILED to clean up temp file ID: ${tempDoc.getId()}`);
      }
    }
  }
}