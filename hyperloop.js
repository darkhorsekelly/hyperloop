// --- CONFIGURATION ---
// Update these constants to match your specific sheet names and column numbers.
// Make sure to run the createTrigger function once to set up the onEdit trigger.
const INPUT_SHEET_NAME = "";
const REPORT_SHEET_NAME = "";
const URL_COLUMN_NUMBER = 1; // Column where the image URL is pasted.
const TARGET_CELL_COLUMN_NUMBER = 1; // Column which contains the target cell address (e.g., "A10") for the report.
const MAX_WIDTH = ""; // Define the maximum width constraint in pixels

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
    const activeRow = range.getRow();
    const targetCellAddress = sheet.getRange(activeRow, TARGET_CELL_COLUMN_NUMBER).getValue();

    if (!targetCellAddress) {
      console.error(`${functionName}: ERROR! No target cell address found in cell ${TARGET_CELL_COLUMN_NUMBER}${activeRow}. Cannot proceed. Please check the ${INPUT_SHEET_NAME} sheet.`);
      return;
    }

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
      console.error(`${functionName}: ERROR! The address '${targetCellAddress}' (from cell ${TARGET_CELL_COLUMN_NUMBER}${activeRow}) is not a valid range in the '${REPORT_SHEET_NAME}' sheet.`);
      console.error(`--> Underlying Error: ${rangeError.message}`);
      return;
    }

    // 4. Determine the action based on the new and old cell values
    const url = e.value;
    const oldUrl = e.oldValue;

    // --- SCENARIO A: A valid URL was pasted or changed ---
    if (url && typeof url === 'string' && url.startsWith("http")) {
      console.log(`${functionName}: Valid URL detected in ${range.getA1Notation()}. Processing for row ${activeRow}.`);
      measureAndSetRowHeight(url, reportSheet, targetRange, activeRow);
    }

    // --- SCENARIO B: The cell was cleared ---
    else if (!url && oldUrl && typeof oldUrl === 'string' && oldUrl.startsWith("http")) {
      console.log(`${functionName}: URL removed from ${range.getA1Notation()}. Clearing target cell and resetting row.`);
      
      const targetRow = targetRange.getRow();

      // Clear the target cell in the report sheet.
      targetRange.clearContent();

      // Reset the row height to default.
      reportSheet.setRowHeight(targetRow, 21);

      console.log(`${functionName}: Cleared content in ${targetRange.getA1Notation()} and reset row ${targetRow} height to default.`);
    }

    // --- SCENARIO C: The cell was edited to a non-URL value ---
    else {
      console.log(`${functionName}: The new value in ${range.getA1Notation()} is not a valid URL. No action taken.`);
    }
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

      // 3. Check if image is too wide and calculate final dimensions.
      let finalHeight;
      let finalWidth;

      if (imageWidth > MAX_WIDTH) {
        // The image is too wide; it needs to be scaled down.
        const aspectRatio = imageHeight / imageWidth;
        finalWidth = maxWidth;
        finalHeight = Math.round(finalWidth * aspectRatio); // Calculate new height based on aspect ratio

        console.log(`Image width (${imageWidth}px) exceeds max width (${maxWidth}px). Resizing to ${finalWidth}w x ${finalHeight}h.`);

      } else {
        // The image is within the width limit, so use its original dimensions.
        finalWidth = imageWidth;
        finalHeight = imageHeight;

        console.log(`Image width (${imageWidth}px) is within the allowed limit. Using original dimensions.`);
      }

      // 4. Set the row height in the report sheet using the final calculated height.
      const targetRow = targetRange.getRow();

      // Use finalHeight for the calculation
      const newHeight = Math.round(finalHeight) + 5;
      console.log(`Setting row ${targetRow} to height ${newHeight}px.`);
      reportSheet.setRowHeight(targetRow, newHeight);

      // 5. IMPORTANT: Write the formula using mode 4 with the final dimensions.
      // Use finalHeight and finalWidth in the formula
      const formula = `=IMAGE("${url}", 4, ${finalHeight}, ${finalWidth})`;
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