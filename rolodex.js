/***** CONFIGURATION *****/
const TEMPLATE_SHEET = ""; // Sheet where the municipality is selected and outputs go.
const TEMPLATE_MUNI_A1 = ""; // Cell with the municipality dropdown.

// Rolodex (Source of Truth) Sheet
const ROLODEX_ID = ""; // The ID of the Rolodex Google Sheet.
const ROLODEX_TAB = ""; // The specific tab within the Rolodex sheet.
const ROLODEX_MUNI_COL = 1; // Column contains municipality names.
const ROLODEX_TEXT_COL = 2; // Column contains the full rich text instructions.

// These markers define the start and end of each section within the rich text.
const SECTIONS = [
  { name: "Appraiser", start: /1\.\s*▇\s*Appraiser:/, stop: /2\.\s*▇\s*Taxes:/,     targetA1: "B14" },
  { name: "Taxes",     start: /2\.\s*▇\s*Taxes:/,     stop: /3\.\s*▇\s*Utilities:/,  targetA1: "B22" },
  { name: "Utilities", start: /3\.\s*▇\s*Utilities:/, stop: /4\.\s*▇\s*Permits:/,   targetA1: "B30" },
  { name: "Permits",   start: /4\.\s*▇\s*Permits:/,   stop: /5\.\s*▇\s*Code:/,      targetA1: "B43" },
  { name: "Code",      start: /5\.\s*▇\s*Code:/,      stop: /6\.\s*▇\s*Special:/,   targetA1: "B68" },
  { name: "Special",   start: /6\.\s*▇\s*Special:/,   stop: /7\.\s*▇\s*Contact Info:/, targetA1: "B93" }
];


/***** MENU *****/

/**
 * Creates a "Rolodex" menu in the spreadsheet UI when the file is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Rolodex")
    .addItem("Fill Instructions from Rolodex", "fillFormattedSections")
    .addSeparator()
    .addItem("Enable Auto-Update", "enableMuniAuto")
    .addItem("Disable Auto-Update", "disableMuniAuto")
    .addToUi();
}


/***** PUBLIC ACTIONS *****/

/**
 * Main function to find the selected municipality's data and populate the sheet.
 */
function fillFormattedSections() {
  const functionName = "fillFormattedSections";
  console.log(`--- ${functionName}: Starting execution ---`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET);
    if (!templateSheet) {
      SpreadsheetApp.getUi().alert(`Error: Sheet "${TEMPLATE_SHEET}" not found.`);
      console.error(`${functionName}: Template sheet not found.`);
      return;
    }

    const muni = templateSheet.getRange(TEMPLATE_MUNI_A1).getDisplayValue().trim();
    if (!muni) {
      SpreadsheetApp.getUi().alert(`Please select a municipality in cell ${TEMPLATE_MUNI_A1}.`);
      console.log(`${functionName}: No municipality selected. Aborting.`);
      return;
    }
    console.log(`${functionName}: Looking for municipality: "${muni}"`);

    const rolodexSheet = SpreadsheetApp.openById(ROLODEX_ID).getSheetByName(ROLODEX_TAB);
    if (!rolodexSheet) {
      SpreadsheetApp.getUi().alert(`Error: Rolodex tab "${ROLODEX_TAB}" not found in source sheet.`);
      console.error(`${functionName}: Rolodex tab not found.`);
      return;
    }
    
    // Find the row number for the selected municipality in the Rolodex.
    const muniRow = findRowInSheet_(rolodexSheet, muni, ROLODEX_MUNI_COL);

    if (muniRow === -1) {
      SECTIONS.forEach(s => templateSheet.getRange(s.targetA1).clearContent());
      SpreadsheetApp.getUi().alert(`Municipality not found in Rolodex: "${muni}"`);
      console.warn(`${functionName}: Municipality "${muni}" not found. Cleared target cells.`);
      return;
    }
    console.log(`${functionName}: Found "${muni}" at row ${muniRow}.`);

    // Fetch the rich text from the correct row and column.
    const richTextValue = rolodexSheet.getRange(muniRow, ROLODEX_TEXT_COL).getRichTextValue();
    if (!richTextValue || richTextValue.getText().trim() === "") {
      SECTIONS.forEach(s => templateSheet.getRange(s.targetA1).clearContent());
      console.warn(`${functionName}: No rich text found for "${muni}". Cleared target cells.`);
      return;
    }

    // For each section, extract the relevant part of the rich text and paste it.
    console.log(`${functionName}: Extracting and populating sections.`);
    SECTIONS.forEach(sec => {
      console.log(`\nProcessing section: ${sec.name}`);
      const extractedRichText = extractRichTextSection_(richTextValue, sec.start, sec.stop);
      const destinationCell = templateSheet.getRange(sec.targetA1);
      
      if (extractedRichText) {
        destinationCell.setRichTextValue(extractedRichText);
        console.log(`  -> SUCCESS: Populated cell ${sec.targetA1}.`);
      } else {
        destinationCell.clearContent();
        console.log(`  -> INFO: No content found for section. Cleared cell ${sec.targetA1}.`);
      }
    });

    console.log(`--- ${functionName}: Execution finished successfully ---`);
  } catch (err) {
    console.error(`${functionName}: An unexpected error occurred: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`An error occurred. Please check the logs.`);
  }
}


/***** AUTOMATION TRIGGERS *****/

/**
 * Creates an installable "onEdit" trigger to run the script automatically.
 */
function enableMuniAuto() {
  const ss = SpreadsheetApp.getActive();
  disableMuniAuto(); // Clear any old triggers first.
  
  ScriptApp.newTrigger("muniEditHandler")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  console.log("Auto-update trigger enabled.");
  SpreadsheetApp.getUi().alert(`Auto-update enabled. Changing cell ${TEMPLATE_MUNI_A1} will now refresh the instructions.`);
}

/**
 * Deletes the "onEdit" trigger to stop automatic updates.
 */
function disableMuniAuto() {
  let triggerDeleted = false;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "muniEditHandler") {
      ScriptApp.deleteTrigger(t);
      triggerDeleted = true;
    }
  });
  
  if (triggerDeleted) {
    console.log("Auto-update trigger disabled.");
    SpreadsheetApp.getUi().alert("Auto-update disabled.");
  }
}

/**
 * The handler function that the trigger calls on any edit.
 * @param {Object} e The event object passed by the onEdit trigger.
 */
function muniEditHandler(e) {
  try {
    if (!e || !e.range) {
      console.log("muniEditHandler: Event object is missing or malformed.");
      return;
    }
    
    const sheet = e.range.getSheet();
    // Check if the edit happened in the correct cell on the correct sheet
    if (sheet.getName() !== TEMPLATE_SHEET || e.range.getA1Notation() !== TEMPLATE_MUNI_A1) {
      return;
    }

    const newValue = e.value;
    
    // SCENARIO 1: A municipality was selected (the cell has a value)
    if (newValue && newValue.trim() !== '') {
      console.log(`muniEditHandler: Value detected in ${TEMPLATE_MUNI_A1}. Running fillFormattedSections...`);
      fillFormattedSections();
    
    // SCENARIO 2: The municipality was cleared (the cell is empty)
    } else {
      console.log(`muniEditHandler: Municipality cleared from ${TEMPLATE_MUNI_A1}. Clearing instruction fields.`);
      SECTIONS.forEach(sec => {
        sheet.getRange(sec.targetA1).clearContent();
      });
    }

  } catch (err) {
    console.error(`muniEditHandler Error: ${err.message}\n${err.stack}`);
  }
}


/***** HELPER FUNCTIONS *****/

/**
 * Finds the row number of a search term in a specific column of a sheet.
 * @param {Sheet} sheet The sheet to search in.
 * @param {string} searchTerm The text to find.
 * @param {number} colNum The column number to search in.
 * @return {number} The row number (1-based), or -1 if not found.
 */
function findRowInSheet_(sheet, searchTerm, colNum) {
  const values = sheet.getRange(1, colNum, sheet.getLastRow(), 1).getDisplayValues();
  const searchTermUpper = searchTerm.toUpperCase();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].trim().toUpperCase() === searchTermUpper) {
      return i + 1; // Return 1-based row number
    }
  }
  return -1; // Not found
}

/**
 * Extracts a substring of a RichTextValue while preserving formatting.
 * @param {RichTextValue} richText The source rich text.
 * @param {RegExp} startRe The regex for the starting marker.
 * @param {RegExp} stopRe The regex for the ending marker.
 * @return {RichTextValue} The extracted rich text section, or null.
 */
function extractRichTextSection_(richText, startRe, stopRe) {
  const text = richText.getText();
  
  console.log(`  - Searching for start marker: ${startRe}`);
  const startMatch = text.match(startRe);
  console.log(`    -> Match result: ${startMatch ? `"${startMatch[0]}"` : "NULL"}`);
  if (!startMatch) return null;
  
  let startIdx = text.indexOf(startMatch[0]) + startMatch[0].length;

  let endIdx = text.length;
  const afterContent = text.slice(startIdx);
  
  console.log(`  - Searching for stop marker: ${stopRe}`);
  const stopMatch = afterContent.match(stopRe);
  console.log(`    -> Match result: ${stopMatch ? `"${stopMatch[0]}"` : "NULL"}`);
  if (stopMatch) {
    endIdx = startIdx + afterContent.indexOf(stopMatch[0]);
  }
  
  console.log(`    -> Found raw text slice from index ${startIdx} to ${endIdx}.`);

  // Trim leading/trailing whitespace from the slice by adjusting indices
  while (startIdx < endIdx && /\s/.test(text[startIdx])) {
    startIdx++;
  }
  while (endIdx > startIdx && /\s/.test(text[endIdx - 1])) {
    endIdx--;
  }
  
  if (endIdx <= startIdx) {
    console.log('    -> Section is empty after trimming whitespace. Returning null.');
    return null;
  }
  
  console.log(`    -> Final trimmed text slice is from index ${startIdx} to ${endIdx}.`);

  // Build a new rich text value from the extracted range
  const subText = text.substring(startIdx, endIdx);
  const builder = SpreadsheetApp.newRichTextValue().setText(subText);

  richText.getRuns().forEach(run => {
    const runStart = run.getStartIndex();
    const runEnd = run.getEndIndex();
    
    // Find the overlapping range between the run and our desired section
    const overlapStart = Math.max(runStart, startIdx);
    const overlapEnd = Math.min(runEnd, endIdx);

    if (overlapEnd > overlapStart) {
      // Translate the overlap to the new substring's coordinates
      const localStart = overlapStart - startIdx;
      const localEnd = overlapEnd - startIdx;
      
      builder.setTextStyle(localStart, localEnd, run.getTextStyle());
      const link = run.getLinkUrl();
      if (link) {
        builder.setLinkUrl(localStart, localEnd, link);
      }
    }
  });

  return builder.build();
}

