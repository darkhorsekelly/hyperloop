/**
 * @OnlyCurrentDoc
 * This script adds a timestamp to a target cell when a corresponding
 * dropdown cell is changed to "Yes".
 */

// --- CONFIGURATION ---
// This object maps the dropdown cells (keys) to their timestamp cells (values).
const TIMESTAMP_CONFIG = {
  // "Dropdown A1 Notation": "Timestamp A1 Notation"
  "A1": "A2"
  // etc.
};

const TIMEZONE = "America/New_York"; // Eastern Timezone
const TIMESTAMP_FORMAT = "M/d/yyyy HH:mm:ss"; // e.g., 9/10/2025 14:26:31

/**
 * Runs automatically when a user edits the spreadsheet.
 * @param {Object} e The event object.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedCellA1 = range.getA1Notation();
  const newValue = e.value;
  const oldValue = e.oldValue;

  // 1. Check if the edited cell is one of our configured dropdowns.
  // If not, exit the script immediately.
  const targetCellA1 = TIMESTAMP_CONFIG[editedCellA1];
  if (!targetCellA1) {
    console.log(`Edit in cell ${editedCellA1} is not configured for timestamping. Exiting.`);
    return;
  }

  const targetCell = sheet.getRange(targetCellA1);

  // 2. Logic for handling the change.
  // If changed TO "Yes" (from "No", blank, or even "Yes" again).
  if (newValue === 'Yes') {
    const timestamp = Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT);
    console.log(`Setting timestamp "${timestamp}" in cell ${targetCellA1} due to change in ${editedCellA1}.`);
    targetCell.setValue(timestamp);
  
  // If changed FROM "Yes" TO "No".
  } else if (oldValue === 'Yes' && newValue === 'No') {
    console.log(`Clearing timestamp in cell ${targetCellA1} due to change in ${editedCellA1} from "Yes" to "No".`);
    targetCell.clearContent();
  }
}
