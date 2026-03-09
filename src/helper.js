
/*** HELPER FUNCTIONS ***/


/**
 * Shows an informational dialog with an "OK" button.
 *
 * @param {string} title - The title of the dialog box.
 * @param {string} body - The body message displayed in the dialog.
 */
function okButton(title, body) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, body, ui.ButtonSet.OK);
}


/**
 * Shows a confirmation dialog in Google Sheets UI and returns whether user confirmed.
 *
 * @param {string} title - The title of the dialog box.
 * @param {string} body - The body message displayed in the dialog.
 * @param {GoogleAppsScript.Base.ButtonSet} [buttonSet=ui.ButtonSet.YES_NO] - 
 *        (Optional) The button set to use, defaults to Yes/No.
 * @return {boolean} True if the user clicked "Yes", false otherwise.
 */
function confirmButtonPress(title, body) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(title, body, ui.ButtonSet.YES_NO);
  return response === ui.Button.YES;
}


/**
 * getStudioSheet
 *  Get a sheet reference for the Studio sheet
 *
 * @param {string} spreadSheetName - Name of the spreadsheet (unused, kept for consistency)
 * @param {string} sheetName - Name of the sheet (tab) to retrieve
 * @returns {Sheet} Sheet reference from the active spreadsheet
 */
function getStudioSheet(spreadSheetName, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * getWardSheet
 * Get a sheet reference for the Ward sheet
 *
 * @param {string} spreadSheetName - Name of the spreadsheet (unused, kept for consistency)
 * @param {string} sheetName - Name of the sheet (tab) to retrieve
 * @returns {Sheet} Sheet reference from the active spreadsheet
 */
function getWardSheet(spreadSheetName, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * getDataSheet
 *  Get a sheet reference for the Data sheet
 *
 * @param {string} spreadSheetName - Name of the spreadsheet file
 * @param {string} sheetName - Name of the sheet (tab) to retrieve
 * @returns {Sheet} Sheet reference from the spreadsheet
 */
function getDataSheet(spreadSheetName, sheetName) {
  return getCwdSpreadsheet(spreadSheetName).getSheetByName(sheetName);
}

/**
 * getDataBackupSheet
 *  Get a sheet reference for the Data (backup) sheet
 *
 * @param {string} spreadSheetName - Name of the spreadsheet file
 * @param {string} sheetName - Name of the sheet (tab) to retrieve
 * @param {string} backupFolderName - Name of the subfolder under the current file’s parent
 * @returns {Sheet} Sheet reference from the spreadsheet in the backup folder
 */
function getDataBackupSheet(spreadSheetName, sheetName, backupFolderName) {
  return getCwdSpreadsheet(spreadSheetName, backupFolderName).getSheetByName(sheetName);
}


/**
 * getCwdSpreadsheet
 * Find a Google Sheet near the current spreadsheet.
 * - Looks in each parent folder of the current file.
 * - Optionally first steps into a named subfolder.
 * - Picks the newest file if there are duplicates.
 *
 * @param {string} fileName - Name of the spreadsheet file to open
 * @param {string|null} subFolder - Optional subfolder name under the parent(s)
 * @returns {Spreadsheet}
 */
function getCwdSpreadsheet(fileName, subFolder=null) {
  const thisFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // Get the parent folder(s) of the current file
  const parents = thisFile.getParents();
  if (!parents.hasNext()) {
    throw new Error("This file has no parent folder (maybe in Shared Drive root).");
  }

  // First Folder in parent iterator
  let parent = parents.next();

  // Find the sub-folder if subFolder !== null
  if (subFolder !== null) {
    const subfolders = parent.getFoldersByName(subFolder);
    if (!subfolders.hasNext()) {
      throw new Error(`No folder named ${subFolder} found in the current folder.`);
    }
    parent = subfolders.next();
  }

  // Search in parent folder for fileName
  const files = parent.getFilesByName(fileName);
  if (!files.hasNext()) {
    throw new Error(`No file named ${fileName} found.`);
  }
  // First file in file iterator
  const file = files.next();

  // Open it as a Spreadsheet
  return SpreadsheetApp.openById(file.getId());
}


/**
 * displayMsg
 * Display message received by the previous shift
 * 
 * @param {Sheet} dataSheet - Studio data sheet reference
 * @param {Sheet} studioSheet - Studio sheet reference
 * @param {Object} studioCellPos - Stores all cell positions, incl. ranges, for studio form input elements
 *    (e.g. {  showAired: 'E12', songRequests: 'O8:O31', studioVisits: 'I9:M31' ...})
 */
function displayMsg(dataSheet, studioSheet, studioCellPos) {

  // Cell to display prev shift message
  const msgCell = studioSheet.getRange(studioCellPos.displayMsgCell);
  const leftBorder = studioSheet.getRange(studioCellPos.displayMsgCellLeftBorder);
  const topBorder = studioSheet.getRange(studioCellPos.displayMsgCellTopBorder);
  const cornerBorder = studioSheet.getRange(studioCellPos.displayMsgCellCornerBorder);

  msgCell.clear()
  leftBorder.clear()
  topBorder.clear()
  cornerBorder.clear()

  const msg = prevShiftMsg(dataSheet);
  if (msg === null) return;

  const dateStr = msg.date
    ? Utilities.formatDate(msg.date, "Australia/Melbourne", "dd/MM/yyyy")
    : "(unknown date)";

  const header = `\nMessage from ${dateStr} shift!\n\n`;
  const fullText = header + String(msg.text) + '\n';

  // Build styles
  const styleHeader = SpreadsheetApp.newTextStyle().setFontSize(14).build();
  const styleBody   = SpreadsheetApp.newTextStyle().setFontSize(10).build();

  // Build rich text
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(1, header.length, styleHeader)
    .setTextStyle(header.length, fullText.length, styleBody)
    .build();

  msgCell.setRichTextValue(rich);

  // Add coloured cells to make message look pretty
  msgCell.setBackground("#FEE9A8");
  leftBorder.setBackground("#F9D258");
  topBorder.setBackground("#F9D258");
  cornerBorder.setBackground("#E9BA2A");
}


/**
 * prevShiftMsg
 * Get message left by previous shift
 *
 * @param {Sheet} dataSheet - Studio data sheet reference
 * @return {Object} Object like {date: last_shift_date, text: msg_from_last_shift}
 */
function prevShiftMsg(dataSheet) {
  const lastRow = dataSheet.getLastRow();  
  const numRows = 10;                       
  const startRow = Math.max(lastRow - numRows + 1, 1);

  // Get dates / shiftMsgs only from the last 10 rows for compute efficiency
  const dates = dataSheet.getRange(startRow, 1, numRows, 1).getValues().flat();
  const shiftMsgs = dataSheet.getRange(startRow, 21, numRows, 1).getValues().flat();

  let maxDate = null;
  let maxIndex = -1;

  const today = new Date();
  today.setHours(0, 0, 0, 0); // today's date at 00:00

  // Get index of max date that is also less than current date
  dates.forEach((v, i) => {
    if (v instanceof Date && !isNaN(v)) {
      const d = new Date(v);
      d.setHours(0, 0, 0, 0); // normalize candidate date

      if (d < today && (maxDate === null || d > maxDate)) {
        maxDate = d;
        maxIndex = i;
      }
    }
  });

  // Return message if message from last shift is non-empty and date is valid
  if (maxIndex !== -1 && String(shiftMsgs[maxIndex] ?? '').trim() !== '') {
    return { date: dates[maxIndex], text: String(shiftMsgs[maxIndex] ?? '').trim() };
  } else {
    return null;
  }
}


/**
 * buildRowData
 * Get input data from studio & ward forms and build row
 * ready to append to studio data form
 *
 * @param {Sheet} studioSheet - Studio sheet reference
 * @param {Sheet} wardSheet - Ward sheet reference
 * @param {Object} studioCellPos - Stores all cell positions, incl. ranges, for studio form input elements
 *    (e.g. {  showAired: 'E12', songRequests: 'O8:O31', studioVisits: 'I9:M31' ...})
 * @param {Object} wardCellPos - Stores all cell positions, incl. ranges, for ward form input elements
 *    (e.g. {  wardVisitsNonBool: 'H9:L38', wardVisitsBool: 'M9:M38', ward2RoomVisits: 'E11'  ...})
 * @return {Array<string|integer|boolean>} Array which will be appended as a row to studio data sheet
 */
function buildRowData(studioSheet, wardSheet, studioCellPos, wardCellPos) {
  // Column indexes for bool fields in ward / studio tables
  const studioVisitsGiftsIdx = 3;
  const studioVisitsShowIdx = 4;
  const wardGiftsIdx = 5;

  // Get data from studio & ward forms
  const studioFormData = getStudioFormData(studioSheet, studioCellPos);
  const wardData = getWardFormData(wardSheet, wardCellPos);

  // Build row
  return [
    studioFormData.date,                                                      // Date
    new Set([                                                                 // Volunteer Count
      ...studioFormData.volsOnShift, 
      ...studioFormData.volsMakeup, 
      ...studioFormData.volsTeamLeads, 
      ...studioFormData.volsDJs
    ]).size,
    studioFormData.studioVisits.length,                                       // Studio Visits Count
    wardData.wardVisits.length,                                               // Ward Visits Count
    wardData.uniqueWardsVisited.join(', '),                                   // Wards Visited
    boolColTrueCount(studioFormData.studioVisits, studioVisitsGiftsIdx)       // Gifts Given Count
      + boolColTrueCount(wardData.wardVisits, wardGiftsIdx),
    boolColTrueCount(studioFormData.studioVisits, studioVisitsShowIdx),       // Radio Show Guests Count
    studioFormData.songRequests.length,                                       // Song Requests Count
    studioFormData.radioWorks,                                                // Radio Works
    studioFormData.sustServiceWorks,                                          // Sustaining Service Works
    studioFormData.showAired,                                                 // Radio Show Aired
    studioFormData.volsTeamLeads.join(', '),                                  // Team leads List
    studioFormData.volsDJs.join(', '),                                        // DJs List
    studioFormData.volsOnShift.join(', '),                                    // Volunteer List
    studioFormData.volsMakeup.join(', '),                                     // Makeup Volunteer List
    studioFormData.songRequests.join(', '),                                   // Song Requests
    studioFormData.studioVisitsStr,                                           // Studio Visits
    wardData.wardVisitsStr,                                                   // Ward Visits
    studioFormData.addNotes,                                                  // Additional Notes
    studioFormData.wrapUp,                                                    // Wrap Up
    studioFormData.nextShiftMsg,                                              // Message for Next Shift
    Utilities.formatDate(new Date(), "Australia/Melbourne", "dd/MM/yyyy HH:mm:ss")  // Submission Datetime
  ];

}


/**
 * clearStudioForm
 * Clear inputs in studio form for re-use
 *
 * @param {Sheet} studioSheet - Studio sheet reference
 * @param {Object} studioCellPos - Stores all cell positions, incl. ranges, for studio form input elements
 *    (e.g. {  showAired: 'E12', songRequests: 'O8:O31', studioVisits: 'I9:M31' ...})
 */
function clearStudioForm(studioSheet, studioCellPos) {
  
  // Reset date input to todays date
  studioSheet.getRange(studioCellPos.date).setFormula("=TODAY()");

  // For all input cells in studio form...
  for (const [field, cellPos] of Object.entries(studioCellPos)) {
    // Checkbox (bool) inputs to set to unchecked (false)
    if (["radioWorks", "sustServiceWorks", "showAired", "studioVisitsBool"].includes(field)) {
      studioSheet.getRange(cellPos).setValue(false);
    }

    // Text inputs to clear content
    else if (['volsTeamLeads', 'volsDJs', 'volsOnShift', 'volsMakeup', 'addNotes', 'wrapUp', 'nextShiftMsg', 'studioVisitsNonBool', 'songRequests', 'songRequestsNames'].includes(field)) {
      studioSheet.getRange(cellPos).clearContent();
    }
  }
}


/**
 * clearWardForm
 * Clear inputs in ward form for re-use
 *
 * @param {Sheet} wardSheet - Ward sheet reference
 * @param {Object} wardCellPos - Stores all cell positions, incl. ranges, for ward form input elements
 *    (e.g. {  wardVisitsNonBool: 'H9:L38', wardVisitsBool: 'M9:M38', ward2RoomVisits: 'E11'  ...})
 */
function clearWardForm(wardSheet, wardCellPos) {

  // For all input cells in ward form...
  for (const [field, cellPos] of Object.entries(wardCellPos)) {
    // Checkbox (bool) inputs to set to unchecked (false)
    if (['wardVisitsBool'].includes(field)) {
      wardSheet.getRange(cellPos).setValue(false);
    }

    // Text inputs to clear content
    else {
      wardSheet.getRange(cellPos).clearContent();
    }
  }
}


/**
 * boolColTrueCount
 * For some generic table, get the number of checked (true) cells in a checkbox (bool) field
 *
 * @param {Array<Array<string|boolean>>} table - 2D array representing a generic table
 * @param {integer} colIndex - Column index where bool field exists
 * @return {integer} Number of true values at col: colIndex in table
 */
function boolColTrueCount(table, colIndex) {
  return table
    .map(row => row[colIndex])
    .filter(row => row)
    .length;
}

