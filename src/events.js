/*** CONSTANTS ***/

const DATA_SPREADSHEET = 'Studio Stats & Data';
const DATA_PRIV_SHEET = 'PrivateData';

const DATA_BACKUP_SPREADSHEET = 'Studio Stats & Data - Backup';
const DATA_BACKUP_PRIV_SHEET = 'PrivateData';
const BACKUP_FOLDER = 'Backup';

const STUDIO_SPREADSHEET = 'Studio Forms';
const STUDIO_SHEET = 'Studio - Form (Main)';
const WARD_SHEET = 'Wards - Form';
const VOLUNTEERS_LIST_SHEET = 'Volunteers List';

/*** Customise your ranges here ***/
const WARD_CELL_POS = {
  wardVisits: 'H9:M48',
  wardVisitsNonBool: 'H9:L48',
  wardVisitsBool: 'M9:M48',
  nursesStationInputs: 'E11:E46'
};
const STUDIO_CELL_POS = {
  date: 'C10',
  volsTeamLeads: 'C14',
  volsDJs: 'C22',
  volsOnShift: 'C30',
  volsMakeup: 'C38',
  addNotes: 'E14',
  wrapUp: 'E30',
  nextShiftMsg: 'E22',
  radioWorks: 'G14',
  sustServiceWorks: 'G15',
  showAired: 'G16',

  studioVisits: 'K9:O43',
  studioVisitsNonBool: 'K9:M43',
  studioVisitsBool: 'N9:O43',

  songRequests: 'R9:R43',
  songRequestsNames: 'Q9:Q43',

  displayMsgCell: 'C5',
  displayMsgCellLeftBorder: 'B5',
  displayMsgCellTopBorder: 'C4',
  displayMsgCellCornerBorder: 'B4'
};


/*** EVENT FUNCTIONS ***/


/**
 * triggerDisplayMsg
 * Func executes when sheets page is loaded in browser and on 5 min timer. Display previous shift message on studio sheet
 */
function triggerDisplayMsg() {
  const studioSheet = getStudioSheet(STUDIO_SPREADSHEET, STUDIO_SHEET);
  const dataSheet = getDataSheet(DATA_SPREADSHEET, DATA_PRIV_SHEET);
  displayMsg(dataSheet, studioSheet, STUDIO_CELL_POS);
}


/**
 * safeTriggerDisplayMsg
 * Wrapper for triggerDisplayMsg to suppress email notifications on error.
 */
function safeTriggerDisplayMsg() {
  try {
    triggerDisplayMsg();
  } catch (err) {
    console.log('triggerDisplayMsg failed: ' + err);
  }
}


/**
 * passMsgToNextShift
 * Func executes when 'Pass Message' button is clicked in studio form. Populates nextMsgCell with previous shifts message
 */
function passMsgToNextShift() {
  const studioSheet = getStudioSheet(STUDIO_SPREADSHEET, STUDIO_SHEET);
  const dataSheet = getDataSheet(DATA_SPREADSHEET, DATA_PRIV_SHEET);

  const msg = prevShiftMsg(dataSheet);
  if (msg === null || msg.text.trim() === "") {
    okButton("No Message Available", "The previous shift did not leave a message. You can enter a new message.");
    return;
  }
  const nextMsgCell = studioSheet.getRange(STUDIO_CELL_POS.nextShiftMsg);

  nextMsgCell.setValue(msg.text.trim());
}


/**
 * clearAllForms
 * Func executes when 'Clear' button is clicked in studio form. Clears all inputs in Studio & Ward forms
 */
function clearAllForms() {
  if (!confirmButtonPress('Clear Studio & Ward Forms', 'Clear studio & ward data now? This action cannot be undone.')) return;

  const studioSheet = getStudioSheet(STUDIO_SPREADSHEET, STUDIO_SHEET);
  const wardSheet = getWardSheet(STUDIO_SPREADSHEET, WARD_SHEET);
  clearStudioForm(studioSheet, STUDIO_CELL_POS);
  clearWardForm(wardSheet, WARD_CELL_POS);
}


/**
 * submitData
 * Func executes when 'Submit' button is clicked in studio form. Appends studio & ward form inputs to studio data sheet
 */
function submitData() {
  if (!confirmButtonPress(
    "Submit Studio & Ward Forms", 
    "Submit studio & ward data now? This action cannot be undone.\n\nIf you want to forward the previous shift’s message, please use the \"Pass on Previous Shift's Message\" button before continuing."
  )) return;

  const studioSheet = getStudioSheet(STUDIO_SPREADSHEET, STUDIO_SHEET);
  const wardSheet = getWardSheet(STUDIO_SPREADSHEET, WARD_SHEET);
  const dataSheet = getDataSheet(DATA_SPREADSHEET, DATA_PRIV_SHEET);
  const dataBackupSheet = getDataBackupSheet(DATA_BACKUP_SPREADSHEET, DATA_BACKUP_PRIV_SHEET, BACKUP_FOLDER)

  const rowData = buildRowData(studioSheet, wardSheet, STUDIO_CELL_POS, WARD_CELL_POS);
  dataSheet.appendRow(rowData);
  dataBackupSheet.appendRow(rowData);

  // Clear studio and ward forms after submission
  clearStudioForm(studioSheet, STUDIO_CELL_POS);
  clearWardForm(wardSheet, WARD_CELL_POS);

  /*
  TODO: Clear previous shifts msg at the top
  - Run displayMsg
  - Edit line 153 in helper.gs to not display msg (return early) if todays shift is submitted
  */
}


/**
 * onEdit
 * Func executes when any cell is edited. Checks if volunteers list is edited so values can be trimmed
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const volsListCol = 2;
  const volsListStartRow = 7;

  // If user edited volunteer list (in volunteers list sheet)
  if (sheet.getName() === VOLUNTEERS_LIST_SHEET && col === volsListCol && e.range.getRow() >= volsListStartRow) {
    // Trim all names in volunteer list
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - volsListStartRow + 1;
    const range = sheet.getRange(volsListStartRow, col, numRows, 1);
    const values = range.getValues();
    const trimmed = values
      .map(r => [String(r[0]).trim()]);
      //.sort((a, b) => a[0].localeCompare(b[0]));
    range.setValues(trimmed);
  }
}

