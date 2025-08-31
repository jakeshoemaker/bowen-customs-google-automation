const CONFIG = {
  SHEET_NAME: 'DONTRUNMaster', // Master will run sheet
  CALENDAR_ID: 'MATS-CALENDAR-ID',
  TITLE_COL: 3, // 0 based index
  INSTALL_START_COL: 21, // 0 based index

  INSTALLS_PER_WEEK_CELL: 'B2',
  CAB_CHASSIS_CELL: 'B11',

  MILLISECONDS_PER_DAY: 1000 * 60 * 60 * 24
};

// Sheet == Config
// InstallsPerWeek == B1 [1, 0]
// CabChassis Cell == B11 [1, 10]
// If BuiltType == CabChassis, Build takes 1.5 weeks
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e.range.getSheet();
  if (
    sheet.getName() !== CONFIG.SHEET_NAME &&
    e.range.getColumn() !== CONFIG.INSTALL_START_COL
  ) {
    logToDocument(`Skipping Edit \n`);
    return;
  }

  const conf = getSheetConfig();
  if (!conf) return;

  const installsPerWeek = conf.getRange(CONFIG.INSTALLS_PER_WEEK_CELL).getValue();
  const cabChassisBuildType = conf.getRange(CONFIG.CAB_CHASSIS_CELL).getBackground();


  const dr = sheet.getDataRange();
  const sheetData = dr.getValues();
  const backgrounds = dr.getBackgrounds();

  const logs = [];
  const updates = [];
  const editedRow = e.range.getRow();
  for (let i = editedRow - 1; i < sheetData.length; i++) {
    const anchorRowIndex = i - installsPerWeek;
    if (anchorRowIndex < 0) continue;

    // Determine clearance time based the build is a cab chassis build
    const anchorInstallDate = validateDate(sheetData[anchorRowIndex][CONFIG.INSTALL_START_COL]);

    // const prevRow = sheetData[i - 1];
    // const currentRow = sheetData[i];
    //
    // const prevInstallStart = validateDate(prevRow[CONFIG.INSTALL_START_COL]);
    // const currentInstallStart = validateDate(currentRow[CONFIG.INSTALL_START_COL]);
    //
    // if (!prevInstallStart || !currentInstallStart) break;
    //
    // if (hasCollision(prevInstallStart, currentInstallStart)) {
    //   const shiftedInstallDate = new Date(prevInstallStart.getTime() + oneWeek);
    //   sheetData[i][CONFIG.INSTALL_START_COL] = shiftedInstallDate;
    //   logs.push(`Row ${i + 1} shifted to ${shiftedInstallDate.toDateString()}`);
    // }
  }

  // Write updated data back to the sheet
  sheet
    .getRange(1, 1, sheetData.length, sheetData[0].length)
    .setValues(sheetData);

  logs.push(['Done\n']);
  logToDocument(logs.join('\n')); // Log the resulting operation
}

/**
 * Recalculates the Install Start date for rows by shifting the start date
 * to align with the build type, as well as, the current production capacity
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet the spreadsheet to process
 * @param {number} fromRow the starting row to process (defaults to 1)
 */
function recalculateBuildDates(sheet, fromRow = 1) {
  const data = sheet.getDataRange().getValues();
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  let hasChanges = false;

  // loop thru each row, starting from the
  for (let i = fromRow; i < data.length; i++) {
    const prevRow = data[i - 1];
    const currentRow = data[i];

    const prevDate = new Date(prevRow[CONFIG.INSTALL_START_COL]);
    const currentDate = new Date(currentRow[CONFIG.INSTALL_START_COL]);
    const originalDate = new Date(currentDate);

    // check to make sure dates are valid before processing
    if (isNaN(prevDate.getTime()) || isNaN(currentDate.getTime()))
      return;

    // COLLISION CHECK: Does the current row's week fall on or before the previous row's week?
    if (getWeekNumber(currentDate) <= getWeekNumber(prevDate)) {
      // get build-type of PREV row, which is causing the collision
      const collidingBuildType = 
    }
  }
}

/**
 * Logs a message to the Document with the current date and time.
 * @param message The message to log.
 */
function logToDocument(message: string) {
  const docName = 'bowen-appscript-debug-log';
  const timestamp = new Date().toISOString();

  const files = DriveApp.getFilesByName(docName);

  const doc = files.hasNext()
    ? DocumentApp.openById(files.next().getId())
    : DocumentApp.create(docName);


  doc.getBody().appendParagraph(`[${timestamp}] ${message}`)
  doc.saveAndClose();
}

/**
 * Checks to make sure the cell value is a valid date object
 */
function validateDate(value: unknown): Date | null {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  return null;
}

/**
 * Checks if the current install date is within the timeframe it takes
 * to finish building the prevDate
 */
function hasCollision(
  prevDate: Date,
  currentDate: Date,
  cabChassisBuildType: boolean = false): boolean {
  const diffDays = (
    currentDate.getTime() - prevDate.getTime()
  ) / CONFIG.MILLISECONDS_PER_DAY;

  return cabChassisBuildType
    ? diffDays < 14
    : diffDays < 7; // CabChassis builds take 1.5 weeks
}

/**
 * Checks if the current row is a CabChassis Build
 */
function isCabChassisBuild(row, cabChassisRowColor) {

}

/**
 * Fetches the config sheet
 */
function getSheetConfig() {
  const configSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config");

  if (!configSheet) {
    logToDocument('Config sheet not found \n');
    return;
  }

  return configSheet
}
