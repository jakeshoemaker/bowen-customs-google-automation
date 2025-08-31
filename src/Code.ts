/**
 * Configuration: Adjust these settings to match your sheet.
 */
const CONFIG = {
  SHEET_NAME: 'Master',
  CONFIG_SHEET_NAME: 'Config',
  CALENDAR_ID: 'MATS_CALENDAR_ID', // <-- IMPORTANT: Change this!
  TITLE_COL: 4,                    // Column A: Customer Build Title
  INSTALL_START_COL: 22,           // Column V: Install Start date (the trigger column)
  PREVENT_CALENDAR_MGMT: true,     // A helper to prevent calendar management while testing

  // Define the required buffer in weeks for each build type.
  // The keys here MUST match the values in Column B of your 'Config' sheet.
  BUILD_SHIFTS: {
    'CabChassis': 2,
    'Standard': 1
  }
};

/**
 * Main trigger. Runs on any edit and calls the schedule recalculation.
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const { range } = e;
  const sheet = range.getSheet();
  const editedRow = range.getRow();

  // 1. EXIT EARLY: Ignore edits on the wrong sheet, wrong column, or header row.
  if (
    sheet.getName() !== CONFIG.SHEET_NAME ||
    range.getColumn() !== CONFIG.INSTALL_START_COL ||
    editedRow === 1
  ) {
    return;
  }

  // 2. RECALCULATE THE SCHEDULE
  recalculateSchedule(sheet, editedRow);
}

/**
 * Recalculates the schedule from the edited row downwards, respecting
 * weekly capacity and consecutive build buffers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to process.
 * @param {number} startRow The row number where the edit occurred.
 */
function recalculateSchedule(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRow: number
) {
  // 1. LOAD CONFIG & DATA
  const appConfig = loadAppConfig();
  if (!appConfig) return; // Stop if config is invalid

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  let hasChanges = false;

  // 2. PRE-FILL CAPACITY: Count installs in weeks *before* the change.
  const weeklyInstallCounts = new Map();
  for (let i = 1; i < startRow - 1; i++) { // Loop up to the row *before* the startRow
    const date = new Date(data[i][CONFIG.INSTALL_START_COL - 1]);
    if (!isNaN(date.getTime())) {
      const week = getWeekNumber(date);
      weeklyInstallCounts.set(week, (weeklyInstallCounts.get(week) || 0) + 1);
    }
  }

  // 3. PROCESS FROM EDITED ROW DOWNWARDS
  for (let i = startRow - 1; i < data.length; i++) {
    const originalDate = new Date(data[i][CONFIG.INSTALL_START_COL - 1]);
    let proposedDate = new Date(originalDate);
    if (isNaN(proposedDate.getTime())) continue;

    // A. RESPECT BUFFER FROM PREVIOUS ROW
    const prevFinalDate = new Date(
      data[i - 1][CONFIG.INSTALL_START_COL - 1]
    );
    if (!isNaN(prevFinalDate.getTime())) {
      const blockerColor = backgrounds[i - 1][CONFIG.INSTALL_START_COL - 1];
      const blockerBuildType =
        (blockerColor === appConfig.cabChassisColor)
          ? 'CabChassis'
          : 'Standard';
      const weeksToShift = CONFIG.BUILD_SHIFTS[blockerBuildType] || 1;

      const minAllowedDate = new Date(prevFinalDate);
      minAllowedDate.setDate(minAllowedDate.getDate() + weeksToShift * 7);

      // If proposedDate is before the minAllowedDate, set it as new minAllowedDate
      if (proposedDate < minAllowedDate) {
        proposedDate = minAllowedDate;
      }
    }

    // B. RESPECT WEEKLY CAPACITY
    let week = getWeekNumber(proposedDate);
    while ((weeklyInstallCounts.get(week) || 0) >= appConfig.installsPerWeek) {
      proposedDate.setDate(proposedDate.getDate() + 7); // Push to the next week
      week = getWeekNumber(proposedDate);
    }

    // C. FINALIZE & UPDATE
    weeklyInstallCounts.set(week, (weeklyInstallCounts.get(week) || 0) + 1);

    if (proposedDate.getTime() !== originalDate.getTime()) {
      data[i][CONFIG.INSTALL_START_COL - 1] = proposedDate;
      hasChanges = true;
      const eventTitle = data[i][CONFIG.TITLE_COL - 1];
      updateCalendarEvent(cal, eventTitle, originalDate, proposedDate);
    }
  }

  // 4. WRITE BACK if changes were made
  if (hasChanges) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    console.log("Schedule recalculated from row " + startRow);
  }
}

/**
 * Loads critical settings from the 'Config' sheet.
 * @returns {object|null} An object with config values or null if sheet is missing.
 */
function loadAppConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!configSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Config' sheet not found!");
    return null;
  }
  return {
    installsPerWeek: configSheet.getRange("B1").getValue(),
    cabChassisColor: configSheet.getRange("B2").getBackground()
  };
}

// Helper function getWeekNumber() remains the same
function getWeekNumber(d: Date) {
  // create a clean, timezone neutral copy of the date
  const date: any = new Date(
    Date.UTC(
      d.getFullYear(),
      d.getMonth(),
      d.getDate()
    ));

  // Find thursday of that week. (ISO Standard revolves around thrusdays)
  const dayOfWeek = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayOfWeek);

  // find first day of that year
  const year = date.getUTCFullYear();
  const firstDayOfYear: any = new Date(Date.UTC(year, 0, 1));

  // calculate difference in days, divide by 7, round up, giving us the week number
  const millisecondsPerDay = 86400000;
  const daysBetween = (date - firstDayOfYear) / millisecondsPerDay;
  const weekNo = Math.ceil((daysBetween + 1) / 7);

  // combine year + weekNo to give unique numbers
  return year * 100 + weekNo;
}

// Helper function updateCalendarEvent() remains the same
function updateCalendarEvent(
  calendar: GoogleAppsScript.Calendar.Calendar,
  eventTitle: string,
  oldDate: Date,
  newDate: Date) {
  if (!CONFIG.PREVENT_CALENDAR_MGMT) {
    logToDocument("Skipping calendar update");
    return;
  }

  try {
    const events = calendar.getEventsForDay(oldDate);
    const targetEvent = events.find(event => event.getTitle().includes(eventTitle));
    if (targetEvent) {
      const startTimeMs = targetEvent.getStartTime().getMilliseconds()
      const endTimeMs = targetEvent.getEndTime().getMilliseconds();
      const duration = endTimeMs - startTimeMs;

      const newEndDate = new Date(newDate.getTime() + duration);
      targetEvent.setTime(newDate, newEndDate);
      console.log(`Calendar updated for: ${eventTitle}`);
    }
  } catch (error) {
    console.error(`Failed to update calendar for ${eventTitle}:`, error);
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
