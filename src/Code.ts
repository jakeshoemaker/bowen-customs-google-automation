const CONFIG = {
  SHEET_NAME: 'DONTRUNMaster',
  CALENDAR_ID: 'MATS-CALENDAR-ID',
  TITLE_COL: 3,
  INSTALL_START_COL: 21,
  PRODUCTION_CAPACITY: 30,
  BUILD_TYPE_COL: 31,

  BUILD_SHIFTS: {
    'Complex': 2,
    'Standard': 1,
    'Simple': 1
  },
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

  const editedValue = e.range.getValue();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  Logger.log(
    `Edit triggered:\n Col: ${editedCol}\n Row: ${editedRow}\n Value: ${editedValue}`
  );

  const logs = [];
  const oneWeek = CONFIG.MILLISECONDS_PER_DAY * 7;
  const sheetData = sheet.getDataRange().getValues();

  for (let i = editedRow; i < sheetData.length; i++) {
    const prevRow = sheetData[i - 1];
    const currentRow = sheetData[i];

    const prevInstallStart = validateDate(prevRow[CONFIG.INSTALL_START_COL]);
    const currentInstallStart = validateDate(currentRow[CONFIG.INSTALL_START_COL]);

    if (!prevInstallStart || !currentInstallStart) continue;

    if (hasCollision(prevInstallStart, currentInstallStart)) {
      const shiftedInstallDate = new Date(prevInstallStart.getTime() + oneWeek);
      sheetData[i][CONFIG.INSTALL_START_COL] = shiftedInstallDate;
      logs.push(`Row ${i + 1} shifted to ${shiftedInstallDate.toDateString()}`);
    }
  }

  // Write updated data back to the sheet
  sheet
    .getRange(1, 1, sheetData.length, sheetData[0].length)
    .setValues(sheetData);

  logs.push(['Done\n']);
  logToDocument(logs.join('\n')); // Log the resulting operation
}

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

function validateDate(value: unknown): Date | null {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  return null;
}

function hasCollision(prevDate: Date, currentDate: Date): boolean {
  const diffDays = (
    currentDate.getTime() - prevDate.getTime()
  ) / (1000 * 60 * 60 * 24);

  return diffDays < 7; // Less than 1 week apart
}
