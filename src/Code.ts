const CONFIG = {
  SHEET_NAME: 'Master',
  CALENDAR_ID: 'MATS-CALENDAR-ID',
  TITLE_COL: 3,
  INSTALL_START_COL: 21,
  PRODUCTION_CAPACITY: 30,
  BUILD_TYPE_COL: 31,

  BUILD_SHIFTS: {
    'Complex': 2,
    'Standard': 1,
    'Simple': 1
  }
};

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e.range.getSheet();
  if (
    sheet.getName() !== CONFIG.SHEET_NAME &&
    e.range.getColumn() !== CONFIG.INSTALL_START_COL
  ) {
    return;
  }

  const editedValue = e.range.getValue();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  Logger.log(
    `Edit triggered:\n Col: ${editedCol}\n Row: ${editedRow}\n Value: ${editedValue}`
  );

  const sheetData = sheet.getDataRange().getValues();
  for (let i = editedRow; i < sheetData.length; i++) {
    const row = sheetData[i];
    const customer = row[0];
    const vehicle = row[1];
    const installStart = row[CONFIG.INSTALL_START_COL]; // FIXED index
    const valid = validateDate(installStart);

    if (!valid) {
      logToDocument(`Row ${i + 1}: Invalid date`);
    } else {
      logToDocument(
        `Row ${i + 1}: Customer: ${customer}, Vehicle: ${vehicle}, Install Start: ${installStart}`
      );
    }
  }
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
