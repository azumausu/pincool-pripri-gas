export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e.source.getActiveSheet();

  if (!sheet) return;

  const addedRow = sheet.getActiveRange()?.getRow();

  if (addedRow === undefined) return;

  sheet.getRange(addedRow, 1).setValue(primaryKey);
}
