import { generateUUID } from './uuid-generate-module';
import { getColIndex, getHeaderRowIndex } from '../create-table/sheet-module';
import {
  DEFINE_SHEET_NAME,
  DEFINE_SHEET_UUID_ROW_OFFSET,
} from '../../constants/define_sheet';
import { UUID_KEY_NAME } from '../../constants/common';

export function appendUUIDWithDefineSheet(
  e: GoogleAppsScript.Events.SheetsOnEdit
) {
  const sheet = e.source.getActiveSheet();

  if (!sheet) return;
  if (sheet.getSheetName() !== DEFINE_SHEET_NAME) return;

  const activeRange = sheet.getActiveRange();
  if (!activeRange) return;

  const headerRowIndex = getHeaderRowIndex(sheet);
  const uuidRowIndex = headerRowIndex + DEFINE_SHEET_UUID_ROW_OFFSET;

  const defineRange = sheet.getRange(
    1,
    1,
    sheet.getLastRow(),
    sheet.getLastColumn()
  );
  const defineSheetValues = defineRange.getValues();
  const uuidColIndex = getColIndex(
    defineSheetValues[uuidRowIndex],
    UUID_KEY_NAME
  );

  const uuidColNumber = uuidColIndex + 1;
  const editStartRow = activeRange.getRow();
  const editNumRows = activeRange.getNumRows();
  // 変更があった行全てのUUIDに更新をかけるか確認していく
  for (let i = editStartRow; i < editStartRow + editNumRows; i++) {
    const uuidRange = sheet.getRange(i, uuidColNumber);
    const editRow = sheet.getRange(
      i,
      uuidColNumber + 1,
      1,
      sheet.getLastColumn()
    );

    // targetRowの全てのCellが空白の時はuuidRangeを空白に設定する
    if (editRow.isBlank()) {
      uuidRange.setValue('');
      continue;
    }

    if (!uuidRange.isBlank()) continue;

    // 空白の時だけUUIDを生成
    uuidRange.setValue(generateUUID());
  }
}
