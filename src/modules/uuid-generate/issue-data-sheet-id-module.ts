import { generateUUID } from './uuid-generate-module';
import {
  DEFINE_SHEET_NAME,
  READ_ROW_MARKER,
  UUID_KEY_NAME,
} from '../../constants/constant';
import { getColIndex, getRowIndex } from '../create-table/sheet-module';

export function appendUUIDWithDefineSheet(
  e: GoogleAppsScript.Events.SheetsOnEdit
) {
  const sheet = e.source.getActiveSheet();

  if (!sheet) return;
  if (sheet.getSheetName() !== DEFINE_SHEET_NAME) return;

  const activeRange = sheet.getActiveRange();
  if (!activeRange) return;

  const defineRange = sheet.getRange(
    1,
    1,
    sheet.getLastRow(),
    sheet.getLastColumn()
  );

  const defineSheetValues = defineRange.getValues();
  const headerRowIndex = getRowIndex(defineSheetValues, 1, READ_ROW_MARKER);

  // headerが存在しない
  if (!headerRowIndex)
    throw new Error('Defineシートのヘッダーに無効な編集がされました。');

  const uuidColIndex = getColIndex(
    defineSheetValues[headerRowIndex],
    UUID_KEY_NAME
  );
  const uuidColNumber = uuidColIndex + 1;

  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  // 変更があった行全てのUUIDに更新をかけるか確認していく
  for (let i = startRow; i < startRow + numRows; i++) {
    const uuidRange = sheet.getRange(i, uuidColNumber);
    const targetRow = sheet.getRange(
      i,
      uuidColNumber + 1,
      1,
      sheet.getLastColumn()
    );

    // targetRowの全てのCellが空白の時はuuidRangeを空白に設定する
    if (targetRow.isBlank()) {
      uuidRange.setValue('');
      continue;
    }

    if (!uuidRange.isBlank()) continue;

    // 空白の時だけUUIDを生成
    uuidRange.setValue(generateUUID());
  }
}
