// 参照シートの読み込み
import {
  READ_ROW_MARKER,
  REFERENCE_KEY_NAME,
  REFERENCE_VALUE_NAME,
} from './constants/constant';
import { getColIndex, getRowIndex } from './sheet-module';

export function createReferenceMap(
  referenceSheet: GoogleAppsScript.Spreadsheet.Sheet
): Map<string, string> | undefined {
  const referenceMap = new Map<string, string>();
  const referenceRange = referenceSheet.getRange(
    1,
    1,
    referenceSheet.getLastRow(),
    referenceSheet.getLastColumn()
  );
  const referenceSheetValues = referenceRange.getValues();
  const referenceHeaderRowIndex = getRowIndex(
    referenceSheetValues,
    1,
    READ_ROW_MARKER
  );

  // Headerが定義されていない
  if (referenceHeaderRowIndex === undefined) {
    Logger.log(
      `参照シート「${referenceSheet.getSheetName()}」にヘッダーが正しく定義されていません。`
    );
    return;
  }

  const keyColIndex = getColIndex(
    referenceSheetValues[referenceHeaderRowIndex],
    REFERENCE_KEY_NAME
  );
  const valueColIndex = getColIndex(
    referenceSheetValues[referenceHeaderRowIndex],
    REFERENCE_VALUE_NAME
  );

  const dataStartRowNumber = referenceHeaderRowIndex + 1;
  const keyValues = referenceSheet
    .getRange(dataStartRowNumber, keyColIndex + 1, referenceSheet.getLastRow())
    .getValues();
  const valueValues = referenceSheet
    .getRange(
      dataStartRowNumber,
      valueColIndex + 1,
      referenceSheet.getLastRow()
    )
    .getValues();

  for (let i = 0; i < keyValues.length; i++) {
    referenceMap.set(keyValues[i][0], valueValues[i][0]);
  }
  return referenceMap;
}
