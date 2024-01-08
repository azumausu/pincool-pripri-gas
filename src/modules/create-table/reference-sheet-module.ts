// 参照シートの読み込み
import {
  createPullDown,
  createReferenceSwitchFormula,
  getCellName,
  getColIndex,
  getHeaderRowIndex,
  insertDataSheetHeader,
  moveColumnData,
} from './sheet-module';
import {
  CELL_NAME,
  REFERENCE_KEY_NAME,
  REFERENCE_VALUE_NAME,
} from '../../constants/common';
import { DATA_SHEET_START_ROW_OFFSET } from '../../constants/data_sheet';

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
  const referenceHeaderRowIndex = getHeaderRowIndex(referenceSheet);

  const keyColIndex = getColIndex(
    referenceSheetValues[referenceHeaderRowIndex],
    REFERENCE_KEY_NAME
  );
  const valueColIndex = getColIndex(
    referenceSheetValues[referenceHeaderRowIndex],
    REFERENCE_VALUE_NAME
  );

  const headerNumber = referenceHeaderRowIndex + 1;
  const dataStartRowNumber = headerNumber + 1;
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

// 参照シートを挿入する
export function insertDataSheetReferenceColumn(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  referenceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  uuid: string,
  variableName: string,
  displayName: string,
  insertDataColNumber: number,
  dataSheetHeaderRowNumber: number,
  dataSheetUUIDToColMap: Map<string, number>
) {
  const referenceInsertDataColNumber = insertDataColNumber + 1;
  // すでにデータが存在しているかを確認して、存在している場合は何もしない
  const referenceDataUUID = `${uuid}_ref`;
  const referenceDataColNumber = dataSheetUUIDToColMap.get(referenceDataUUID);
  if (referenceDataColNumber !== undefined) {
    if (referenceDataColNumber === referenceInsertDataColNumber) return;

    // データシートの同じ異なる位置に存在する場合はコピーする
    moveColumnData(
      dataSheet,
      referenceDataColNumber,
      referenceInsertDataColNumber
    );
    return;
  }

  // 参照シートの読み込み
  const referenceMap = createReferenceMap(referenceSheet);
  if (!referenceMap) throw new Error();

  moveColumnData(dataSheet, insertDataColNumber, referenceInsertDataColNumber);

  insertDataSheetHeader(
    dataSheet,
    `${uuid}_ref`,
    `${variableName}_ref`,
    `${displayName}(ref)`,
    insertDataColNumber
  );

  dataSheet
    .getRange(
      dataSheetHeaderRowNumber + DATA_SHEET_START_ROW_OFFSET,
      insertDataColNumber,
      dataSheet.getLastRow()
    )
    .setValues(
      dataSheet
        .getRange(
          dataSheetHeaderRowNumber + DATA_SHEET_START_ROW_OFFSET,
          insertDataColNumber,
          dataSheet.getLastRow()
        )
        .getValues()
        .map(x => x.map(y => referenceMap.get(y as string)))
    );

  // プルダウンの作成
  createPullDown(
    dataSheet,
    dataSheetHeaderRowNumber + DATA_SHEET_START_ROW_OFFSET,
    insertDataColNumber,
    [...referenceMap.values()]
  );

  // 参照シートのValueをKeyに変換する関数を作成
  const formula = createReferenceSwitchFormula(referenceMap);
  const dataSheetDataStartRowNumber =
    dataSheetHeaderRowNumber + DATA_SHEET_START_ROW_OFFSET;
  const dataSheetDataEndRowNumber = dataSheet.getLastRow();
  dataSheet
    .getRange(
      dataSheetDataStartRowNumber,
      referenceInsertDataColNumber,
      dataSheetDataEndRowNumber,
      1
    )
    .setValues(
      dataSheet
        .getRange(
          dataSheetDataStartRowNumber,
          insertDataColNumber,
          dataSheetDataEndRowNumber,
          1
        )
        .getValues()
        .map((x, rowIndex) =>
          x.map(() =>
            formula.replace(
              CELL_NAME,
              `${getCellName(
                rowIndex + dataSheetDataStartRowNumber,
                insertDataColNumber
              )}`
            )
          )
        )
    );
}
