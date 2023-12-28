import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_DISPLAY_NAME_ROW_OFFSET,
  DATA_SHEET_KEY_NAME_ROW_OFFSET,
  DATA_SHEET_UUID_ROW_OFFSET,
} from '../../constants/data_sheet';
import { CELL_NAME, HEADER_START_MARKER } from '../../constants/common';

export function createDataSheetUUIDToColMap(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet
): Map<string, number> {
  const dataSheetUUIDToColMap = new Map<string, number>();

  const dataSheetHeaderRowIndex = getHeaderRowIndex(dataSheet);
  const dataSheetUUIDRowIndex =
    dataSheetHeaderRowIndex + DATA_SHEET_UUID_ROW_OFFSET;
  const dataSheetUUIDRowNumber = dataSheetUUIDRowIndex + 1;

  for (let i = 0; i < dataSheet.getLastColumn(); i++) {
    const colIndex = i + DATA_SHEET_COL_OFFSET;
    const colNumber = colIndex + 1;

    // uuidを取得
    const dataUUIDRange = dataSheet.getRange(
      dataSheetUUIDRowNumber,
      colNumber,
      1,
      1
    );
    if (dataUUIDRange.isBlank()) continue;

    // uuidとデータのkey-valueを作成
    dataSheetUUIDToColMap.set(dataUUIDRange.getValue(), colNumber);
  }

  return dataSheetUUIDToColMap;
}

// 定義シートのIdとDisplayNameを持つ列を追加する
export function insertDataSheetHeader(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  uuid: string,
  key: string,
  displayName: string,
  insertDataColNumber: number,
  dataSheetHeaderRowNumber: number
) {
  // uuidを挿入
  dataSheet
    .getRange(
      dataSheetHeaderRowNumber + DATA_SHEET_UUID_ROW_OFFSET,
      insertDataColNumber
    )
    .setValue(uuid);

  // dataシートに項目名を挿入
  dataSheet
    .getRange(
      dataSheetHeaderRowNumber + DATA_SHEET_KEY_NAME_ROW_OFFSET,
      insertDataColNumber
    )
    .setValue(key);

  // dataシートに表示名を挿入
  dataSheet
    .getRange(
      dataSheetHeaderRowNumber + DATA_SHEET_DISPLAY_NAME_ROW_OFFSET,
      insertDataColNumber
    )
    .setValue(displayName);
}

// Headerの開始行のインデックスを取得する
export function getHeaderRowIndex(
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): number {
  const index = getRowIndex(
    sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues(),
    1,
    HEADER_START_MARKER
  );

  if (index === undefined)
    throw new Error(
      `Sheet:${sheet.getSheetName()}の1列目にHeaderの開始位置(${HEADER_START_MARKER})が存在しません。`
    );

  return index;
}

// SWITCH関数を作成する
// SWITCH(CELL, value, key, value, key, value, key, ...)
export function createReferenceSwitchFormula(
  referenceMap: Map<string, string>
) {
  let formula = `=SWITCH(${CELL_NAME}`;
  referenceMap.forEach((value, key) => {
    formula += `, "${value}", "${key}"`;
  });
  formula += `)`;

  return formula;
}

// プルダウンの作成
export function createPullDown(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  startRowNumber: number,
  colNumber: number,
  pullDownList: string[]
) {
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(pullDownList, true)
    .build();
  sheet.getRange(startRowNumber, colNumber).setDataValidation(validationRule);
}

// 列の値の移動
export function moveColumnData(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  fromColumn: number,
  toColumn: number
) {
  const numRows = sheet.getLastRow();

  const rangeFrom = sheet.getRange(1, fromColumn, numRows);
  const rangeTo = sheet.getRange(1, toColumn, numRows);

  rangeTo.setValues(rangeFrom.getValues());
  rangeFrom.clear();
}

// 特定の文字列が入っている行のインデックスを取得する
export function getRowIndex(
  cellValues: unknown[][],
  searchColumnNumber: number,
  keyValue: string
) {
  const colIndex = searchColumnNumber - 1;
  for (let row = 0; row < cellValues.length; row++) {
    if (cellValues[row][colIndex] === keyValue) {
      return row;
    }
  }
}

// 特定の文字列が入っている列の淫デッっくすを取得する
export function getColIndex(cellValues: unknown[], keyValue: string) {
  for (let col = 0; col < cellValues.length; col++) {
    if (cellValues[col] === keyValue) {
      return col;
    }
  }

  throw new Error(`カラム「${keyValue}」が存在しません。`);
}

export function getCellName(row: number, column: number) {
  // 列のアルファベット名を取得
  const columnLetter = columnToLetter(column);
  // セル名を組み立てる（列のアルファベット名と行番号の組み合わせ）
  return columnLetter + row;
}

function columnToLetter(column: number) {
  let temp,
    letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
