import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_DISPLAY_NAME_ROW_OFFSET,
  DATA_SHEET_VARIABLE_NAME_ROW_OFFSET,
  DATA_SHEET_UUID_ROW_OFFSET,
} from '../../constants/data_sheet';
import {
  CELL_NAME,
  DEFINE_VARIABLE_NAME,
  DISPLAY_NAME,
  HEADER_START_MARKER,
  UUID_KEY_NAME,
} from '../../constants/common';
import { DataSheetMetadata } from '../../types/data-sheet-metadata';
import { DEFINE_SHEET_DATA_START_ROW_OFFSET } from '../../constants/define_sheet';
import { REFERENCE_SHEET_NAME } from '../../constants/reference_sheet';
import { createReferenceMap } from './reference-sheet-module';
import { SheetInsertionMetadata } from '../../types/sheet-insertion-metadata';

// 定義シートで定義されているデータからデータシートに挿入する情報を作成する
export function createSheetInsertionMetadata(
  defineSheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const metadata: SheetInsertionMetadata[] = [];

  const headerRowIndex = getHeaderRowIndex(defineSheet);
  const headerRowNumber = headerRowIndex + 1;
  const lastRowNumber = defineSheet.getLastRow();
  const lastColNumber = defineSheet.getLastColumn();
  const uuidColIndex = getColIndex(
    defineSheet.getRange(headerRowNumber, 1, 1, lastColNumber).getValues()[0],
    UUID_KEY_NAME
  );
  const variableNameColIndex = getColIndex(
    defineSheet.getRange(headerRowNumber, 1, 1, lastColNumber).getValues()[0],
    DEFINE_VARIABLE_NAME
  );
  const displayNameColIndex = getColIndex(
    defineSheet.getRange(headerRowNumber, 1, 1, lastColNumber).getValues()[0],
    DISPLAY_NAME
  );
  const referenceSheetColIndex = getColIndex(
    defineSheet.getRange(headerRowNumber, 1, 1, lastColNumber).getValues()[0],
    REFERENCE_SHEET_NAME
  );
  const uuidColNumber = uuidColIndex + 1;
  const variableNameColNumber = variableNameColIndex + 1;
  const displayNameColNumber = displayNameColIndex + 1;
  const referenceSheetColNumber = referenceSheetColIndex + 1;

  let insertionColNumber = 1;
  for (let i = 0; i < lastRowNumber - headerRowIndex; i++) {
    const readRowNumber =
      headerRowNumber + DEFINE_SHEET_DATA_START_ROW_OFFSET + i;
    const uuid = defineSheet
      .getRange(readRowNumber, uuidColNumber, 1, 1)
      .getValue();
    const variableName = defineSheet
      .getRange(readRowNumber, variableNameColNumber, 1, 1)
      .getValue();
    const displayName = defineSheet
      .getRange(readRowNumber, displayNameColNumber, 1, 1)
      .getValue();
    const referenceSheetNameRange = defineSheet.getRange(
      readRowNumber,
      referenceSheetColNumber,
      1,
      1
    );
    const referenceSheetName = referenceSheetNameRange.getValue() as string;

    const referenceValueExists = !referenceSheetNameRange.isBlank();
    if (referenceValueExists) {
      // 参照カラム側の追加
      const referenceSheet = spreadSheet.getSheetByName(referenceSheetName);
      if (!referenceSheet)
        throw new Error(`参照シート(${referenceSheetName})が存在しません。`);
      const referenceMap = createReferenceMap(referenceSheet);
      if (!referenceMap)
        throw new Error(
          `参照シート(${referenceSheetName})のデータが存在しません。`
        );

      metadata.push({
        uuid: `${uuid}_ref`,
        variableName: `${variableName}`,
        displayName: `${displayName}`,
        // 自身が参照シートなので持っていない
        referenceSheetName: referenceSheetName,
        insertionColumnNumber: insertionColNumber,
        isReferenceColumn: true,
        hasReferenceColumn: false,
        referenceColumnNumber: -1,
        referenceMap: referenceMap,
      });
      insertionColNumber++;

      // 実データ側のカラムを追加
      metadata.push({
        uuid: uuid,
        variableName: `${variableName}`,
        displayName: `値（${displayName}）`,
        referenceSheetName: referenceSheetName,
        insertionColumnNumber: insertionColNumber,
        isReferenceColumn: false,
        hasReferenceColumn: true,
        referenceColumnNumber: insertionColNumber - 1,
        referenceMap: referenceMap,
      });
      insertionColNumber++;
      continue;
    }

    metadata.push({
      uuid: uuid,
      variableName: variableName,
      displayName: displayName,
      referenceSheetName: '',
      insertionColumnNumber: insertionColNumber,
      isReferenceColumn: false,
      hasReferenceColumn: false,
      referenceColumnNumber: -1,
      referenceMap: null,
    });
    insertionColNumber++;
  }

  return metadata;
}

export function createDataSheetUUIDToSheetInsertionMetadataMap(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const uuidToMetadataMap = new Map<string, DataSheetMetadata>();

  const headerRowIndex = getHeaderRowIndex(dataSheet);
  const uuidRowIndex = headerRowIndex + DATA_SHEET_UUID_ROW_OFFSET;
  const variableNameRowIndex =
    headerRowIndex + DATA_SHEET_VARIABLE_NAME_ROW_OFFSET;
  const displayNameRowIndex =
    headerRowIndex + DATA_SHEET_DISPLAY_NAME_ROW_OFFSET;
  const uuidRowNumber = uuidRowIndex + 1;

  for (let i = 0; i < dataSheet.getLastColumn(); i++) {
    const colIndex = i + DATA_SHEET_COL_OFFSET;
    const colNumber = colIndex + 1;

    // uuidを取得
    const dataUUIDRange = dataSheet.getRange(uuidRowNumber, colNumber, 1, 1);
    if (dataUUIDRange.isBlank()) continue;

    const metadata: DataSheetMetadata = {
      uuid: dataUUIDRange.getValue() as string,
      variableName: dataSheet
        .getRange(variableNameRowIndex, colNumber, 1, 1)
        .getValue() as string,
      displayName: dataSheet
        .getRange(displayNameRowIndex, colNumber, 1, 1)
        .getValue() as string,
      columnNumber: colNumber,
      data: dataSheet
        .getRange(
          headerRowIndex + 1,
          colNumber,
          dataSheet.getLastRow() - headerRowIndex,
          1
        )
        .getValues(),
    };
    uuidToMetadataMap.set(dataUUIDRange.getValue(), metadata);
  }

  return uuidToMetadataMap;
}

// 定義シートのIdとDisplayNameを持つ列を追加する
export function insertDataSheetHeader(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  uuid: string,
  key: string,
  displayName: string,
  insertDataColNumber: number
) {
  const dataSheetHeaderRowNumber = getHeaderRowIndex(dataSheet) + 1;

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
      dataSheetHeaderRowNumber + DATA_SHEET_VARIABLE_NAME_ROW_OFFSET,
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
  sheet
    .getRange(startRowNumber, colNumber, sheet.getLastRow(), 1)
    .setDataValidation(validationRule);
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

  rangeFrom.clearDataValidations();
  rangeFrom.clearContent();
}

// 列の値のコピー(dataシートのヘッダー開始位置からペーストを開始する)
export function copyColumnDataForDataSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  copyData: unknown[][],
  copyToColumn: number
) {
  const dataSheetHeaderRowIndex = getHeaderRowIndex(sheet);
  const dataSheetUUIDRowIndex =
    dataSheetHeaderRowIndex + DATA_SHEET_UUID_ROW_OFFSET;
  const dataSheetUUIDRowNumber = dataSheetUUIDRowIndex + 1;
  const numRows = copyData.length;

  const rangeTo = sheet.getRange(dataSheetUUIDRowNumber, copyToColumn, numRows);

  rangeTo.clearContent();

  rangeTo.setValues(copyData);
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
