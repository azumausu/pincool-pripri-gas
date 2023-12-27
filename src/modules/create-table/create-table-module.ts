import { createReferenceMap } from './reference-sheet-module';
import {
  createPullDown,
  createReferenceSwitchFormula,
  getCellName,
  getColIndex,
  getHeaderRowIndex,
  moveColumnData,
} from './sheet-module';
import {
  CELL_NAME,
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_DISPLAY_NAME_ROW_OFFSET,
  DATA_SHEET_KEY_NAME_ROW_OFFSET,
  DATA_SHEET_NAME,
  DATA_SHEET_START_ROW_OFFSET,
  DEFINE_KEY_NAME,
  DEFINE_SHEET_NAME,
  DISPLAY_NAME,
  REFERENCE_SHEET_NAME,
} from '../../constants/constant';

export function apply() {
  // シートの取得
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const defineSheet = spreadSheet.getSheetByName(DEFINE_SHEET_NAME);
  const dataSheet = spreadSheet.getSheetByName(DATA_SHEET_NAME);
  if (!defineSheet) throw new Error(`Defineシートが存在しません。`);
  if (!dataSheet) throw new Error('Dataシートが存在しません。');

  const defineRange = defineSheet.getRange(
    1,
    1,
    defineSheet.getLastRow(),
    defineSheet.getLastColumn()
  );

  const defineSheetValues = defineRange.getValues();
  const defineSheetHeaderRowIndex = getHeaderRowIndex(defineSheet);

  const idColIndex = getColIndex(
    defineSheetValues[defineSheetHeaderRowIndex],
    DEFINE_KEY_NAME
  );
  const displayNameColIndex = getColIndex(
    defineSheetValues[defineSheetHeaderRowIndex],
    DISPLAY_NAME
  );
  const referenceSheetColIndex = getColIndex(
    defineSheetValues[defineSheetHeaderRowIndex],
    REFERENCE_SHEET_NAME
  );

  // defineシートのHeader一行下からデータを「項目名」と「表示名」のデータを取得していき、
  // そのデータをdataシート側のHeaderに追加していく
  let referenceCount = 0;
  const dataSheetHeaderRowIndex = getHeaderRowIndex(dataSheet);
  const dataSheetHeaderRowNumber = dataSheetHeaderRowIndex + 1;

  for (
    let i = 1;
    i < defineSheetValues.length - defineSheetHeaderRowIndex;
    i++
  ) {
    const defineSheetRowIndex = defineSheetHeaderRowIndex + i;
    const insertDataColNumber = referenceCount + i + DATA_SHEET_COL_OFFSET;
    const id = defineSheetValues[defineSheetRowIndex][idColIndex];
    const displayName =
      defineSheetValues[defineSheetRowIndex][displayNameColIndex];

    insertColumn(
      dataSheet,
      id,
      displayName,
      insertDataColNumber,
      dataSheetHeaderRowNumber
    );

    // Referenceシートが定義されているか確認
    const referenceValue = defineSheetValues[defineSheetRowIndex][
      referenceSheetColIndex
    ] as string;

    // 参照シートがなければそのままカラム追加して終了
    if (!referenceValue || referenceValue.length === 0) continue;

    const referenceSheet = spreadSheet.getSheetByName(referenceValue);
    if (!referenceSheet) continue;

    insertReferenceColumn(
      dataSheet,
      referenceSheet,
      id,
      displayName,
      insertDataColNumber,
      dataSheetHeaderRowNumber
    );

    // 参照シート用の列を作成したのでインクリメントする
    referenceCount++;
  }
}

// 定義シートのIdとDisplayNameを持つ列を追加する
function insertColumn(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  key: string,
  displayName: string,
  insertDataColNumber: number,
  dataSheetHeaderRowNumber: number
) {
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

// 参照シートを挿入する
function insertReferenceColumn(
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  referenceSheet: GoogleAppsScript.Spreadsheet.Sheet,
  key: string,
  displayName: string,
  insertDataColNumber: number,
  dataSheetHeaderRowNumber: number
) {
  const referenceMap = createReferenceMap(referenceSheet);
  if (!referenceMap) throw new Error();

  moveColumnData(dataSheet, insertDataColNumber, insertDataColNumber + 1);

  insertColumn(
    dataSheet,
    `${key}_ref`,
    `${displayName}(ref)`,
    insertDataColNumber,
    dataSheetHeaderRowNumber
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
      insertDataColNumber + 1,
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
