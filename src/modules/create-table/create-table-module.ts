import { insertDataSheetReferenceColumn } from './reference-sheet-module';
import {
  createDataSheetUUIDToColMap,
  getColIndex,
  getHeaderRowIndex,
  insertDataSheetHeader,
  moveColumnData,
} from './sheet-module';
import { DEFINE_SHEET_NAME } from '../../constants/define_sheet';
import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_NAME,
} from '../../constants/data_sheet';
import {
  DEFINE_KEY_NAME,
  DISPLAY_NAME,
  UUID_KEY_NAME,
} from '../../constants/common';
import { REFERENCE_SHEET_NAME } from '../../constants/reference_sheet';

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
  const defineHeaderRowValues = defineSheetValues[defineSheetHeaderRowIndex];

  const uuidColIndex = getColIndex(defineHeaderRowValues, UUID_KEY_NAME);
  const idColIndex = getColIndex(defineHeaderRowValues, DEFINE_KEY_NAME);
  const displayNameColIndex = getColIndex(defineHeaderRowValues, DISPLAY_NAME);
  const referenceSheetColIndex = getColIndex(
    defineHeaderRowValues,
    REFERENCE_SHEET_NAME
  );

  // defineシートのHeader一行下からデータを「項目名」と「表示名」のデータを取得していき、
  // そのデータをdataシート側のHeaderに追加していく
  let referenceCount = 0;
  const dataSheetHeaderRowIndex = getHeaderRowIndex(dataSheet);
  const dataSheetHeaderRowNumber = dataSheetHeaderRowIndex + 1;
  const dataSheetColNumber =
    defineSheetValues.length - defineSheetHeaderRowIndex;

  // データシートにすでに存在するデータの列番号とUUIDのマップを作成する
  const dataSheetUUIDToColMap = createDataSheetUUIDToColMap(dataSheet);

  for (let i = 1; i < dataSheetColNumber; i++) {
    const defineSheetRowIndex = defineSheetHeaderRowIndex + i;
    const defineSheetCurrentRowValues = defineSheetValues[defineSheetRowIndex];
    const insertDataColNumber = referenceCount + i + DATA_SHEET_COL_OFFSET;
    const uuid = defineSheetCurrentRowValues[uuidColIndex];
    const key = defineSheetCurrentRowValues[idColIndex];
    const displayName = defineSheetCurrentRowValues[displayNameColIndex];

    const dataColNumber = dataSheetUUIDToColMap.get(uuid);
    if (dataColNumber !== undefined) {
      // すでにデータシートの同じ位置に同一のデータが存在する場合は次のループへ
      // TODO: ここでreferenceシート側の値を確認してreferenceCountをインクリメントする必要がある？
      if (dataColNumber === insertDataColNumber) continue;

      // データシートの同じ異なる位置に存在する場合はコピーする
      // TODO: ここで、referenceシート側もmoveする必要がある？
      moveColumnData(dataSheet, dataColNumber, insertDataColNumber);
      continue;
    }

    insertDataSheetHeader(
      dataSheet,
      uuid,
      key,
      displayName,
      insertDataColNumber,
      dataSheetHeaderRowNumber
    );

    if (
      !tryInsertReferenceValue(
        spreadSheet,
        dataSheet,
        defineSheetCurrentRowValues[referenceSheetColIndex] as string,
        uuid,
        key,
        displayName,
        insertDataColNumber,
        dataSheetHeaderRowNumber,
        dataSheetUUIDToColMap
      )
    )
      continue;

    // 参照シート用の列を作成したのでインクリメントする
    referenceCount++;
  }
}

function tryInsertReferenceValue(
  spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  dataSheet: GoogleAppsScript.Spreadsheet.Sheet,
  referenceValue: string,
  uuid: string,
  key: string,
  displayName: string,
  insertDataColNumber: number,
  dataSheetHeaderRowNumber: number,
  dataSheetUUIDToColMap: Map<string, number>
) {
  // Referenceシートが定義されているか確認
  if (!referenceValue || referenceValue.length === 0) return false;

  const referenceSheet = spreadSheet.getSheetByName(referenceValue);

  // 参照シートがなければそのままカラム追加して終了
  if (!referenceSheet) return false;

  insertDataSheetReferenceColumn(
    dataSheet,
    referenceSheet,
    uuid,
    key,
    displayName,
    insertDataColNumber,
    dataSheetHeaderRowNumber,
    dataSheetUUIDToColMap
  );

  return true;
}
