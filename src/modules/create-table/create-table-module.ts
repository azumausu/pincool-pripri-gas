import { insertDataSheetReferenceColumn } from './reference-sheet-module';
import {
  getColIndex,
  getHeaderRowIndex,
  insertDataSheetHeader,
} from './sheet-module';
import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_NAME,
  DEFINE_KEY_NAME,
  DEFINE_SHEET_NAME,
  DISPLAY_NAME,
  REFERENCE_SHEET_NAME,
  UUID_KEY_NAME,
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

  const uuidColIndex = getColIndex(
    defineSheetValues[defineSheetHeaderRowIndex],
    UUID_KEY_NAME
  );
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
    const uuid = defineSheetValues[defineSheetRowIndex][uuidColIndex];
    const key = defineSheetValues[defineSheetRowIndex][idColIndex];
    const displayName =
      defineSheetValues[defineSheetRowIndex][displayNameColIndex];

    insertDataSheetHeader(
      dataSheet,
      uuid,
      key,
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

    insertDataSheetReferenceColumn(
      dataSheet,
      referenceSheet,
      uuid,
      key,
      displayName,
      insertDataColNumber,
      dataSheetHeaderRowNumber
    );

    // 参照シート用の列を作成したのでインクリメントする
    referenceCount++;
  }
}
