import { createReferenceMap } from './reference-sheet-module';
import { getColIndex, getRowIndex, moveColumnData } from './sheet-module';
import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_NAME,
  DEFINE_KEY_NAME,
  DEFINE_SHEET_NAME,
  DISPLAY_NAME,
  READ_ROW_MARKER,
  REFERENCE_SHEET_NAME,
} from './constants/constant';

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
  const headerRowIndex = getRowIndex(defineSheetValues, 1, READ_ROW_MARKER);

  // headerが存在しない。
  if (!headerRowIndex)
    throw new Error('Defineシートのヘッダーに無効な編集がされました。');

  const idColIndex = getColIndex(
    defineSheetValues[headerRowIndex],
    DEFINE_KEY_NAME
  );
  const displayNameColIndex = getColIndex(
    defineSheetValues[headerRowIndex],
    DISPLAY_NAME
  );
  const referenceSheetColIndex = getColIndex(
    defineSheetValues[headerRowIndex],
    REFERENCE_SHEET_NAME
  );

  // defineシートのHeader一行下からデータを「項目名」と「表示名」のデータを取得していき、
  // そのデータをdataシート側のHeaderに追加していく
  let referenceCount = 0;
  const dataSheetHeaderRowIndex = getRowIndex(
    dataSheet.getRange(1, 1, dataSheet.getLastRow(), 1).getValues(),
    1,
    READ_ROW_MARKER
  );
  const dataSheetHeaderRowNumber = dataSheetHeaderRowIndex + 1;

  for (let i = 1; i < defineSheetValues.length - headerRowIndex; i++) {
    const defineSheetRowIndex = headerRowIndex + i;
    const insertDataColNumber = referenceCount + i + DATA_SHEET_COL_OFFSET;
    const id = defineSheetValues[defineSheetRowIndex][idColIndex];
    const displayName =
      defineSheetValues[defineSheetRowIndex][displayNameColIndex];

    // dataシートにIdを挿入
    dataSheet
      .getRange(dataSheetHeaderRowNumber, insertDataColNumber)
      .setValue(id);

    // dataシートに表示名を挿入
    dataSheet
      .getRange(dataSheetHeaderRowNumber + 1, insertDataColNumber)
      .setValue(displayName);

    // Referenceシートが定義されているか確認
    const referenceValue = defineSheetValues[defineSheetRowIndex][
      referenceSheetColIndex
    ] as string;

    // 参照シートがなければそのままカラム追加して終了
    if (!referenceValue || referenceValue.length === 0) continue;

    const referenceSheet = spreadSheet.getSheetByName(referenceValue);
    if (!referenceSheet) continue;

    const referenceMap = createReferenceMap(referenceSheet);
    if (!referenceMap) throw new Error();

    moveColumnData(dataSheet, insertDataColNumber, insertDataColNumber + 1);

    // dataシートにIdを挿入
    dataSheet
      .getRange(dataSheetHeaderRowNumber, insertDataColNumber)
      .setValue(`${id}_ref`);
    // dataシートに表示名を挿入
    dataSheet
      .getRange(dataSheetHeaderRowNumber + 1, insertDataColNumber)
      .setValue(`${displayName}(ref)`);
    dataSheet
      .getRange(
        dataSheetHeaderRowNumber + 2,
        insertDataColNumber,
        dataSheet.getLastRow()
      )
      .setValues(
        dataSheet
          .getRange(
            dataSheetHeaderRowNumber + 2,
            insertDataColNumber,
            dataSheet.getLastRow()
          )
          .getValues()
          .map(x => x.map(y => referenceMap.get(y as string)))
      );

    // 参照シート用の列を作成したのでインクリメントする
    referenceCount++;
  }
}
