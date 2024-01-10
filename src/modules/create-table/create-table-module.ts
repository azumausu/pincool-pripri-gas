import {
  copyColumnDataForDataSheet,
  createDataSheetUUIDToSheetInsertionMetadataMap,
  createPullDown,
  createReferenceSwitchFormula,
  createSheetInsertionMetadata,
  getCellName,
  getHeaderRowIndex,
  insertDataSheetHeader,
} from './sheet-module';
import { DEFINE_SHEET_NAME } from '../../constants/define_sheet';
import {
  DATA_SHEET_COL_OFFSET,
  DATA_SHEET_NAME,
  DATA_SHEET_ROW_MAX,
  DATA_SHEET_START_ROW_OFFSET,
} from '../../constants/data_sheet';
import { CELL_NAME } from '../../constants/common';
import { DefineSheetCellType } from '../../types/cell-type';

// すでに入っているデータは保つように編集する
// ヘッダーと参照シートは変わっている可能性があるので毎回更新する
export function apply() {
  // シートの取得
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const defineSheet = spreadSheet.getSheetByName(DEFINE_SHEET_NAME);
  const dataSheet = spreadSheet.getSheetByName(DATA_SHEET_NAME);
  if (!defineSheet) throw new Error(`Defineシートが存在しません。`);
  if (!dataSheet) throw new Error('Dataシートが存在しません。');

  // データシートにすでに存在するデータの列番号とUUIDのマップを作成する
  const dataSheetUUIDToMetadataMap =
    createDataSheetUUIDToSheetInsertionMetadataMap(dataSheet);
  const insertionMetadata = createSheetInsertionMetadata(defineSheet);
  const dataSheetHeaderRowIndex = getHeaderRowIndex(dataSheet);
  const dataSheetHeaderRowNumber = dataSheetHeaderRowIndex + 1;
  const dataSheetDataStartRowNumber =
    dataSheetHeaderRowNumber + DATA_SHEET_START_ROW_OFFSET;
  const dataCount = DATA_SHEET_ROW_MAX - dataSheetDataStartRowNumber + 1;

  for (const metadata of insertionMetadata) {
    // 既にデータシート側に反映済みのカラムかを確認
    const currentData = dataSheetUUIDToMetadataMap.get(metadata.uuid);
    const insertionColNumber =
      metadata.insertionColumnNumber + DATA_SHEET_COL_OFFSET;

    // Defineシートと参照シートが持つ情報は毎回更新する

    // -- シートに存在するデータ検証ルールはどんな時は毎回削除する
    dataSheet
      .getRange(dataSheetDataStartRowNumber, insertionColNumber, dataCount, 1)
      .clearDataValidations();

    // -- ヘッダーの更新
    insertDataSheetHeader(
      dataSheet,
      metadata.importTarget,
      metadata.uuid,
      metadata.variableName,
      metadata.displayName,
      insertionColNumber
    );

    // -- 参照シートのプルダウンの更新
    if (metadata.cellType === DefineSheetCellType.ReferencePullDown) {
      createPullDown(
        dataSheet,
        dataSheetDataStartRowNumber,
        insertionColNumber,
        dataCount,
        [...metadata.referenceMap!.values()]
      );
    }

    // データの入力操作

    // -- 参照シートの関数の更新
    if (metadata.cellType === DefineSheetCellType.ReferenceFormula) {
      const formula = createReferenceSwitchFormula(metadata.referenceMap!);
      dataSheet
        .getRange(dataSheetDataStartRowNumber, insertionColNumber, dataCount, 1)
        .setFormulas(
          dataSheet
            .getRange(
              dataSheetDataStartRowNumber,
              insertionColNumber,
              dataCount,
              1
            )
            .getValues()
            .map((x, rowIndex) =>
              x.map(() =>
                formula.replace(
                  CELL_NAME,
                  `${getCellName(
                    rowIndex + dataSheetDataStartRowNumber,
                    metadata.referenceColumnNumber + DATA_SHEET_COL_OFFSET
                  )}`
                )
              )
            )
        );

      continue;
    }

    // -- 元々データが存在していない項目の場合はこの列のデータを全て空白にして終了
    if (currentData === undefined) {
      dataSheet
        .getRange(dataSheetDataStartRowNumber, insertionColNumber, dataCount, 1)
        .clearContent();

      continue;
    }

    // -- 現在のデータとデータ側に変化がない場合は何もしない。
    if (currentData.columnNumber === insertionColNumber) {
      continue;
    }

    // -- 別の列に今の列に挿入するデータがあった場合はデータをコピーする
    // -- ただし、参照シートの関数の場合はすでに挿入ずみなので何もしない
    copyColumnDataForDataSheet(dataSheet, currentData.data, insertionColNumber);
  }
}
