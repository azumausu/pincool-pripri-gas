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
  DATA_SHEET_START_ROW_OFFSET,
} from '../../constants/data_sheet';
import { CELL_NAME } from '../../constants/common';

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
  const dataSheetLastRowNumber = dataSheet.getLastRow();

  for (const metadata of insertionMetadata) {
    // 既にデータシート側に反映済みのカラムかを確認
    const currentData = dataSheetUUIDToMetadataMap.get(metadata.uuid);
    const insertionColNumber =
      metadata.insertionColumnNumber + DATA_SHEET_COL_OFFSET;

    // ヘッダーは更新がかかっている可能性があるため常に更新をかける
    insertDataSheetHeader(
      dataSheet,
      metadata.uuid,
      metadata.variableName,
      metadata.displayName,
      insertionColNumber
    );

    // プルダウンのルールを削除しつつ、参照カラムの場合はプルダウンを追加で設定する。ここもどんな場合でも更新をかける
    dataSheet
      .getRange(
        dataSheetDataStartRowNumber,
        insertionColNumber,
        dataSheet.getLastRow(),
        1
      )
      .clearDataValidations();
    if (metadata.isReferenceColumn) {
      createPullDown(
        dataSheet,
        dataSheetDataStartRowNumber,
        insertionColNumber,
        [...metadata.referenceMap!.values()]
      );
    }

    // 参照カラムの実データの場合は、関数を追加で設定する。ここもどんな場合でも更新をかける
    if (metadata.hasReferenceColumn) {
      const formula = createReferenceSwitchFormula(metadata.referenceMap!);
      dataSheet
        .getRange(
          dataSheetDataStartRowNumber,
          insertionColNumber,
          dataSheetLastRowNumber,
          1
        )
        .setValues(
          dataSheet
            .getRange(
              dataSheetDataStartRowNumber,
              insertionColNumber,
              dataSheetLastRowNumber,
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
    }

    if (currentData !== undefined) {
      // 既に挿入しようとしている同一データが存在している場合は何もしない
      if (currentData.columnNumber === insertionColNumber) {
        continue;
      }

      // データが存在しているが別の列に存在している場合はデータをコピーする
      copyColumnDataForDataSheet(
        dataSheet,
        currentData.data,
        insertionColNumber
      );
      continue;
    }

    // 今入っているデータを破棄する
    const overrideRangeData = dataSheet.getRange(
      dataSheetHeaderRowNumber,
      insertionColNumber,
      dataSheetLastRowNumber,
      1
    );
    overrideRangeData.clearContent();
    overrideRangeData.clearDataValidations();
  }
}
