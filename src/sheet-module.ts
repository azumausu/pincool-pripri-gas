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

  throw new Error(`「${keyValue}」が存在しません。`);
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
