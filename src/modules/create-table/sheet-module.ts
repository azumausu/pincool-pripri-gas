import { CELL_NAME } from '../../constants/constant';

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
