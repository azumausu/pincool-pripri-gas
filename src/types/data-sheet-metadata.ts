export type DataSheetMetadata = {
  // 挿入するUUID
  uuid: string;

  // 項目名
  variableName: string;

  // 表示名
  displayName: string;

  // データが存在している列番号
  columnNumber: number;

  // Headerを除くデータ
  data: unknown[][]; // [row][col]
};
