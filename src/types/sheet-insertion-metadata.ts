export type SheetInsertionMetadata = {
  // 挿入するUUID
  uuid: string;

  importTarget: boolean;

  // 項目名
  variableName: string;

  // 表示名
  displayName: string;

  // 参照シートの名前
  referenceSheetName: string;

  // データが存在している列番号
  insertionColumnNumber: number;

  // 参照データを表示するためのカラムか
  // uuidにrefがつく。pull downを作成するカラム
  isReferenceColumn: boolean;

  // 参照データを持つカラムか
  // pull downの値から関数で値を決めるカラム。
  hasReferenceColumn: boolean;

  // どの列を参照すれば良いか
  referenceColumnNumber: number;

  // 参照シート
  referenceMap: Map<string, string> | null;
};
