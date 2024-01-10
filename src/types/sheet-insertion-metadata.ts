import { DefineSheetCellType } from './cell-type';

export type SheetInsertionMetadata = {
  // 挿入するUUID
  uuid: string;

  // インポート対象か
  importTarget: boolean;

  // セルタイプ
  cellType: DefineSheetCellType;

  // 項目名
  variableName: string;

  // 表示名
  displayName: string;

  // 参照シートの名前
  referenceSheetName: string;

  // データが存在している列番号
  insertionColumnNumber: number;

  // どの列を参照すれば良いか
  referenceColumnNumber: number;

  // 参照シート
  referenceMap: Map<string, string> | null;
};
