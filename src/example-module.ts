/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
const DEFINE_SHEET_NAME = 'define';
const PREVIOUS_ROW_COUNT_KEY = 'prevRowCount';

export function onChangeEvent(e: GoogleAppsScript.Events.SheetsOnChange) {
  Logger.log('処理の開始');

  // 変更された範囲を取得
  const defineSheet = e.source.getSheetByName(DEFINE_SHEET_NAME);

  const prevRowCount = Number(
    PropertiesService.getDocumentProperties().getProperty(
      PREVIOUS_ROW_COUNT_KEY
    )
  );
  const currentRowCount = defineSheet?.getLastRow() ?? 0;
  Logger.log(
    'prevRowCount:' +
      prevRowCount.toString() +
      ' currentRowCount:' +
      currentRowCount.toString()
  );

  // 変更が行全体である場合
  if (prevRowCount < currentRowCount) {
    // 行が挿入されたか削除されたか判定
    Logger.log('新しい行が追加されました。');
  } else if (prevRowCount > currentRowCount) {
    Logger.log('行が削除されました。');
  } else {
    Logger.log('変更なし');
  }

  PropertiesService.getDocumentProperties().setProperty(
    PREVIOUS_ROW_COUNT_KEY,
    currentRowCount.toString()
  );
}
