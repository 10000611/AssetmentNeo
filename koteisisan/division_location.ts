function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("資産分類区分、利用部門、管理部門をマスタの番号に変換");
  const assetTypeSheet = workbook.getWorksheet("資産分類マスタ");
  const masterSheet = workbook.getWorksheet("部門マスタ");
  const locationSheet = workbook.getWorksheet("ロケーションマスタ");

  const usedRange = sheet.getUsedRange();
  const rowCount = usedRange.getRowCount();

  // L列: 資産分類区分（列11）
  const lRange = sheet.getRangeByIndexes(1, 11, rowCount - 1, 1);
  const lColumn = lRange.getValues();

  // M列: 利用部門（列12）
  const mRange = sheet.getRangeByIndexes(1, 12, rowCount - 1, 1);
  const mColumn = mRange.getValues();

  // N列: 管理部門（列13）
  const nRange = sheet.getRangeByIndexes(1, 13, rowCount - 1, 1);
  const nColumn = nRange.getValues();

  // O列: ロケーション（列14）
  const oRange = sheet.getRangeByIndexes(1, 14, rowCount - 1, 1);
  const oColumn = oRange.getValues();

  // 資産分類マスタ取得（A列:コード, B列:資産分類名）
  const assetValues = assetTypeSheet.getUsedRange().getValues();
  const assetMap = new Map<string, string>();
  for (let i = 1; i < assetValues.length; i++) {
    const code = assetValues[i][0]?.toString().trim();
    const name = assetValues[i][1]?.toString().trim();
    if (code && name) {
      assetMap.set(name, code);
    }
  }

  // 部門マスタ取得（A列:コード, B列:部門名）
  const masterValues = masterSheet.getUsedRange().getValues();
  const masterMap = new Map<string, string>();
  for (let i = 1; i < masterValues.length; i++) {
    const code = masterValues[i][0]?.toString().trim();
    const name = masterValues[i][1]?.toString().trim();
    if (code && name) {
      masterMap.set(name, code);
    }
  }

  // ロケーションマスタ取得（A列:コード, B列:ロケーション名）
  const locationValues = locationSheet.getUsedRange().getValues();
  const locationMap = new Map<string, string>();
  for (let i = 1; i < locationValues.length; i++) {
    const code = locationValues[i][0]?.toString().trim();
    const name = locationValues[i][1]?.toString().trim();
    if (code && name) {
      locationMap.set(name, code);
    }
  }

  // 資産分類区分（L列）変換
  for (let i = 0; i < lColumn.length; i++) {
    const assetName = lColumn[i][0]?.toString().trim();
    lColumn[i][0] = assetName && assetMap.has(assetName) ? assetMap.get(assetName) : "";
  }
  lRange.setValues(lColumn);

  // 利用部門（M列）変換
  for (let i = 0; i < mColumn.length; i++) {
    const deptName = mColumn[i][0]?.toString().trim();
    mColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
  }
  mRange.setValues(mColumn);

  // 管理部門（N列）変換
  for (let i = 0; i < nColumn.length; i++) {
    const deptName = nColumn[i][0]?.toString().trim();
    nColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
  }
  nRange.setValues(nColumn);

  // ロケーション（O列）変換
  for (let i = 0; i < oColumn.length; i++) {
    const locName = oColumn[i][0]?.toString().trim();
    oColumn[i][0] = locName && locationMap.has(locName) ? locationMap.get(locName) : "";
  }
  oRange.setValues(oColumn);
}
