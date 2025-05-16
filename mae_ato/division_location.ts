function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("資産分類区分、ロケーションをマスタの番号に変換");
    const assetTypeSheet = workbook.getWorksheet("資産分類マスタ");
    const locationSheet = workbook.getWorksheet("ロケーションマスタ");

    const usedRange = sheet.getUsedRange();
    const rowCount = usedRange.getRowCount();

    // L列: 資産分類区分（列11）
    const lRange = sheet.getRangeByIndexes(1, 11, rowCount - 1, 1);
    const lColumn = lRange.getValues();

    // N列: ロケーション（列13）
    const nRange = sheet.getRangeByIndexes(1, 13, rowCount - 1, 1);
    const nColumn = nRange.getValues();

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

    // ロケーション（N列）変換
    for (let i = 0; i < nColumn.length; i++) {
        const locName = nColumn[i][0]?.toString().trim();
        nColumn[i][0] = locName && locationMap.has(locName) ? locationMap.get(locName) : "";
    }
    nRange.setValues(nColumn);
}
