function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("AssetmentNeo投入データ(部門・ロケーション変換)");
    const masterSheet = workbook.getWorksheet("部門マスタ");
    const locationSheet = workbook.getWorksheet("ロケーションマスタ");

    const usedRange = sheet.getUsedRange();
    const rowCount = usedRange.getRowCount();

    // L列（列11）
    const lRange = sheet.getRangeByIndexes(1, 11, rowCount - 1, 1);
    const lColumn = lRange.getValues();

    // M列（列12）
    const mRange = sheet.getRangeByIndexes(1, 12, rowCount - 1, 1);
    const mColumn = mRange.getValues();

    // N列（列13） 
    const nRange = sheet.getRangeByIndexes(1, 13, rowCount - 1, 1);
    const nColumn = nRange.getValues();

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

    // L列（部門）変換
    for (let i = 0; i < lColumn.length; i++) {
        const deptName = lColumn[i][0]?.toString().trim();
        lColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
    }
    lRange.setValues(lColumn);

    // M列（部門）変換
    for (let i = 0; i < mColumn.length; i++) {
        const deptName = mColumn[i][0]?.toString().trim();
        mColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
    }
    mRange.setValues(mColumn);

    // N列（ロケーション）変換
    for (let i = 0; i < nColumn.length; i++) {
        const locName = nColumn[i][0]?.toString().trim();
        nColumn[i][0] = locName && locationMap.has(locName) ? locationMap.get(locName) : "";
    }
    nRange.setValues(nColumn);
}
