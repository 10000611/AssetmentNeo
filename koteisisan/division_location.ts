function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("資産分類区分、利用部門、管理部門をマスタの番号に変換");
    const assetTypeSheet = workbook.getWorksheet("資産分類マスタ");
    const masterSheet = workbook.getWorksheet("部門マスタ");
    const locationSheet = workbook.getWorksheet("ロケーションマスタ");

    const usedRange = sheet.getUsedRange();
    const rowCount = usedRange.getRowCount();

    // M列: 資産分類区分（列12）
    const mRange = sheet.getRangeByIndexes(1, 12, rowCount - 1, 1);
    const mColumn = mRange.getValues();

    // N列: 利用部門（列13）
    const nRange = sheet.getRangeByIndexes(1, 13, rowCount - 1, 1);
    const nColumn = nRange.getValues();

    // O列: 管理部門（列14）
    const oRange = sheet.getRangeByIndexes(1, 14, rowCount - 1, 1);
    const oColumn = oRange.getValues();

    // P列: ロケーション（列15）
    const pRange = sheet.getRangeByIndexes(1, 15, rowCount - 1, 1);
    const pColumn = pRange.getValues();

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

    // 資産分類区分（M列）変換
    for (let i = 0; i < mColumn.length; i++) {
        const assetName = mColumn[i][0]?.toString().trim();
        mColumn[i][0] = assetName && assetMap.has(assetName) ? assetMap.get(assetName) : "";
    }
    mRange.setValues(mColumn);

    // 利用部門（N列）変換
    for (let i = 0; i < nColumn.length; i++) {
        const deptName = nColumn[i][0]?.toString().trim();
        nColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
    }
    nRange.setValues(nColumn);

    // 管理部門（O列）変換
    for (let i = 0; i < oColumn.length; i++) {
        const deptName = oColumn[i][0]?.toString().trim();
        oColumn[i][0] = deptName && masterMap.has(deptName) ? masterMap.get(deptName) : "";
    }
    oRange.setValues(oColumn);

    // ロケーション（P列）変換
    for (let i = 0; i < pColumn.length; i++) {
        const locName = pColumn[i][0]?.toString().trim();
        pColumn[i][0] = locName && locationMap.has(locName) ? locationMap.get(locName) : "";
    }
    pRange.setValues(pColumn);
}
