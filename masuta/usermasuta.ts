function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("元データ");
    const locationSheet = workbook.getWorksheet("ロケーションマスタ");

    const usedRange = sheet.getUsedRange();
    const rowCount = usedRange.getRowCount();

    // B列: ロケーション名（列1）
    const bRange = sheet.getRangeByIndexes(1, 1, rowCount - 1, 1);
    const bColumn = bRange.getValues();

    // ロケーションマスタ取得（A列:コード, B列:ロケーション名）
    const locationValues = locationSheet.getUsedRange().getValues();
    const locationMap = new Map<string, string>();
    for (let i = 1; i < locationValues.length; i++) {
        const code = locationValues[i][0]?.toString().trim(); // A列
        const name = locationValues[i][1]?.toString().trim(); // B列
        if (code && name) {
            locationMap.set(name, code);
        }
    }

    // ロケーション名 → コードに置換
    for (let i = 0; i < bColumn.length; i++) {
        const locName = bColumn[i][0]?.toString().trim();
        bColumn[i][0] = locName && locationMap.has(locName) ? locationMap.get(locName) : "";
    }

    // B列に上書き
    bRange.setValues(bColumn);
}
