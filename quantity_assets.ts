function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("AssetmentNeo投入データ(部門・ロケーション変換)");

  // 1000行を取得
  const maxRows = 1000;
  const numCols = 42;
  const dataStartRow = 1; // A2 から開始（0-indexedで1）
  const dataRange = sheet.getRangeByIndexes(dataStartRow, 0, maxRows - 1, numCols);
  const dataValues = dataRange.getValues();

  // ヘッダー取得（1行目）
  const header = sheet.getRange("A1:AP1").getValues()[0];

  const quantityColIndex = 15; // 「数量」列（P列）

  let expandedData: (string | number | boolean | null)[][] = [];

  for (let row of dataValues) {
    // 空行はスキップ（No列などが空かで判定）
    if (row.every(cell => cell === null || cell === "")) {
      continue;
    }

    const quantity = Number(row[quantityColIndex]);
    if (isNaN(quantity) || quantity < 1) continue;

    for (let i = 1; i <= quantity; i++) {
      let newRow = [...row];
      newRow[quantityColIndex] = i; // 「数量」列を1〜Nの連番に
      expandedData.push(newRow);
    }
  }

  // 出力先：展開結果シート
  let outputSheet = workbook.getWorksheet("AssetmentNeo投入データ(数量分加工後)");
  if (!outputSheet) {
    outputSheet = workbook.addWorksheet("AssetmentNeo投入データ(数量分加工後)");
  } else {
    outputSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);
  }

  // 出力：ヘッダー＋展開データ
  outputSheet.getRangeByIndexes(0, 0, expandedData.length + 1, numCols)
    .setValues([header, ...expandedData]);
}
