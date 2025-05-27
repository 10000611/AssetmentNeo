function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("数量_マスタ変換");


  // A2 から AP列まで（最大1000行）を取得
  const maxRows = 1000;
  const numCols = 42;
  const dataStartRow = 1; // A2 から開始（0-indexedで1）
  const dataRange = sheet.getRangeByIndexes(dataStartRow, 0, maxRows - 1, numCols);
  const dataValues = dataRange.getValues();

  // ヘッダー取得（1行目）
  const header = sheet.getRange("A1:AP1").getValues()[0];

  const quantityColIndex = 16; // Q列:「数量」列
  const priceExclTaxColIndex = 8;  // I列: 購入額合価(税抜)
  const taxColIndex = 9;           // J列: 消費税
  const priceInclTaxColIndex = 10; // K列: 購入額合価(税込)

  let expandedData: (string | number | boolean | null)[][] = [];

  for (let row of dataValues) {
    // 空行はスキップ（No列などが空かで判定）
    if (row.every(cell => cell === null || cell === "")) {
      continue;
    }

    const quantity = Number(row[quantityColIndex]);
    if (isNaN(quantity) || quantity < 1) continue;

    // 金額（整数）を取得
    const totalExclTax = Number(row[priceExclTaxColIndex]) || 0;
    const totalTax = Number(row[taxColIndex]) || 0;
    const totalInclTax = Number(row[priceInclTaxColIndex]) || 0;

    // 金額を数量で割る。Math.floorで小数点なし。小数点以下切り捨て。(1円以下)
    const baseExclTax = Math.floor(totalExclTax / quantity);
    const baseTax = Math.floor(totalTax / quantity);
    const baseInclTax = Math.floor(totalInclTax / quantity);

    // 総額から数量-1の金額を引く。
    const remainderExclTax = totalExclTax - baseExclTax * (quantity - 1);
    const remainderTax = totalTax - baseTax * (quantity - 1);
    const remainderInclTax = totalInclTax - baseInclTax * (quantity - 1);

    for (let i = 1; i <= quantity; i++) {
      let newRow = [...row];
      newRow[quantityColIndex] = 1; // 数量列は常に「1」に固定

      const isLast = i === quantity;
      //最後の1件に誤差を加算することで、金額の合計が元と一致するように調整する
      newRow[priceExclTaxColIndex] = isLast ? remainderExclTax : baseExclTax;
      newRow[taxColIndex] = isLast ? remainderTax : baseTax;
      newRow[priceInclTaxColIndex] = isLast ? remainderInclTax : baseInclTax;

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
