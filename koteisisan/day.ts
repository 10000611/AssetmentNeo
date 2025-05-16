function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet("AssetmentNeo投入データ(数量分加工後)");
  const lastRow = sheet.getUsedRange().getRowCount();
  const startRow = 1; // A2から開始（0-index）

  // F列, G列, H列を 'yyyy/mm/dd' の日付文字列に変換
  for (let col = 5; col <= 7; col++) { // F(5), G(6), H(7)
    const dateRange = sheet.getRangeByIndexes(startRow, col, lastRow - 1, 1);
    const dateValues: (string | number | boolean | null)[][] = dateRange.getValues();

    const dateStrings: (string | number | boolean | null)[][] = dateValues.map(
      (row: (string | number | boolean | null)[]): (string | number | boolean | null)[] => {
        const excelDate: string | number | boolean | null = row[0];
        if (typeof excelDate === 'number') {
          const jsDate: Date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
          const yyyy: number = jsDate.getFullYear();
          const mm: string = String(jsDate.getMonth() + 1).padStart(2, '0');
          const dd: string = String(jsDate.getDate()).padStart(2, '0');
          return ["'" + `${yyyy}/${mm}/${dd}`]; // '2025/05/14 形式
        } else {
          return [""];
        }
      }
    );

    dateRange.setValues(dateStrings);
  }
}
