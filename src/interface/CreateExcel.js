const Excel = require('exceljs');
const sizeOf = require('image-size');

async function getSample({tplPath, shift, imgFolder, scale}) {

  const outname = tplPath.replace(".xlsx", "-created.xlsx");

  const newbook = new Excel.Workbook();

  for (let index = 0; index < 1; index++) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(tplPath);

    const worksheet = workbook.getWorksheet("sample");
    // シートを新規
    let newsheet = newbook.addWorksheet("No" + index);
    // シート属性設定
    newsheet.rowBreaks = worksheet.rowBreaks;
    newsheet.properties = worksheet.properties;
    newsheet.pageSetup = worksheet.pageSetup;
    newsheet.headerFooter = worksheet.headerFooter;
    newsheet.dataValidations = worksheet.dataValidations;
    newsheet.views = worksheet.views;
    newsheet.autoFilter = worksheet.autoFilter;
    newsheet.sheetProtection = worksheet.sheetProtection;
    newsheet.tables = worksheet.tables;
    newsheet.conditionalFormattings = worksheet.conditionalFormattings;

    const lastRowNum = worksheet.lastRowNum;
    worksheet.eachRow((row, rowNum) => {
      if (rowNum !== lastRowNum) {
        const r = newsheet.addRow();
        Object.assign(r, row);
      }
    });

    // 印刷範囲を取得
    let printArea = worksheet.pageSetup.printArea;

    let printWidthHeight = printArea.split(":");

    // 印刷タイトルを取得
    let printTitlesRow = worksheet.pageSetup.printTitlesRow;
    let printTitleHeight = printTitlesRow.split(":")[1];

    // 印刷範囲の頭の行と縦
    let printStart = printWidthHeight[0];
    let printStartRow = printStart.replace(/[^\d]/g, '');
    let printStartColumn = printStart.replace(/[^a-zA-Z]/g, '');

    // 印刷範囲の尾の行と縦
    let printEnd = printWidthHeight[1];
    let printEndRow = printEnd.replace(/[^\d]/g, '');
    let printEndColumn = printEnd.replace(/[^a-zA-Z]/g, '');

    let printAreaHeight = parseInt(printEndRow) - parseInt(printTitleHeight);

    let pageSize = 0;

    // 画像を挿入
    for (let index = 0; index < 5; index++) {
      // ページ
      pageSize = index;
      // 移動を計算
      let move = parseInt(printTitleHeight) + parseInt(printAreaHeight * pageSize);

      if (index > 0) {
        worksheet.eachRow((row, rowNum) => {
          if (rowNum > printTitleHeight && rowNum !== lastRowNum) {
            const r = newsheet.addRow();
            r.height = row.height;
            row.eachCell({
              includeEmpty: true
            }, function (cell, colNumber) {
              r.getCell(colNumber).value = cell.value;
              r.getCell(colNumber).style = cell.style;
            });
          }
        });
      }

      // 改ページ
      let changePage = newsheet.getRow(parseInt(printEndRow) + parseInt(printAreaHeight * pageSize));
      changePage.addPageBreak();

      // イメージ読み取
      const imageId1 = newbook.addImage({
        filename: 'C:/Users/505304/Desktop/test/image/1/1.png',
        extension: 'png',
      });

      let imgprop = sizeOf('C:/Users/505304/Desktop/test/image/1/1.png');
      
      let imgConfig = {};
      // イメージ位置調整
      imgConfig.tl = {
        col: 1,
        row: parseInt(move) + parseInt(shift)
      };
      // イメージサイズ調整
      imgConfig.ext = {
        width: imgprop.width * scale,
        height: imgprop.height * scale
      };
      // イメージ挿入
      newsheet.addImage(imageId1, imgConfig);
    }

    // 印刷範囲を計算
    printEnd = printEndColumn + (parseInt(printEndRow) + (parseInt(printAreaHeight) * pageSize));
    // 印刷範囲を調整
    newsheet.pageSetup.printArea = printStart + ":" + printEnd;
  }

  await newbook.xlsx.writeFile(outname);
}

module.exports = getSample;