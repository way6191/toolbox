const Excel = require('exceljs');
const sizeOf = require('image-size');

async function getSample() {
  let scale = 0.4;
  let shift = 0;

  const filename = "C:/Users/505304/Desktop/test/sample.xlsx"
  const outname = "C:/Users/505304/Desktop/test/エビデンス.xlsx"
  
  const newbook = new Excel.Workbook();

  for (let index = 0; index < 3; index++) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filename);

    const worksheet  = workbook.getWorksheet('sample');
    // シートを新規
    let newsheet = newbook.addWorksheet("No"+index);
    newsheet.model = Object.assign(worksheet.model, {
      mergeCells: worksheet.model.merges
    });
    newsheet.name = "No"+index;

    // 印刷範囲を取得
    let printArea = worksheet.pageSetup.printArea;

    let printWidthHeight = printArea.split(":");

    // 印刷タイトルを取得
    let printTitlesRow  = worksheet.pageSetup.printTitlesRow ;
    let printTitleHeight = printTitlesRow.split(":")[1];
    
    // 印刷範囲の頭の行と縦
    let printStart = printWidthHeight[0];
    let printStartRow = printStart.replace(/[^\d]/g,'');
    let printStartColumn = printStart.replace(/[^a-zA-Z]/g,'');

    // 印刷範囲の尾の行と縦
    let printEnd = printWidthHeight[1];
    let printEndRow = printEnd.replace(/[^\d]/g,'');
    let printEndColumn = printEnd.replace(/[^a-zA-Z]/g,'');

    let printAreaHeight = parseInt(printEndRow) - parseInt(printTitleHeight);
    
    // イメージ挿入開始行
    let imageInsertStart = parseInt(printTitleHeight) + 1;

    // 背景を準備
    const rows = newsheet.getRows(imageInsertStart, printAreaHeight);

    newsheet.getRow(printEndRow).addPageBreak();

    let pageSize = 0;
    
    // 画像を挿入
    for (let index = 0; index < 5; index++) {
      pageSize = index;
      // 移動を計算
      let move = imageInsertStart + printAreaHeight * index;
      // ページをコピー
      if (index > 0) {
        let idx = move;
        let lastrow;
        rows.forEach(row => {
          lastrow = newsheet.insertRow(idx++, row.values, 'o');
        });
        lastrow.addPageBreak();
      }
      // イメージ読み取
      const imageId1 = newbook.addImage({
        filename: 'C:/Users/505304/Desktop/test/image/1/1.png',
        extension: 'png',
      });
  
      let imgprop = sizeOf('C:/Users/505304/Desktop/test/image/1/1.png');
      // イメージサイズ調整
      let imgConfig = {};
      imgConfig.tl = { col: 1, row: parseInt(move) + parseInt(shift) - 1};
      imgConfig.ext = { width: imgprop.width * scale, height: imgprop.height * scale};
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
