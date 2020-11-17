const Excel = require('exceljs');
const sizeOf = require('image-size');
var fs = require('fs');
var path = require('path');
const util = require('util');

const readdir = util.promisify(fs.readdir);

/**
 * imgFolder取得
 * @param Folder 対象フォルダー
 */
async function getFolder(folder) {
  let folders;
  try {
    folders = await readdir(folder);
  } catch (e) {
    console.log(e);
  }
  return folders || [];
}

/**
 * img取得
 * @param imgFolder 対象フォルダー
 */
async function getImg(imgFolder) {
  let imgs;
  try {
    imgs = await readdir(imgFolder);
  } catch (e) {
    console.log(e);
  }
  return imgs || [];
}

async function createExcel({
  tplPath,
  shift,
  folder,
  scale
}) {
  const outname = tplPath.replace(".xlsx", "-created.xlsx");

  const imgFolders = await getFolder(folder);

  const newbook = new Excel.Workbook();

  for (let index = 0; index < imgFolders.length; index++) {
    foldername = imgFolders[index];

    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(tplPath);

    const worksheet = workbook.getWorksheet("sample");

    //  フォルダーパース
    let imgFolder = path.join(folder, foldername);

    // シートを新規
    let newsheet = newbook.addWorksheet(foldername);
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

    const imgs = await getImg(imgFolder);

    let pageSize = 0;

    // 画像を挿入
    for (let i = 0; i < imgs.length; i++) {
      imgName = imgs[i];
      // ページ
      pageSize = i;
      // イメージパース
      let img = path.join(imgFolder, imgName);
      let imgtype = path.extname(img).replace(".", "");
      // 移動を計算
      let move = parseInt(printTitleHeight) + parseInt(printAreaHeight * pageSize);

      if (i > 0) {
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
        filename: img,
        extension: imgtype,
      });

      let imgprop = sizeOf(img);

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

module.exports = createExcel;