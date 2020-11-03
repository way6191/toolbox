const ExcelJS = require('exceljs');

exports.createExcel= function() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('My Sheet');

  workbook.xlsx.writeFile("C:/Users/505304/Desktop/test/1.xlsx");
}
