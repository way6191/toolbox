const Excel = require('exceljs');
const path = require('path');

async function createExcel() {
  const filename = 
  path.join(__dirname, "sample/sample.xlsx")
  console.log(filename);
  
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename);

  console.log(workbook);

  // const sheet = workbook.addWorksheet('My Sheet');

  // workbook.xlsx.writeFile("C:/Users/way/Desktop/test/1.xlsx");
}
