const express = require('express')
const app = express()
const port = 9090

const Excel = require('exceljs');
const path = require('path');

async function createExcel(req, res) {
  const filename = 
  path.join(__dirname, "sample/sample.xlsx")
  console.log(filename);
  
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename);

  res.send(workbook);

  // const sheet = workbook.addWorksheet('My Sheet');

  // workbook.xlsx.writeFile("C:/Users/way/Desktop/test/1.xlsx");
}


app.get('/', (req, res) => {
    
  createExcel(req, res);
  

    // const sheet = workbook.addWorksheet('My Sheet');

    // workbook.xlsx.writeFile("C:/Users/way/Desktop/test/1.xlsx");

})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})