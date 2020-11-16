const express = require('express')
const createExcel = require('./interface/CreateExcel.js')

const app = express()
const port = 9999

app.get('/createExcel', (req, res) => {
  createExcel();
  res.end();
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})