const express = require('express')
const cors = require('cors')
const bodyParser = require('body-parser')

const createExcel = require('./interface/CreateExcel.js')

const app = express()
const port = 9191

app.use(cors())
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.post('/createExcel', (req, res) => {
  const param = req.body;
  createExcel(param);
  res.end();
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})