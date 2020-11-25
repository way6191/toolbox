const express = require('express')
const cors = require('cors')
const bodyParser = require('body-parser')

const createExcel = require('./interface/CreateExcel.js')
const getAll = require('./interface/GetReleaseInfo.js')

const app = express()
const port = 9191

app.use(cors())
app.use(bodyParser.urlencoded({
  extended: true
}));
app.use(bodyParser.json());

app.post('/createExcel', (req, res) => {
  const param = req.body;
  let result = false;

  createExcel(param, isok => {
    result = isok;
    res.end(JSON.stringify(result));
  });
})

app.get('/getReleaseInfo', (req, res) => {
  const param = req.query;

  let result = getAll(param.folder);

  res.end(JSON.stringify(result));
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})