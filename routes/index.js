var express = require('express');
var router = express.Router();
const XLSX = require('xlsx');

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.get('/excel', (req, res, next) => {
  console.log(`/excel called`);
  const workbook = XLSX.readFile(__dirname + "/../test.xlsx");
  const sheets = workbook.SheetNames;
  const dec = sheets[13];
  console.log(dec);
  const worksheet = workbook.Sheets[dec];
  res.json(worksheet);
});

module.exports = router;
