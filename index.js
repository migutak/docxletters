const express = require('express');
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
const bodyParser = require("body-parser");
const router = express.Router();
const cors = require('cors')

//include the routes file
var demand2 = require('./routes/demand2');
var demand1 = require('./routes/demand1');
var overduecc = require('./routes/overduecc');
var suspensioncc = require('./routes/suspensioncc');
var prelistingcc = require('./routes/prelistingcc');
var prelisting = require('./routes/prelisting');
var postlistingsecured = require('./routes/postlistingsecured');
var postlistingunsecured = require('./routes/postlistingunsecured');
var postlistingunsecuredcc = require('./routes/postlistingunsecuredcc');
var day40 = require('./routes/day40');
var day90 = require('./routes/day90');
var day30 = require('./routes/day30');
var prelistingremedial = require('./routes/prelistingremedial');
var ipfcancellation = require('./routes/ipfcancellation');
var ipfcancellationwithsend = require('./routes/ipfcancellationwithsend');
var revocation = require('./routes/revocation');

app.use('/docx/demand2', demand2);
app.use('/docx/demand1', demand1);
app.use('/docx/overduecc', overduecc);
app.use('/docx/suspension', suspensioncc);
app.use('/docx/prelistingcc', prelistingcc);
app.use('/docx/prelisting', prelisting);
app.use('/docx/postlistingunsecured', postlistingunsecured);
app.use('/docx/postlistingunsecuredcc', postlistingunsecuredcc);
app.use('/docx/postlistingsecured', postlistingsecured);
app.use('/docx/day40', day40);
app.use('/docx/day90', day90);
app.use('/docx/day30', day30);
app.use('/docx/prelistingremedial', prelistingremedial);
app.use('/docx/ipfcancellation', ipfcancellation);
app.use('/docx/ipfcancellationwithsend', ipfcancellationwithsend);
app.use('/docx/revocation', revocation);

router.get('/', function (req, res) {
    res.json({ message: 'Demand letters ready Home!' }); 
});
  
  app.use(bodyParser.json());
  app.use(bodyParser.urlencoded({extended: true}));
  app.use(cors())


  //add the router
app.use('/docx', router);
app.listen(process.env.port || 8004);

console.log('Running at Port 8004');