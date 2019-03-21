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

////////
app.use('/demand2', demand2);
app.use('/demand1', demand1);
app.use('/overduecc', overduecc);
app.use('/suspension', suspensioncc);
app.use('/prelistingcc', prelistingcc);
app.use('/prelisting', prelisting);
app.use('/postlistingunsecured', postlistingunsecured);
app.use('/postlistingunsecuredcc', postlistingunsecuredcc);
app.use('/postlistingsecured', postlistingsecured);
app.use('/day40', day40);
app.use('/day90', day90);
app.use('/day30', day30);
app.use('/prelistingremedial', prelistingremedial);

router.get('/', function (req, res) {
    res.json({ message: 'Demand letters ready!' }); 
  });
  
  app.use(bodyParser.json());
  app.use(bodyParser.urlencoded({extended: true}));
  app.use(cors())


  //add the router
app.use('/', router);
app.listen(process.env.port || 8004);

console.log('Running at Port 8004');