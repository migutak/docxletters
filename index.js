const express = require('express');
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
const bodyParser = require("body-parser");
const router = express.Router();
const cors = require('cors')

var demand1 = require('./routes/demand1');

app.use('/demand1', demand1);

router.get('/', function (req, res) {
    res.json({ message: 'hooray! welcome to our rest video api!' }); 
  });
  
  app.use(bodyParser.json());
  app.use(bodyParser.urlencoded({extended: true}));
  app.use(cors())


  //add the router
app.use('/', router);
app.listen(process.env.port || 8004);

console.log('Running at Port 8004');