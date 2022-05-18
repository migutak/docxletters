const express = require('express');
const app = express();
var morgan = require('morgan');
const ecsFormat = require('@elastic/ecs-morgan-format');
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
var revocationwithsend = require('./routes/revocationwithsend');
var repossession = require('./routes/repossession');
var repossessionwithemail = require('./routes/repossessionwithemail');
var release = require('./routes/release');
var investigations = require('./routes/investigations');
var investigationswithemail = require('./routes/investigationswithemail');
var valuation = require('./routes/valuation');
var valuationwithemail = require('./routes/valuationwithemail');
var repossessionsendphy = require('./routes/repossessionsendphy');
var repossessionsendphy_word = require('./routes/repossessionsendphy_word');

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
app.use('/docx/revocationwithsend', revocationwithsend);
app.use('/docx/repossession', repossession);
app.use('/docx/repossessionwithemail', repossessionwithemail);
app.use('/docx/release', release);
app.use('/docx/investigators', investigations);
app.use('/docx/investigatorswithemail', investigationswithemail);
app.use('/docx/valuation', valuation);
app.use('/docx/valuationwithemail', valuationwithemail);
app.use('/docx/repossessionsendphy', repossessionsendphy);
app.use('/docx/repossessionsendphy_word', repossessionsendphy_word);

router.get('/', function (req, res) { 
  res.json({ message: 'Demand letters ready Home!' });
});

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(morgan(ecsFormat()));

app.use(cors())

//add the router
app.use('/docx', router);
app.listen(process.env.port || 8004);

console.log('Running at Port 8004');