var express = require('express');
var router = express.Router();
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const word2pdf = require('word2pdf-promises');
const cors = require('cors')

var data = require('./data.js');

const LETTERS_DIR = data.filePath;
const IMAGEPATH = data.imagePath;

const { Document, Paragraph, Packer, TextRun } = docx;

router.use(bodyParser.urlencoded({
  extended: true
}));

router.use(bodyParser.json());
router.use(cors())
 
/*router.use(function (req, res, next) {
  res.setHeader('Access-Control-Allow-Origin', 'http://localhost:4200');
  res.setHeader('Access-Control-Allow-Methods', 'POST');
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');
  res.setHeader('Access-Control-Allow-Credentials', true);
  next();
});*/

router.post('/download', function (req, res) {
  const letter_data = req.body;
  const GURARANTORS = req.body.guarantors || [];
  const INCLUDELOGO = req.body.showlogo;
  const DATA = req.body.accounts;
  const DATE = dateFormat(new Date(), "dd-mmm-yyyy");
  //
  //
  const document = new Document();
  
    const footer1 = new TextRun("Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group Managing Director & CEO), M. Malonza (Vice Chairman),")
      .size(16)
    const parafooter1 = new Paragraph()
    parafooter1.addRun(footer1).center();
    document.Footer.addParagraph(parafooter1);
    const footer2 = new TextRun("J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, R. Simani (Mrs), L. Karissa, G. Mburia.")
      .size(16)
    const parafooter2 = new Paragraph()
    parafooter1.addRun(footer2).center();
    document.Footer.addParagraph(parafooter2);

    //logo start
    if (INCLUDELOGO == true) {
    document.createImage(fs.readFileSync(IMAGEPATH + "coop.jpg"), 350, 60, {
      floating: {
        horizontalPosition: {
          offset: 1000000,
        },
        verticalPosition: {
          offset: 1014400,
        },
        margins: {
          top: 0,
          bottom: 201440,
        },
      },
    });
  }
  // logo end

  document.createParagraph("The Co-operative Bank of Kenya Limited").right();
  document.createParagraph("Co-operative Bank House").right();
  document.createParagraph("Haile Selassie Avenue").right();
  document.createParagraph("P.O.Box 48231-00100 GPO, Nairobi").right();
  document.createParagraph("Tel: (020) 3276100").right();
  document.createParagraph("Fax: (020) 2227747/2219831").right();
  document.createParagraph("Website: www.co-opbank.co.ke").right();

  document.createParagraph(" ");

  document.createParagraph("Our Ref: DAY40/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'dd-mmm-yyyy' ));
  const pddate = new Paragraph();
  ddate.size(20);
  pddate.addRun(ddate);
  document.addParagraph(pddate);

  const register = new TextRun("BY REGISTERED POST");
  const pregister = new Paragraph();
  register.size(20);
  pregister.addRun(register);
  pregister.right();
  document.addParagraph(pregister);

  const copy = new TextRun("Copy by ordinary Mail");
  const pcopy = new Paragraph();
  copy.size(20);
  pcopy.addRun(copy);
  pcopy.right();
  document.addParagraph(pcopy);

  document.createParagraph(" ");
  const name = new TextRun(letter_data.custname);
  const pname = new Paragraph();
  name.size(20);
  pname.addRun(name);
  document.addParagraph(pname);

  const address = new TextRun(letter_data.address + '- ' + letter_data.postcode);
  const paddress = new Paragraph();
  address.size(20);
  paddress.addRun(address);
  document.addParagraph(paddress);

  document.createParagraph(" ");
  document.createParagraph("Dear Sir/Madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: NOTIFICATION OF SALE OF PROPERTY L.R NO. xxxxxxxxxxxxxxxxx	");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("We refer to our notices dated xxxxxxxxxxxxx, xxxxxxxxxxxxxxx and xxxxxxxxxxxxxxxxxx");

  document.createParagraph(" ");
  const txt3 = new TextRun("As you are fully aware and despite the notices mentioned above, you have not rectified the default and you owe the Bank the sum of "+letter_data.accounts[0].currency +' '+numeral(Math.abs(letter_data.accounts[0].oustbalance)).format('0,0.0')+"DR as at "+DATE+" in respect of a facility granted to "+letter_data.custname+", full particulars whereof are well within your knowledge. ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("The said facility is secured by inter alia, a legal charge over L.R NO. xx registered in the name of xxxxxx.");
  document.createParagraph(" ");

  const note1 = new TextRun("TAKE NOTICE");
  note1.bold();
  note1.size(24);
  const note = new TextRun("TAKE NOTICE that pursuant to the provisions of Section 96(2) of the Land Act, 2012 the Bank intends to exercise its statutory power of sale over L.R NO. xxxxxx aforesaid after expiry of FORTY (40) DAYS from the date of service of this Notice upon yourself unless you rectify the default and all the outstanding balances owned to the Bank are fully settled within the aforesaid period. ");
  const pnote = new Paragraph();
  
  pnote.addRun(note);
  pnote.justified();
  document.addParagraph(pnote);

  document.createParagraph(" ");
  document.createParagraph("Please note that any repayment arrangements entered into between yourselves and the Bank and/or any payments made by you after the date of this notice shall be accepted by the Bank strictly on account and without prejudice to the Bank's right to proceed and realize its securities as aforesaid.");
  document.createParagraph(" ");
  document.createParagraph("FURTHER NOTE that pursuant to the provisions of Section 103 of the Land Act, 2012, you are at liberty to apply to the Court for any relief that the Court may deem fit against the Bank's remedy.");
  

  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph("                                                                                                    JOYCE MBINDA");
  document.createParagraph(letter_data.arocode                                           +    "                                                                                     MANAGER REMEDIAL MANAGEMENT");
  document.createParagraph("CREDIT MANAGEMENT DIVISION.                                    CREDIT MANAGEMENT DIVISION.");


  if (GURARANTORS.length>0) {
    document.createParagraph("cc: ");

    for (g = 0; g < GURARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GURARANTORS[g].guarantorname);
      document.createParagraph(GURARANTORS[g].address);
    }
  }

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "day40.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "day40.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + 'day40.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.acc + DATE + "day40.pdf",
              filename: letter_data.acc + DATE + "day40.pdf"
            })
          }, error => {
            console.log('error ...', error)
            res.json({
              result: 'error',
              message: 'Exception occured'
            });
          })
      }
      convert();
    } else {
      //
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.acc + DATE + "day40.docx",
        filename: letter_data.acc + DATE + "day40.docx"
      })
    }
  }).catch((err) => {
    console.log(err);
    res.json({
      result: 'error',
      message: 'Exception occured'
    });
  });
});

module.exports = router;
