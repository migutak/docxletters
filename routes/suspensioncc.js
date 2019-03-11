var express = require('express');
var router = express.Router();
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const bodyParser = require("body-parser");
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
  const GURARANTORS = req.body.guarantors;
  const INCLUDELOGO = req.body.showlogo;
  const DATA = req.body.accounts;
  const DATE = dateFormat(new Date(), "isoDate");

  const document = new Document();
  if (INCLUDELOGO == true) {
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

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph(" ");

  document.createParagraph(" ");
  document.createParagraph(" ");

  document.createParagraph(" ");

  const ref = new TextRun("Our Ref: SUSPENSION/" + letter_data.cardacct + '/' + DATE);
  const paragraphref = new Paragraph();
  ref.bold();
  ref.size(28);
  paragraphref.addRun(ref);
  document.addParagraph(paragraphref);

  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
  const pddate = new Paragraph();
  ddate.font("Garamond");
  ddate.size(28);
  pddate.addRun(ddate);
  document.addParagraph(pddate);

  document.createParagraph(" ");

  const name = new TextRun(letter_data.cardname);
  const pname = new Paragraph();
  name.size(28);
  pname.addRun(name);
  document.addParagraph(pname);

  const address = new TextRun(letter_data.address + '- ' + letter_data.rpcode);
  const paddress = new Paragraph();
  address.size(28);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.size(28);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  document.createParagraph(" ");

  document.createParagraph("Dear sir/madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: CO-OPCARD ACCOUNT NO: " + letter_data.cardacct);
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  const txt = new TextRun("Your Co-opcard offers many exclusive benefits in addition to the unsecured credit facility. In order for you to enjoy these benefits to the full, proper maintenance of the account is vital.  We regret this has not been the case.");
  const ptxt = new Paragraph();
  txt.size(24);
  ptxt.addRun(txt);
  ptxt.justified();
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt5 = new TextRun("Your account has been suspended for non-payment of your bills and currently your account reflects a balance of Kshs. " + letter_data.OUT_BALANCE + " and this does not include any bills that we may not have received. The account also continues to accrue 1.083% interest and 5% late payment charges on outstanding balance and overdue amount every month respectively.");
  const ptxt5 = new Paragraph();
  txt5.size(24);
  ptxt5.addRun(txt5);
  ptxt5.justified();
  document.addParagraph(ptxt5);

  document.createParagraph(" ");
  const txt2 = new TextRun("We are now giving you notice that your personal information and credit account details will be disclosed to the Credit Reference Bureau, in accordance with the Banking Act and CRB regulations 2013. Be advised that any credit defaults will remain on your credit file for up to five years from the date of settlement. ");
  const ptxt2 = new Paragraph();
  txt2.size(24);
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);


  document.createParagraph(" ");
  document.createParagraph("Yours sincerely, ");
  document.createParagraph(" ");
  // sign

  document.createImage(fs.readFileSync(IMAGEPATH + "sign_rose.png"), 100, 70);

  //sign
  document.createParagraph(" ");
  const sign = new TextRun("ROSE KARAMBU ");
  const psign = new Paragraph();
  sign.size(24);
  psign.addRun(sign);
  document.addParagraph(psign);

  const signtext = new TextRun("COLLECTIONS SUPPORT MANAGER.");
  const paragraphsigntext = new Paragraph();
  signtext.bold();
  signtext.underline();
  signtext.size(28);
  paragraphsigntext.addRun(signtext);
  document.addParagraph(paragraphsigntext);

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "suspension.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.cardacct + DATE + "suspension.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + 'suspension.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.cardacct + DATE + "suspension.pdf",
              filename: letter_data.acc + DATE + "suspension.pdf"
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
      // res.sendFile(path.join(LETTERS_DIR + letter_data.cardacct + DATE + 'suspension.docx'));
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.cardacct + DATE + "suspension.docx",
        filename: letter_data.cardacct + DATE + "suspension.docx"
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