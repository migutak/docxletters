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

const {
  Document,
  Paragraph,
  Packer,
  TextRun
} = docx;

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
  if (INCLUDELOGO == 'Y') {
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

    document.createImage(fs.readFileSync("./coop.jpg"), 350, 60, {
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

  const ref = new TextRun("Our Ref: OVERDUE/" + letter_data.cardacct + '/' + DATE);
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
  name.font("Garamond");
  name.size(28);
  pname.addRun(name);
  document.addParagraph(pname);

  const address = new TextRun(letter_data.address + '- ' + letter_data.rpcode);
  const paddress = new Paragraph();
  address.font("Garamond");
  address.size(28);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.font("Garamond");
  city.size(28);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  const txtdear = new TextRun("Dear sir/madam ");
  const ptxtdear = new Paragraph();
  txtdear.font("Garamond");
  txtdear.size(28);
  ptxtdear.addRun(txtdear);
  document.addParagraph(ptxtdear);
  document.createParagraph(" ");

  const headertext = new TextRun("RE: YOUR CARD ACCOUNT NUMBER: " + letter_data.cardacct);
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.font("Garamond");
  headertext.size(28);
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  const txt1 = new TextRun("We would like to draw your attention to your Co-op card account which is currently overdue. The total amount overdue is Kshs " + numeral(letter_data.EXP_PMNT).format('0,0.00') + " while your current outstanding balance is Kshs " + numeral(letter_data.OUT_BALANCE).format('0,0.00') + " ");
  const ptxt1 = new Paragraph();
  txt1.size(24);
  txt1.font("Garamond");
  ptxt1.addRun(txt1);
  ptxt1.justified();
  document.addParagraph(ptxt1);

  document.createParagraph(" ");
  const txt2 = new TextRun("We therefore request you to send payment of the above overdue amount immediately to avoid escalation of the interest and late payment charges accruing at 1.083% and 5% every month respectively. ");
  const ptxt2 = new Paragraph();
  txt2.size(24);
  txt2.font("Garamond");
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("If you have any query regarding the above amount or suspect that your payment has been delayed, please feel free to contact the undersigned. If payment has already been sent, please ignore this letter. ");
  const ptxt3 = new Paragraph();
  txt3.size(24);
  txt3.font("Garamond");
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  const txt4 = new TextRun("Payment can be made via Mpesa Paybill No. 400200 Account No. CR " + letter_data.cardacct + " ");
  const ptxt4 = new Paragraph();
  txt4.size(24);
  txt4.font("Garamond");
  txt4.bold();
  ptxt4.addRun(txt4);
  document.addParagraph(ptxt4);

  document.createParagraph(" ");
  const txt5 = new TextRun("We appreciate the opportunity to serve you. ");
  const ptxt5 = new Paragraph();
  txt5.size(24);
  txt5.font("Garamond");
  ptxt5.addRun(txt5);
  document.addParagraph(ptxt5);

  document.createParagraph(" ");
  const txt6 = new TextRun("Kindly provide us with your email address by replying through Cardcentre@co-opbank.co.ke to enable us serve you better. ");
  const ptxt6 = new Paragraph();
  txt6.size(24);
  txt6.font("Garamond");
  txt6.bold();
  ptxt6.addRun(txt6);
  ptxt6.justified();
  document.addParagraph(ptxt6);

  document.createParagraph(" ");
  document.createParagraph("Yours sincerely, ");
  document.createParagraph(" ");
  // sign

  document.createImage(fs.readFileSync("./sign_rose.png"), 100, 70);

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
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "overdue.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.cardacct + DATE + "overdue.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + 'overdue.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.cardacct + DATE + "overdue.pdf",
              filename: letter_data.acc + DATE + "overduecc.pdf"
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
      // res.sendFile(path.join(LETTERS_DIR + letter_data.cardacct + DATE + 'overdue.docx'));
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.cardacct + DATE + "overdue.docx",
        filename: letter_data.acc + DATE + "overdue.docx"
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