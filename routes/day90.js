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

const { Document, Paragraph, Packer, TextRun } = docx;

var data = require('./data.js');

const LETTERS_DIR = data.filePath;

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
  //
  //
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

  document.createParagraph("Our Ref: DAY90/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'fullDate' ));
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
  document.createParagraph("Dear sir/madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: OUTSTANDING LIABILITIES DUE TO THE BANK ON ACCOUNT OF "+letter_data.acc+": BASE NO. xxxxxxxxx ");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("We refer to our notice dated xxxxxxxxxxx. ");

  document.createParagraph(" ");
  const txt3 = new TextRun("As you are fully aware and despite the referenced notice, the above account is in arrears of Kes. xxxxxxx dr as at (Date) which continues to accrue interest at xxx% per annum (equivalent to Kenya Bank's Reference Rate (KBRR) currently at xxxx% plus a margin of xxx% (K)) and late penalties of 0.5% per month and further the total outstanding sum due to the Bank as at (Date) is Kes. ……….. dr which continues to accrue interest at xxx% per annum (equivalent to Kenya Bank's Reference Rate (KBRR) currently at xxxx% plus a margin of xxx% (K)).");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("The liabilities are secured by way of a Legal charge over the properties: ");
  document.createParagraph("L.R.NO. xxxxxxxxxxxxxxxx I.N.O xxxxxxxxxxxxxx")
  document.createParagraph(" ");

  
  document.createParagraph("TAKE NOTICE that pursuant to the provisions of Section 90 of the Land Act, 2012, the Bank intends to take action and exercise remedies provided in this Section after the expiry of THREE (3) MONTHS from the date of service of this Notice upon yourself if you do not rectify the default by repaying the outstanding sum of Kes. xxxxxxxxxxxdr which includes the ");
  document.createParagraph(" ");
  document.createParagraph("Please be advised that if you fail to remedy the default and repay the outstanding amount as stated above the Bank shall exercise any of the remedies as stipulated in Section 90 (3) of the Land Act, 2012 against you which includes:");
  document.createParagraph("File suit against you for money due and owing ");
  document.createParagraph("Appoint a receiver of the income of the charged property ");
  document.createParagraph("Lease or sublease the charged property ");
  document.createParagraph("Enter into possession of the charged Property ");
  document.createParagraph("Sell the charged property. ");

  document.createParagraph(" ");
  document.createParagraph("FURTHER NOTE that pursuant to the provisions of Sections 90(2) (e) and 103 of the Land Act, 2012, you are at liberty to apply to the Court for any relief that the Court may deem fit to grant against the Bank's remedies. ");


  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(letter_data.manager);
  document.createParagraph("RELATIONSHIP OFFICER                                        HEAD - REMEDIAL MANAGEMENT");


  if (GURARANTORS) {
    document.createParagraph("cc: ");

    for (g = 0; g < GURARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GURARANTORS[g].name);
      document.createParagraph(GURARANTORS[g].address);
    }
  }

  document.createParagraph(" ");
  document.createParagraph("This letter is valid without a signature ");

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "day90.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "day90.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + 'day90.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.acc + DATE + "day90.pdf",
              filename: letter_data.acc + DATE + "day90.pdf"
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
      // res.sendFile(path.join(LETTERS_DIR + letter_data.acc + DATE + 'day90.docx'));
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.acc + DATE + "day90.docx",
        filename: letter_data.acc + DATE + "day90.docx"
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