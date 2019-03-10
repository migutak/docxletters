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
  //
  //
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

  document.createParagraph("Our Ref: POSTLISTING/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
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
  document.createParagraph(letter_data.custname);
  document.createParagraph(letter_data.address);
  document.createParagraph(" ");

  document.createParagraph("Dear sir/madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("REF: DEMAND FOR OUTSTANDING BALANCES AND NOTICE OF LISTING AT THE CREDIT REFERENCE BUREAU ISSUED PURSUANT TO REGULATION 50 (1)(b) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013.");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  headertext.size(18);
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("We note with regret that you have failed to act on our various requests to you to either repay or regularise your account number; "+letter_data.acc+" that is in arrears. You can make payments through account number "+letter_data.acc+" at any branch of Co-operative Bank.");
  document.createParagraph(" ");
  document.createParagraph("The outstanding liabilities stand at Kshs. "+ letter_data.accounts[0].oustbalance+"Dr as at "+DATE+" broken down as follows:-");
  document.createParagraph(" ");

  document.createParagraph("Principle Loan                    Kshs:  " + numeral(letter_data.accounts[0].oustbalance).format('0,0.00'));
  document.createParagraph("Loan Arrears                       Kshs:  " + numeral(letter_data.accounts[0].totalarrears).format('0,0.00'));
  document.createParagraph(" ");

  document.createParagraph(" ");
  document.createParagraph("Please note that interest continues to accrue at various Bank rates until the outstanding balance is paid in full. ");
  document.createParagraph(" ");
  document.createParagraph("Kindly make the necessary arrangements to repay the outstanding balance within the next Fourteen (14) days from the date of this letter, i.e. on or before (Date), failure to which we shall have no option but to exercise any of the remedies below against you, to recover the said outstanding amount at your risk as to costs and expenses arising without further reference to you;");
  document.createParagraph("1.	Appoint an External Debt Collector. ");
  document.createParagraph("2.	File suit against you. ");
  document.createParagraph(" ");

  document.createParagraph("After revisions in 2012/2013 to the Banking Act (Cap 488), Central Bank Act, Microfinance Act, 2006 and the CRB Regulations, Banks and Microfinance Banks have been mandated to share information on all their borrowers, and their loan information with registered Credit Reference Bureaus (CRBs). This means that the CRBs will now hold information on both good and bad borrowers. A good loan repayment pattern will reflect in a borrower's credit report resulting in an attractive credit profile, which can allow a borrower to negotiate preferential loan agreements with lenders.  ");
  document.createParagraph(" ");
  document.createParagraph("Thus, in compliance to the law, and having borrowed with Co-operative Bank of Kenya Limited, we have forwarded your  ");

 

  document.createParagraph(" ");
  //start crb
  const crb = new TextRun("TransUnion CRB                                                   Metropol CRB");
  const pcrb = new Paragraph();
  crb.size(20);
  crb.bold();
  pcrb.addRun(crb);
  document.addParagraph(pcrb);

  const crb1 = new TextRun("2nd Floor, Prosperity House,                                   1st Floor, Shelter Afrique Centre, Upper Hill, Nairobi. ");
  const pcrb1 = new Paragraph();
  crb1.size(20);
  pcrb1.addRun(crb1);
  document.addParagraph(pcrb1);

  const crb2 = new TextRun("Westlands Road, Off Museum Hill,                         P.O Box 35331 - 00200 ");
  const pcrb2 = new Paragraph();
  crb2.size(20);
  pcrb2.addRun(crb2);
  document.addParagraph(pcrb2);

  const crb3 = new TextRun("Westlands, Nairobi. P.O. Box 46406, 00100           NAIROBI, KENYA. ");
  const pcrb3 = new Paragraph();
  crb3.size(20);
  pcrb3.addRun(crb3);
  document.addParagraph(pcrb3);

  const crb4 = new TextRun("NAIROBI, KENYA Telephone: +254 (0) 20          Telephone: +254 (0) 20 2689881/27113575  ");
  const pcrb4 = new Paragraph();
  crb4.size(20);
  pcrb4.addRun(crb4);
  document.addParagraph(pcrb4);

  const crb5 = new TextRun("51799/3751360/2/4/5 Fax: +254 (0) 20 3751344    Fax: +254 (0) 20273572 ");
  const pcrb5 = new Paragraph();
  crb5.size(20);
  pcrb5.addRun(crb5);
  document.addParagraph(pcrb5);

  const crb6 = new TextRun("Email: info@transunion.co.ke                                 Email: creditbureau@metropol.co.ke ");
  const pcrb6 = new Paragraph();
  crb6.size(20);
  pcrb6.addRun(crb6);
  document.addParagraph(pcrb6);

  const crb9 = new TextRun("Website: www.crbafrica.com                                  www.metropolcorporation.com  ");
  const pcrb9 = new Paragraph();
  crb9.size(20);
  // crb9.underline();
  crb9.color("blue")
  pcrb9.addRun(crb9);
  document.addParagraph(pcrb9);
  //stop crb

  document.createParagraph(" ");
  const txt = new TextRun("You are encouraged to ensure that your loan payments are always up-to date and regularly obtain your credit report from the bureaus above to ascertain the accuracy of your information. ");
  const ptxt = new Paragraph();
  txt.size(18);
  ptxt.addRun(txt);
  ptxt.justified();
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt2 = new TextRun("In need, please feel free to contact the undersigned at Credit Management Division, Co-operative House Building, and Mezzanine 2.Tel:020-3276xxx/0711049xxx/0732106xxx.");
  const ptxt2 = new Paragraph();
  txt2.size(18);
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("Be advised accordingly.");
  const ptxt3 = new Paragraph();
  txt3.size(18);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph(letter_data.manager);
  document.createParagraph("Officer-Remedial Management Department                                         Manager-Remedial Management Department");


  if (GURARANTORS) {
    document.createParagraph("cc: ");

    for (g = 0; g < GURARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GURARANTORS[g].name);
      document.createParagraph(GURARANTORS[g].address);
    }
  }

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "postlistingunsecured.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "postlistingunsecured.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + 'postlistingunsecured.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.acc + DATE + "postlistingunsecured.pdf",
              filename: letter_data.acc + DATE + "postlistingunsecured.pdf"
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
      // res.sendFile(path.join(LETTERS_DIR + letter_data.acc + DATE + 'postlistingunsecured.docx'));
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.acc + DATE + "postlistingunsecured.docx",
        filename: letter_data.acc + DATE + "postlistingunsecured.docx"
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