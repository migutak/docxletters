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
// const word2pdf = require('word2pdf');
var data = require('./data.js');
const cors = require('cors')

// Define font files
var fonts = {
  Roboto: {
    normal: 'fonts/Roboto-Regular.ttf',
    bold: 'fonts/Roboto-Medium.ttf',
    italics: 'fonts/Roboto-Italic.ttf',
    bolditalics: 'fonts/Roboto-MediumItalic.ttf'
  }
};

var PdfPrinter = require('pdfmake');
var printer = new PdfPrinter(fonts);

const LETTERS_DIR = data.filePath;
const IMAGEPATH = data.imagePath;

const {
  Document,
  Paragraph,
  Packer,
  TextRun,
  BorderStyle,
  Borders,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom, HorizontalPositionAlign,
  VerticalPositionAlign, floating
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
  const DATE = dateFormat(new Date(), "dd-mmm-yyyy");

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

    document.createImage(fs.readFileSync(IMAGEPATH + 'coop.jpg'), 350, 60, {
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

  const ref = new TextRun("Our Ref: PRELISTING/" + letter_data.cardacct + '/' + DATE);
  const paragraphref = new Paragraph();
  ref.bold();
  ref.size(24);
  paragraphref.addRun(ref);
  document.addParagraph(paragraphref);

  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'dd-mmm-yyyy'));
  const pddate = new Paragraph();
  ddate.size(24);
  pddate.addRun(ddate);
  document.addParagraph(pddate);

  document.createParagraph(" ");

  const name = new TextRun(letter_data.cardname);
  const pname = new Paragraph();
  name.size(20);
  pname.addRun(name);
  document.addParagraph(pname);

  const address = new TextRun(letter_data.address + '- ' + letter_data.rpcode);
  const paddress = new Paragraph();
  address.size(20);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.size(20);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  const dear = new TextRun("Dear Sir/Madam ");
  const pdear = new Paragraph();
  dear.size(20);
  pdear.addRun(dear);
  document.addParagraph(pdear);
  document.createParagraph(" ");

  const headertext = new TextRun("PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1)(a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.size(20);
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  paragraphheadertext.center();
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  const txt = new TextRun("We wish to inform you that in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs). ");
  const ptxt = new Paragraph();
  txt.size(20);
  ptxt.addRun(txt);
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt1 = new TextRun("A default in your card debt repayment will result in a negative impact on your credit record. If your card debt is classified as Non-Performing as per the Banking Act & Prudential Guidelines and/or as per the Microfinance Act, your credit profile at the CRBs will be adversely affected. ");
  const ptxt1 = new Paragraph();
  txt1.size(20);
  ptxt1.addRun(txt1);
  document.addParagraph(ptxt1);

  document.createParagraph(" ");
  const txt2 = new TextRun("Please note that your card account number " + letter_data.cardacct + ", card number " + letter_data.cardnumber + " is currently in default. It is outstanding at " + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + "DR with arrears of " + numeral(Math.abs(letter_data.exp_pmnt)).format('0,0.00') + "DR, having not paid the full installment(s) for 60 days. This card debt continues to accrue interest at a rate of 1.083% per month, on the daily outstanding balance and late payment fees at the rate of 5% on the arrears amount plus an excess fee of Kshs.1,000.00 monthly (if the total balance is above the limit).");
  const ptxt2 = new Paragraph();
  txt2.size(20);
  ptxt2.addRun(txt2);
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("We hereby notify you that we will proceed to adversely list you with the CRBs if your card debt becomes non-performing. To avoid an adverse listing, you are advised to clear the outstanding arrears within 30 days from the date of this letter. Payment can be made via Mpesa Paybill No. 400200 Account No. CR " + letter_data.cardacct + " ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  const txt4 = new TextRun("You have a right of access to your credit report at the CRBs and you may dispute any erroneous information. You may request for your report by contacting the CRBs at the following addresses: ");
  const ptxt4 = new Paragraph();
  txt4.size(20);
  ptxt4.addRun(txt4);
  document.addParagraph(ptxt4);

  document.createParagraph(" ");
  //start crb
  const crb = new TextRun("TransUnion CRB                                         Metropol CRB                                      Creditinfo CRB Kenya Ltd");
  const pcrb = new Paragraph();
  crb.size(18);
  crb.bold();
  pcrb.addRun(crb);
  document.addParagraph(pcrb);

  const crb1 = new TextRun("Delta Corner Annex, Ring Road,	       1st Floor, Shelter Afrique Centre, 	       Park Suites, Office 12, Second Floor");
  const pcrb1 = new Paragraph();
  crb1.size(18);
  pcrb1.addRun(crb1);
  document.addParagraph(pcrb1);

  const crb2 = new TextRun("Westlands Lane, Nairobi. 	                       Upper Hill, Nairobi.  	                       Parklands Road, Nairobi.");
  const pcrb2 = new Paragraph();
  crb2.size(18);
  pcrb2.addRun(crb2);
  document.addParagraph(pcrb2);

  const crb3 = new TextRun("P.O. Box 46406, 00100, NAIROBI, KENYA        P.O Box 35331–00200 NAIROBI. 	        P.O. Box 38941-00623, NAIROBI ");
  const pcrb3 = new Paragraph();
  crb3.size(16);
  pcrb3.addRun(crb3);
  document.addParagraph(pcrb3);

  const crb4 = new TextRun("Tel: +254 (0) 20 51799/3751360/2/4/5         Tel: +254(0)20 2689881/27113575       Tel: 020 3757272  ");
  const pcrb4 = new Paragraph();
  crb4.size(18);
  pcrb4.addRun(crb4);
  document.addParagraph(pcrb4);

  const crb5 = new TextRun("Fax:+254(0)20 3751344                                Fax: +254 (0) 20273572 ");
  const pcrb5 = new Paragraph();
  crb5.size(18);
  pcrb5.addRun(crb5);
  document.addParagraph(pcrb5);

  const crb6 = new TextRun("Email: info@crbafrica.com	                       Email: creditbureau@metropol.co.ke    Email: cikinfo@creditinfo.co.ke");
  const pcrb6 = new Paragraph();
  crb6.size(18);
  pcrb6.addRun(crb6);
  document.addParagraph(pcrb6);

  const crb9 = new TextRun("Website: www.crbafrica.com	                       Website: www.metropolcorporation.com  Website: ke.creditinfo.com ");
  const pcrb9 = new Paragraph();
  crb9.size(18);
  // crb9.underline();
  crb9.color("blue")
  pcrb9.addRun(crb9);
  document.addParagraph(pcrb9);

  document.createParagraph(" ");
  document.createParagraph("Yours sincerely, ");
  // sign
  //document.createImage(fs.readFileSync(IMAGEPATH + "sign_rose.png"), 100, 50);
  //sign
  const sign = new TextRun(" ");
  const psign = new Paragraph();
  sign.size(20);
  psign.addRun(sign);
  document.addParagraph(psign);

  const signtext = new TextRun("BRANCH MANAGER.");
  const paragraphsigntext = new Paragraph();
  signtext.bold();
  signtext.underline();
  signtext.size(22);
  paragraphsigntext.addRun(signtext);
  document.addParagraph(paragraphsigntext);

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.docx", buffer);
    // if pdf format
    if (letter_data.format == 'pdf') {
      var dd = {
        pageSize: 'A4',
        pageOrientation: 'portrait',
        // [left, top, right, bottom] or [horizontal, vertical] or just a number for equal margins
        pageMargins: [50, 60, 50, 60],
        footer: {
          columns: [
            { text: 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group Managing Director & CEO), M. Malonza (Vice Chairman),J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, R. Simani (Mrs), L. Karissa, G. Mburia.\n\n' }
          ],
          style: 'superMargin'
        },

        content: [
          {
            alignment: 'justify',
            columns: [
              {
                image: 'coop.jpg',
                width: 300
              },
              {
                type: 'none',
                alignment: 'right',
                fontSize: 9,
                ol: [
                  'The Co-operative Bank of Kenya Limited',
                  'Co-operative Bank House',
                  'Haile Selassie Avenue',
                  'P.O. Box 48231-00100 GPO, Nairobi',
                  'Tel: (020) 3276100',
                  'Fax: (020) 2227747/2219831',
                  { text: 'www.co-opbank.co.ke', color: 'blue', link: 'http://www.co-opbank.co.ke' }
                ]
              },
            ],
            columnGap: 10
          },
          "Our Ref: PRELISTING/" + letter_data.cardacct + '/' + DATE,
          '\n' + dateFormat(new Date(), 'dd-mmm-yyyy'),
          '\n' + letter_data.cardname,
          '' + letter_data.address + '-' + letter_data.rpcode,
          letter_data.city,

          '\nDear Sir/Madam',

          {
            text: '\nRE: PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1)(a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013',
            style: 'subheader'
          },

          { text: '\nWe wish to inform you that in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs). ', alignment: 'justify' },

          "\nA default in your card debt repayment will result in a negative impact on your credit record. If your card debt is classified as Non-Performing as per the Banking Act & Prudential Guidelines and/or as per the Microfinance Act, your credit profile at the CRBs will be adversely affected. ",

          {
            text: '\nPlease note that your card account number ' + letter_data.cardacct + ', card number ' + letter_data.cardnumber + ' is currently in default. It is outstanding at ' + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + 'DR with arrears of ' + numeral(Math.abs(letter_data.exp_pmnt)).format('0,0.00') + 'DR, having not paid the full installment(s) for 60 days. This card debt continues to accrue interest at a rate of 1.083% per month, on the daily outstanding balance and late payment fees at the rate of 5% on the arrears amount plus an excess fee of Kshs.1,000.00 monthly (if the total balance is above the limit).',
            fontSize: 10, alignment: 'justify'
          },

          "\nWe hereby notify you that we will proceed to adversely list you with the CRBs if your card debt becomes non-performing. To avoid an adverse listing, you are advised to clear the outstanding arrears within 30 days from the date of this letter. Payment can be made via Mpesa Paybill No. 400200 Account No. CR " + letter_data.cardacct + " ",

          "\nYou have a right of access to your credit report at the CRBs and you may dispute any erroneous information. You may request for your report by contacting the CRBs at the following addresses: ",

          { text: '\nYours sincerely, ' },
          {
            //image: 'sign_rose.png',
            width: 100,
            height: 50
          },
          { text: '  ' },
          { text: 'BRANCH MANAGER.' }

        ],
        styles: {
          header: {
            fontSize: 18,
            bold: true,
            alignment: 'right',
            margin: [0, 190, 0, 80]
          },
          subheader: {
            fontSize: 12,
            bold: true,
            decoration: 'underline'
          },
          superMargin: {
            margin: [20, 0, 40, 0],
            fontSize: 8, alignment: 'center', opacity: 0.5
          },
          quote: {
            italics: true
          },
          small: {
            fontSize: 8
          }
        },
        defaultStyle: {
          fontSize: 10
        }
      }; // end dd

      var options = {
        // ...
      }

      var pdfDoc = printer.createPdfKitDocument(dd, options);
      //pdfDoc.pipe(fs.createWriteStream(LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.pdf"));
      //pdfDoc.end();

      // ensures response is sent only after pdf is created
      writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "prelistingcc.pdf");
      pdfDoc.pipe(writeStream);
      pdfDoc.end();
      writeStream.on('finish', function () {
        // do stuff with the PDF file
        // send response
        res.json({
          result: 'success',
          message: LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.pdf",
          filename: letter_data.cardacct + DATE + "prelistingcc.pdf"
        })
      });
    } else {
      // res.sendFile(path.join(LETTERS_DIR + letter_data.cardacct + DATE + 'prelistingcc.docx'));
      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.docx",
        filename: letter_data.cardacct + DATE + "prelistingcc.docx"
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
