var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const cors = require('cors')
var Minio = require("minio");
require('log-timestamp');

var minioClient = new Minio.Client({
  endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
  port: process.env.MINIO_PORT || 9005,
  useSSL: false,
  accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
  secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});

var data = require('./data.js');

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

const {
  Document,
  Paragraph,
  Packer,
  TextRun
} = docx;

router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());

router.post('/download', function (req, res) {
  const letter_data = req.body;
  const INCLUDELOGO = req.body.showlogo;
  const DATE = dateFormat(new Date(), "dd-mmm-yyyy");

  const document = new Document();
  if (INCLUDELOGO == true) {
    const footer1 = new TextRun(data.footerfirst)
      .size(16)
    const parafooter1 = new Paragraph()
    parafooter1.addRun(footer1).center();
    document.Footer.addParagraph(parafooter1);
    const footer2 = new TextRun(data.footersecond)
      .size(16)
    const parafooter2 = new Paragraph()
    parafooter1.addRun(footer2).center();
    document.Footer.addParagraph(parafooter2);

    //logo start

    document.createImage(fs.readFileSync('coop.jpg'), 350, 60, {
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
  const ddate = new TextRun(dateFormat(new Date(), 'dd-mmm-yyyy'));
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

  const address = new TextRun(letter_data.address);
  const paddress = new Paragraph();
  address.font("Garamond");
  address.size(28);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const address1 = new TextRun(letter_data.rpcode);
  const paddress1 = new Paragraph();
  address1.font("Garamond");
  address1.size(28);
  paddress1.addRun(address1);
  document.addParagraph(paddress1);

  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.font("Garamond");
  city.size(28);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  document.createParagraph("Dear Sir/Madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: OVERDUE CARD PAYMENT ACCOUNT NUMBER: " + letter_data.cardacct);
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.font("Garamond");
  headertext.size(28);
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  const txt1 = new TextRun("We would like to draw your attention to your Co-op card account which is currently overdue. The total amount overdue is Kshs " + numeral(Math.abs(letter_data.exp_pmnt)).format('0,0.00') + "DR while your current outstanding balance is Kshs " + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + "DR ");
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
  const txt4 = new TextRun("Payment can be made via Mpesa Paybill No. 400200 Account No. CR" + letter_data.cardacct + " ");
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

  //document.createImage(fs.readFileSync(IMAGEPATH + "sign_rose.png"), 100, 70);

  //sign
  document.createParagraph(" ");
  const sign = new TextRun(" ");
  const psign = new Paragraph();
  sign.size(24);
  psign.addRun(sign);
  document.addParagraph(psign);

  const signtext = new TextRun("BRANCH MANAGER");
  const branchname = new TextRun(letter_data.branchname);
  const nosigntext = new TextRun("This letter is electronically generated and is valid without a signature");
  nosigntext.bold();
  nosigntext.italic();
  const paragraphsigntext = new Paragraph();
  const paragraphsigntext2 = new Paragraph();
  const paragraphsigntext3 = new Paragraph();
  //signtext.bold();
  //signtext.underline();
  //signtext.size(28);
  paragraphsigntext.addRun(signtext);
  paragraphsigntext2.addRun(branchname);
  paragraphsigntext3.addRun(nosigntext);
  document.addParagraph(paragraphsigntext);
  document.addParagraph(paragraphsigntext2);
  document.addParagraph(psign);
  document.addParagraph(psign);
  document.addParagraph(paragraphsigntext3);

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "overdue.docx", buffer);
    //conver to pdf
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
          "Our Ref: OVERDUE/" + letter_data.cardacct + '/' + DATE,
          '\n' + dateFormat(new Date(), 'dd-mmm-yyyy'),
          '\n' + letter_data.cardname,
          '' + letter_data.address + '-' + letter_data.rpcode,
          letter_data.city,

          '\nDear Sir/Madam',

          {
            text: "\nRE: OVERDUE CARD PAYMENT ACCOUNT NUMBER: " + letter_data.cardacct,
            style: 'subheader'
          },

          "\nWe would like to draw your attention to your Co-op card account which is currently overdue. The total amount overdue is Kshs " + numeral(Math.abs(letter_data.exp_pmnt)).format('0,0.00') + "DR while your current outstanding balance is Kshs " + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + "DR ",


          // { text: '\nWe wish to inform you that in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs). ', alignment: 'justify' },

          "\nWe therefore request you to send payment of the above overdue amount immediately to avoid escalation of the interest and late payment charges accruing at 1.083% and 5% every month respectively. ",

          "\nIf you have any query regarding the above amount or suspect that your payment has been delayed, please feel free to contact the undersigned. If payment has already been sent, please ignore this letter. ",

          {
            text: ['\nPayment can be made via ',
              { text: 'Mpesa Paybill No. 400200 Account No. CR ' + letter_data.cardacct, bold: true },
            ]
          },

          {
            text: '\nWe appreciate the opportunity to serve you. ',
            alignment: 'left'
          },

          {
            text: '\nKindly provide us with your email address by replying through Cardcentre@co-opbank.co.ke to enable us serve you better. ',
            alignment: 'left'
          },

          { text: '\nYours sincerely, ' },
          {
            //image: 'sign_rose.png',
            //width: 100,
            //height: 50
          },
          { text: '\n\nBRANCH MANAGER , ' },
          { text: letter_data.branchname },
          { text: '\n\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },

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

      // ensures response is sent only after pdf is created
      writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.cardacct + DATE + "overduecc.pdf");
      pdfDoc.pipe(writeStream);
      pdfDoc.end();
      writeStream.on('finish', function () {
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "overduecc.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "overduecc.pdf"
        var metaData = {
          'Content-Type': 'text/html',
          'Content-Language': 123,
          'X-Amz-Meta-Testing': 1234,
          'example': 5678
        }
        minioClient.fPutObject(bucket, savedfilename, filelocation, metaData, function (error, objInfo) {
          if (error) {
            console.log(error);
            res.status(500).json({
              success: false,
              error: error.message
            })
          }
          res.json({
            result: 'success',
            message: LETTERS_DIR + accnumber_masked + DATE + "overduecc.pdf",
            filename: accnumber_masked + DATE + "overduecc.pdf",
            savedfilename: savedfilename,
            objInfo: objInfo
          })
        });
        //save to mino end
      });
    } else {
      // save to minio
      const filelocation = LETTERS_DIR + accnumber_masked + DATE + "overduecc.docx";
      const bucket = 'demandletters';
      const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "overduecc.docx"
      var metaData = {
        'Content-Type': 'text/html',
        'Content-Language': 123,
        'X-Amz-Meta-Testing': 1234,
        'example': 5678
      }
      minioClient.fPutObject(bucket, savedfilename, filelocation, metaData, function (error, objInfo) {
        if (error) {
          console.log(error);
          res.status(500).json({
            success: false,
            error: error.message
          })
        }
        res.json({
          result: 'success',
          message: LETTERS_DIR + accnumber_masked + DATE + "overduecc.docx",
          filename: accnumber_masked + DATE + "overduecc.docx",
          savedfilename: savedfilename,
          objInfo: objInfo
        })
      });
      //save to mino end
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
