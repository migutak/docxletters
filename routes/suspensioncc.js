var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const word2pdf = require('word2pdf-promises');
const cors = require('cors')
var Minio = require("minio");

var minioClient = new Minio.Client({
  endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
  port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
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

const { Document, Paragraph, Packer, TextRun } = docx;

router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());

router.post('/download', function (req, res) {
  const letter_data = req.body;
  const INCLUDELOGO = req.body.showlogo;
  const DATE = dateFormat(new Date(), "isoDate");

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

    document.createImage(fs.readFileSync("coop.jpg"), 350, 60, {
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

  document.createParagraph("The Co-operative Bank of Kenya Limited").left();
  document.createParagraph("Head Office").left();
  document.createParagraph("The Co-operative Bank of Kenya Limited").left();
  document.createParagraph("Co-operative Bank House").left();
  document.createParagraph("Haile Selassie Avenue").left();
  document.createParagraph("P.O.Box 48231-00100 GPO, Nairobi").left();
  document.createParagraph("Tel: (020) 3276100").left();
  document.createParagraph("Fax: (020) 2227747/2219831").left();
  document.createParagraph("Website: www.co-opbank.co.ke").left();

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
  const ddate = new TextRun(dateFormat(new Date(), 'dd-mmm-yyyy'));
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

  const address = new TextRun(letter_data.address);
  const paddress = new Paragraph();
  address.size(28);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const address1 = new TextRun(letter_data.rpcode);
  const paddress1 = new Paragraph();
  address1.size(28);
  paddress1.addRun(address1);
  document.addParagraph(paddress1);


  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.size(28);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  document.createParagraph(" ");

  document.createParagraph("Dear Sir/Madam ");
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
  const txt5 = new TextRun("Your account has been suspended for non-payment of your bills and currently your account reflects a balance of Kshs. " + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + "DR and this does not include any bills that we may not have received. The account also continues to accrue 1.083% interest and 5% late payment charges on outstanding balance and overdue amount every month respectively.");
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
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "suspension.docx", buffer);
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
            { text: data.footeroneline }
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
          "Our Ref: SUSPENSION/" + letter_data.cardacct + '/' + DATE,
          '\n' + dateFormat(new Date(), 'dd-mmm-yyyy'),
          '\n' + letter_data.cardname,
          '' + letter_data.address + '-' + letter_data.rpcode,
          letter_data.city,

          '\nDear Sir/Madam',

          {
            text: "\nRE: CO-OPCARD ACCOUNT NO: " + letter_data.cardacct,
            style: 'subheader'
          },

          "\nYour Co-opcard offers many exclusive benefits in addition to the unsecured credit facility. In order for you to enjoy these benefits to the full, proper maintenance of the account is vital.  We regret this has not been the case.",

          "\nYour account has been suspended for non-payment of your bills and currently your account reflects a balance of Kshs. " + numeral(Math.abs(letter_data.out_balance)).format('0,0.00') + "DR and this does not include any bills that we may not have received. The account also continues to accrue 1.083% interest and 5% late payment charges on outstanding balance and overdue amount every month respectively.",

          "\nWe are now giving you notice that your personal information and credit account details will be disclosed to the Credit Reference Bureau, in accordance with the Banking Act and CRB regulations 2013. Be advised that any credit defaults will remain on your credit file for up to five years from the date of settlement. ",


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
      writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "suspension.pdf");
      pdfDoc.pipe(writeStream);
      pdfDoc.end();
      writeStream.on('finish', function () {
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "suspension.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "suspension.pdf"
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
            message: LETTERS_DIR + accnumber_masked + DATE + "suspension.pdf",
            filename: accnumber_masked + DATE + "suspension.pdf",
            savedfilename: savedfilename,
            objInfo: objInfo
          })
        });
        //save to mino end
      });
    } else {
      // save to minio
      const filelocation = LETTERS_DIR + accnumber_masked + DATE + "suspension.docx";
      const bucket = 'demandletters';
      const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "suspension.docx"
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
          message: LETTERS_DIR + accnumber_masked + DATE + "suspension.docx",
          filename: accnumber_masked + DATE + "suspension.docx",
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
