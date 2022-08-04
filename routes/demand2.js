var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const cors = require('cors');
var Minio = require("minio");
require('log-timestamp');

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
const IMAGEPATH = data.imagePath;

const {
  Document,
  Paragraph,
  Packer,
  TextRun
} = docx;

router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());

router.post('/download', async function (req, res) {
  const letter_data = req.body;
  const GUARANTORS = req.body.guarantors || [];
  const INCLUDELOGO = req.body.showlogo;
  const DATA = req.body.accounts;
  const DATE = dateFormat(new Date(), "dd-mmm-yyyy");
  //
  //
  const rawaccnumber = letter_data.acc;
  const memo = rawaccnumber.substr(2, 3);
  const first4 = rawaccnumber.substring(0, 9);
  const last4 = rawaccnumber.substring(rawaccnumber.length - 4);

  const mask = rawaccnumber.substring(4, rawaccnumber.length - 4).replace(/\d/g, "*");
  accnumber_masked = first4 + 'xxxxx';
  const document = new Document();

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
  if (INCLUDELOGO == true) {
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

  const text = new TextRun(" ''Without Prejudice'' ")
  const paragraph = new Paragraph();
  text.bold();
  // document.createParagraph(" ''Without Prejudice'' ").title().heading3();
  paragraph.addRun(text).center();
  document.addParagraph(paragraph);

  document.createParagraph(" ");

  document.createParagraph("Our Ref: DEMAND2/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(DATE);
  const pddate = new Paragraph();
  ddate.size(20);
  pddate.addRun(ddate);
  document.addParagraph(pddate);

  document.createParagraph(" ");
  document.createParagraph(letter_data.custname);
  document.createParagraph(letter_data.address + ' - ' + letter_data.postcode);
  document.createParagraph(" ");

  document.createParagraph("Dear Sir/Madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: OUTSTANDING LIABILITIES A/C NO. " + letter_data.acc + " - " + letter_data.custname + " ");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("Following our 1st notice " + dateFormat(letter_data.demand1date, 'dd-mmm-yyyy') + ", we note with concern that your account/s is/are still in arrears/overdrawn ");
  document.createParagraph("Kindly note that your current balance/s as indicated below and it/they continue/s to accrue interest until payment is made in full.  ");
  document.createParagraph(" ");

  const table = document.createTable(DATA.length + 2, 7);
  /*float({
      horizontalAnchor: TableAnchorType.MARGIN,
      verticalAnchor: TableAnchorType.MARGIN,
      relativeHorizontalPosition: RelativeHorizontalPosition.RIGHT,
      relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
  });*/
  // table.setFixedWidthLayout();
  // table.setWidth('45', WidthType.DXA);
  table.getCell(0, 1).addContent(new Paragraph("Account no"));
  table.getCell(0, 2).addContent(new Paragraph("Principal Loan"));
  table.getCell(0, 3).addContent(new Paragraph("Outstanding Interest "));
  table.getCell(0, 4).addContent(new Paragraph("Principal Arrears"));
  table.getCell(0, 5).addContent(new Paragraph("Total Arrears"));
  table.getCell(0, 6).addContent(
    new Paragraph("Total Outstanding")
  );
  // table rows
  for (i = 0; i < DATA.length; i++) {
    row = i + 1
    table.getCell(row, 1).addContent(new Paragraph(DATA[i].accnumber));
    table.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
    table.getCell(row, 3).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 4).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 5).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].totalarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 6).addContent(new Paragraph(DATA[i].currency.concat(Math.abs(Number.parseInt(DATA[i].oustbalance, 10)) + (Math.abs(Number.parseInt(DATA[i].oustbalance, 10))))));
  }

  document.createParagraph(" ");

  document.createParagraph(" ");
  document.createParagraph("We DEMAND that you pay the amount in arrears, plus the accrued interest within fourteen days (14) from the date hereof. ");


  document.createParagraph(" ");
  const txt = new TextRun("Kindly also note that under the provisions of the Banking (Credit Reference Bureau) Regulations 2013, it is now a mandatory requirement in law that all financial institutions share positive and negative credit information while assessing customers credit worthiness, standing and capacity through duly licensed Credit Reference Bureaus (CRBs) for inclusion and maintenance in their database for purposes of sharing the said information..");
  const ptxt = new Paragraph();
  txt.size(20);
  ptxt.addRun(txt);
  ptxt.justified();
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt2 = new TextRun("We would therefore as a matter of courtesy like to notify you that unless you fully settle all your outstanding arrears with the Bank from the date stated above we shall proceed to adversely update your details and information with the CRBs relating to your credit worthiness and standing.  ");
  const ptxt2 = new Paragraph();
  txt2.size(20);
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("In the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000 ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(letter_data.manager);
  document.createParagraph("BRANCH MANAGER ");
  document.createParagraph(letter_data.branchname);


  if (GUARANTORS.length > 0) {
    document.createParagraph("cc: ");

    for (g = 0; g < GUARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GUARANTORS[g].name);
      document.createParagraph(GUARANTORS[g].address);
    }
  }

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph(" ");
  const bottom = new TextRun("This letter is electronically generated and is valid without a signature ");
  const pbottom = new Paragraph();
  bottom.italic()
  pbottom.addRun(bottom);
  document.addParagraph(pbottom);


  const packer = new Packer();

  try {
    const buffer = await packer.toBuffer(document);
    fs.writeFileSync(LETTERS_DIR + accnumber_masked + DATE + "demand2.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {
      var body = [];
      body.push(['Account Number', 'Principal Loan', 'Outstanding Interest', 'Principal Arrears', 'Total Arrears', 'Total Outstanding'])
      for (i = 0; i < DATA.length; i++) {
        row = i + 1
        body.push([(DATA[i].accnumber).substring(0, 9) + 'xxxxx',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intratearr)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].instamount)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance + DATA[i].instamount)).format('0,0.00') + ' DR'
        ])
      }

      function guarantors() {
        if (GUARANTORS.length > 0) {
          var inc = "\ncc: \n";
          var guar = ''
          for (g = 0; g < GUARANTORS.length; g++) {
            guar = guar + GUARANTORS[g].guarantorname + '\n' + GUARANTORS[g].address + '\n\n';
          }
          return inc + guar;
        }
      }
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
          {
            text: '\n\n\'\'Without Prejudice\'\'\n\n',
            alignment: 'center',
            bold: true
          },
          'Our Ref: DEMAND2/' + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE,
          '\n' + DATE,
          '\n' + letter_data.custname,
          letter_data.address + '-' + letter_data.postcode,

          '\nDear Sir/Madam',

          {
            text: 'RE: OUTSTANDING LIABILITIES A/C NO. ' + accnumber_masked + ' - ' + letter_data.custname + '',
            style: 'subheader'
          },

          '\nFollowing our 1st notice ' + dateFormat(letter_data.demand1date, 'dd-mmm-yyyy') + ', we note with concern that your account/s is/are still in arrears/overdrawn Kindly note that your current balance/s as indicated below and it/they continue/s to accrue interest until payment is made in full.',
          '\n',
          {
            alignment: 'justify',
            fontSize: 9,
            table: {
              body: body
            },
            layout: {
              fillColor: function (rowIndex, node, columnIndex) {
                return (rowIndex % 2 === 0) ? '#CCCCCC' : null;
              }
            }
          },

          { text: '\nWe DEMAND that you pay the amount in arrears, plus the accrued interest within fourteen days (14) from the date hereof. ', fontSize: 10 },

          { text: '\nKindly also note that under the provisions of the Banking (Credit Reference Bureau) Regulations 2013, it is now a mandatory requirement in law that all financial institutions share positive and negative credit information while assessing customers credit worthiness, standing and capacity through duly licensed Credit Reference Bureaus (CRBs) for inclusion and maintenance in their database for purposes of sharing the said information.', fontSize: 10 },

          { text: '\nWe would therefore as a matter of courtesy like to notify you that unless you fully settle all your outstanding arrears with the Bank from the date stated above we shall proceed to adversely update your details and information with the CRBs relating to your credit worthiness and standing. ', fontSize: 10 },

          { text: '\nIn the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000 ', fontSize: 10 },

          { text: '\nYours Faithfully, ' },
          { text: '\n\nBRANCH MANAGER , ' },
          { text: letter_data.branchname },
          { text: '\n\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },

          guarantors()

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
      writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "demand2.pdf");
      pdfDoc.pipe(writeStream);
      pdfDoc.end();
      writeStream.on('finish', async function () {
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "demand2.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "demand2.pdf"
        var metaData = {
          'Content-Type': 'text/html',
          'Content-Language': 123,
          'X-Amz-Meta-Testing': 1234,
          'example': 5678
        }
        
        const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

        res.json({
          result: 'success',
          message: LETTERS_DIR + accnumber_masked + DATE + "demand2.pdf",
          filename: accnumber_masked + DATE + "demand2.pdf",
          savedfilename: savedfilename,
          objInfo: objInfo
        })
        //save to mino end
      });
    } else {
      // save to minio
      const filelocation = LETTERS_DIR + accnumber_masked + DATE + "demand2.docx";
      const bucket = 'demandletters';
      const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "demand2.docx"
      var metaData = {
        'Content-Type': 'text/html',
        'Content-Language': 123,
        'X-Amz-Meta-Testing': 1234,
        'example': 5678
      }
      const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

      res.json({
        result: 'success',
        message: LETTERS_DIR + accnumber_masked + DATE + "demand2.docx",
        filename: accnumber_masked + DATE + "demand2.docx",
        savedfilename: savedfilename,
        objInfo: objInfo
      })

      //save to mino end
    }
  } catch (error) {
    console.log(err);
    res.status(500).json({
      result: 'error',
      message: err.message
    });
  }


});
module.exports = router;
