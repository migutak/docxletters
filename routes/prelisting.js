var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const cors = require('cors');
require('log-timestamp');

var Minio = require("minio");

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

router.post('/download', function (req, res) {
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

  document.createParagraph("Our Ref: PRELISTING/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(DATE);
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
  document.createParagraph(letter_data.address + '-' + letter_data.postcode);
  document.createParagraph(" ");

  document.createParagraph("Dear Sir/Madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("REF: PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1) (a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013:");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("We wish to inform you that, in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs).  ");
  document.createParagraph(" ");
  document.createParagraph("A default in loan repayment will result in a negative impact on your credit record. If your loan is classified as Non-Performing as per the Banking Act & Prudential Guidelines and/or as per the Microfinance Act, your credit profile at the CRBs will be adversely affected.   ");
  document.createParagraph(" ");

  document.createParagraph("Please note that your loans are currently in default with outstanding balances and arrears, having not paid the full instalments. These loans continue to accrue interest at various rates per annum. Here below please find the loan/overdrawn particulars: ");
  document.createParagraph(" ");

  // table 1
  const table1 = document.createTable(DATA.length + 2, 3);
  table1.getCell(0, 1).addContent(new Paragraph("Type of facility"));
  table1.getCell(0, 2).addContent(new Paragraph("Amount (Kshs)"));
  for (i = 0; i < DATA.length; i++) {
    row = i + 1
    table1.getCell(row, 1).addContent(new Paragraph(DATA[i].productcode));
    table1.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
  }
  //end table 1
  document.createParagraph(" ");
  document.createParagraph("This is broken down as follows: ");
  document.createParagraph(" ");
  const table = document.createTable(DATA.length + 2, 9);
  table.getCell(0, 1).addContent(new Paragraph("Loan Account Number"));
  table.getCell(0, 2).addContent(new Paragraph("Principal Loan Balance"));
  table.getCell(0, 3).addContent(new Paragraph("Accrued Interest on principal"));
  table.getCell(0, 4).addContent(new Paragraph("Principal Arrears"));
  table.getCell(0, 5).addContent(new Paragraph("Interest in Arrears"));
  table.getCell(0, 6).addContent(new Paragraph("Penalty Interest"));
  table.getCell(0, 7).addContent(new Paragraph("Days in arrears"));
  table.getCell(0, 8).addContent(new Paragraph("Interest Rate per annum"));
  // table rows
  for (i = 0; i < DATA.length; i++) {
    row = i + 1
    table.getCell(row, 1).addContent(new Paragraph(DATA[i].accnumber));
    table.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
    table.getCell(row, 3).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 4).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 5).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].totalarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 6).addContent(new Paragraph(DATA[i].currency + ' 0.00'));
    table.getCell(row, 7).addContent(new Paragraph('Over 60 days'));
    table.getCell(row, 8).addContent(new Paragraph('14%'));
  }

  document.createParagraph(" ");
  document.createParagraph("Please note that interest continues to accrue at various Bank rates until the outstanding balance is paid in full.. ");


  document.createParagraph(" ");
  const txt = new TextRun("Kindly also note that under the provisions of the Banking (Credit Reference Bureau) Regulations 2013, it is now a mandatory requirement in law that all financial institutions share positive and negative credit information while assessing customers credit worthiness, standing and capacity through duly licensed Credit Reference Bureaus (CRBs) for inclusion and maintenance in their database for purposes of sharing the said information..");
  const ptxt = new Paragraph();
  txt.size(20);
  ptxt.addRun(txt);
  ptxt.justified();
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt2 = new TextRun("Kindly make the necessary arrangements to repay the outstanding balance within the next Fourteen (14) days from the date of this letter, i.e. on or before " + dateFormat(new Date() + 14, "dd-mmm-yyyy") + ", failure to which we shall have no option but to exercise any of the remedies below against you, to recover the said outstanding amount at your risk as to costs and expenses arising without further reference to you;.");
  const ptxt2 = new Paragraph();
  txt2.size(20);
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("We hereby notify you that we will proceed to adversely list you with the CRBs if your loan (s) becomes non-performing. To avoid an adverse listing, you are advised to clear the outstanding arrears.  ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  const txt4 = new TextRun("You have a right of access to your credit report at the CRBs and you may dispute any erroneous information. You may request for your report by contacting the CRBs at the following addresses:.  ");
  const ptxt4 = new Paragraph();
  txt4.size(20);
  ptxt4.addRun(txt4);
  ptxt4.justified();
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
  //st

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

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + accnumber_masked + DATE + "prelisting.docx", buffer);
    //conver to pdf
    // if pdf format
    if (letter_data.format == 'pdf') {

      var body = [];
      var body2 = [];
      body2.push(['Type of facility', 'Amount (Kshs)'])
      for (i = 0; i < DATA.length; i++) {
        row = i + 1
        body2.push([DATA[i].productcode,
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'
        ])
      }

      body.push(['Loan Account Number', 'Principal Loan Balance', 'Accrued Interest on principal', 'Principal Arrears', 'Interest in Arrears', 'Penalty Interest', 'Days in arrears', 'Interest Rate per annum'])
      for (i = 0; i < DATA.length; i++) {
        row = i + 1
        body.push([(DATA[i].accnumber).substring(0, 9) + 'xxxxx',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intratearr)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR',
        DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].instamount)).format('0,0.00') + ' DR',
        DATA[i].currency + ' 0.00',
          'Over 60 days',
          '14%'
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
        pageMargins: [50, 50, 50, 50],
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
                  'www.co-opbank.co.ke'
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
          'Our Ref: PRELISTING/' + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE,
          '\n' + DATE,
          {
            type: 'none',
            alignment: 'right',
            fontSize: 9,
            ol: [
              'BY REGISTERED POST',
              'Copy by ordinary Mail'
            ]
          },

          '\n' + letter_data.custname,
          letter_data.address + '-' + letter_data.postcode,

          '\nDear Sir/Madam',

          {
            text: '\nREF: PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1) (a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013:',
            style: 'subheader'
          },

          "\nWe wish to inform you that, in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs).  ",

          "\nA default in loan repayment will result in a negative impact on your credit record. If your loan is classified as Non-Performing as per the Banking Act & Prudential Guidelines and/or as per the Microfinance Act, your credit profile at the CRBs will be adversely affected.   ",

          "\nPlease note that your loans are currently in default with outstanding balances and arrears, having not paid the full instalments. These loans continue to accrue interest at various rates per annum. Here below please find the loan/overdrawn particulars: ",

          '\n',

          {
            alignment: 'left',
            fontSize: 9,
            table: {
              body: body2
            }
          },
          '\n',
          '\nThis is broken down as follows: ',
          '\n',

          {
            alignment: 'left',
            fontSize: 9,
            table: {
              body: body
            }
          },

          { text: '\nPlease note that interest continues to accrue at various Bank rates until the outstanding balance is paid in full. ', fontSize: 10 },

          { text: '\nKindly also note that under the provisions of the Banking (Credit Reference Bureau) Regulations 2013, it is now a mandatory requirement in law that all financial institutions share positive and negative credit information while assessing customers credit worthiness, standing and capacity through duly licensed Credit Reference Bureaus (CRBs) for inclusion and maintenance in their database for purposes of sharing the said information', fontSize: 10 },

          { text: '\nKindly make the necessary arrangements to repay the outstanding balance within the next Fourteen (14) days from the date of this letter, i.e. on or before 23-Sep-2019, failure to which we shall have no option but to exercise any of the remedies below against you, to recover the said outstanding amount at your risk as to costs and expenses arising without further reference to you;', fontSize: 10 },

          { text: '\nWe hereby notify you that we will proceed to adversely list you with the CRBs if your loan (s) becomes nonperforming. To avoid an adverse listing, you are advised to clear the outstanding arrears.', fontSize: 10 },

          { text: '\nYou have a right of access to your credit report at the CRBs and you may dispute any erroneous information. You may request for your report by contacting the CRBs at the following addresses:', fontSize: 10 },

          '\n',

          {
            columns: [
              {
                type: 'none',
                ol: [
                  'TransUnion CRB',
                  'Delta Corner Annex, Ring Road,',
                  'Westlands Lane, Nairobi.',
                  'P.O. Box 46406, 00100, NAIROBI,',
                  'Tel: +254 (0) 20 51799/3751360/2/4/5',
                  'Fax:+254(0)20 3751344',
                  'Email: info@crbafrica.com',
                  'www.crbafrica.com'
                ]
              },
              {
                type: 'none',
                ol: [
                  'Metropol CRB',
                  '1st Floor, Shelter Afrique Centre,',
                  'Upper Hill, Nairobi',
                  'P.O Box 35331–00200 , NAIROBI',
                  'Tel: +254(0)20 2689881/27113575',
                  'Fax: +254 (0) 20273572',
                  'Email: creditbureau@metropol.co.ke',
                  'www.metropolcorporation.com'
                ]
              },
              {
                type: 'none',
                ol: [
                  'Creditinfo CRB Kenya Ltd',
                  'Park Suites, Office 12, Second Floor,',
                  'Parklands Road, Nairobi.',
                  'P.O. Box 38941-00623, NAIROBI ',
                  'Tel: 020 3757272',
                  'Fax: +254 (0) 20273572',
                  'Email: cikinfo@creditinfo.co.ke',
                  ' Website: ke.creditinfo.com'
                ]
              }
            ]
          },

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
      };

      var options = {
        // ...
      }

      var pdfDoc = printer.createPdfKitDocument(dd, options);
      //pdfDoc.pipe(fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "prelisting.pdf"));
      //pdfDoc.end();
      // ensures response is sent only after pdf is created
      writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "prelisting.pdf");
      pdfDoc.pipe(writeStream);
      pdfDoc.end();
      writeStream.on('finish', function () {
        // do stuff with the PDF file
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "prelisting.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "prelisting.pdf"
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
                message: LETTERS_DIR + accnumber_masked + DATE + "prelisting.pdf",
                filename: accnumber_masked + DATE + "prelisting.pdf",
                savedfilename: savedfilename,
                objInfo: objInfo
            })
        });
        //save to mino end
      });
    } else {
      // save to minio
      const filelocation = LETTERS_DIR + accnumber_masked + DATE + "prelisting.docx";
      const bucket = 'demandletters';
      const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "prelisting.docx"
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
          message: LETTERS_DIR + accnumber_masked + DATE + "prelisting.docx",
          filename: accnumber_masked + DATE + "prelisting.docx",
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
