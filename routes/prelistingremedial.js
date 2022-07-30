var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const word2pdf = require('word2pdf-promises');
const cors = require('cors')
require('log-timestamp');

var Minio = require("minio");

var minioClient = new Minio.Client({
  endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
  port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
  useSSL: false,
  accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
  secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});

var data = require('./data.js');

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
  const GURARANTORS = req.body.guarantors;
  const INCLUDELOGO = req.body.showlogo;
  const DATA = req.body.accounts;
  const DATE = dateFormat(new Date(), "dd-mmm-yyyy");
  //
  //
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

  const nametext = new TextRun(letter_data.custname);
  const pnametext = new Paragraph();
  nametext.allCaps();
  pnametext.addRun(nametext);
  document.addParagraph(pnametext);

  const addresstext = new TextRun(letter_data.address + '-' + letter_data.postcode);
  const paddresstext = new Paragraph();
  addresstext.allCaps();
  paddresstext.addRun(addresstext);
  document.addParagraph(paddresstext);

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
  table.getCell(0, 6).addContent(
    new Paragraph("Penalty Interest")
  );
  table.getCell(0, 7).addContent(
    new Paragraph("Days in arrears")
  );
  table.getCell(0, 8).addContent(
    new Paragraph("Interest Rate per annum")
  );
  // table rows
  for (i = 0; i < DATA.length; i++) {
    row = i + 1
    table.getCell(row, 1).addContent(new Paragraph(DATA[i].accnumber));
    table.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
    table.getCell(row, 3).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 4).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 5).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].totalarrears)).format('0,0.00') + ' DR'));
    table.getCell(row, 6).addContent(new Paragraph(DATA[i].currency + ' 0.00'));
    table.getCell(row, 7).addContent(new Paragraph('Over 60 days'));
    table.getCell(row, 8).addContent(new Paragraph('14%'));
  }



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
  //stop crb

  document.createParagraph(" ");
  const txt2 = new TextRun("In need, please feel free to contact the undersigned at Credit Management Division, Co-operative House Building, and Mezzanine 2.Tel:020-3276122/0711049122/0732106122.");
  const ptxt2 = new Paragraph();
  txt2.size(18);
  ptxt2.addRun(txt2);
  ptxt2.justified();
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph("                                                                                                    JOYCE MBINDA");
  document.createParagraph(letter_data.arocode + "                                                                                     MANAGER REMEDIAL MANAGEMENT");
  document.createParagraph("CREDIT MANAGEMENT DIVISION.                                    CREDIT MANAGEMENT DIVISION.");



  if (GURARANTORS.length > 0) {
    document.createParagraph("cc: ");

    for (g = 0; g < GURARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GURARANTORS[g].guarantorname);
      document.createParagraph(GURARANTORS[g].address);
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
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "prelistingremedial.docx", buffer);
    // if pdf format
    if (letter_data.format == 'pdf') {
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "prelistingremedial.docx")
          .then(data => {
            fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + 'prelistingremedial.pdf', data);
            res.json({
              result: 'success',
              message: LETTERS_DIR + letter_data.acc + DATE + "prelistingremedial.pdf",
              filename: letter_data.acc + DATE + "prelistingremedial.pdf"
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
      // save to minio
      const filelocation = LETTERS_DIR + letter_data.acc + DATE + "prelistingremedial.docx";
      const bucket = 'demandletters';
      const savedfilename = letter_data.acc + '_' + Date.now() + '_' + "prelistingremedial.docx"
      var metaData = {
        'Content-Type': 'text/html',
        'Content-Language': 123,
        'X-Amz-Meta-Testing': 1234,
        'example': 5678
      }
      const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

      res.json({
        result: 'success',
        message: LETTERS_DIR + letter_data.acc + DATE + "prelistingremedial.docx",
        filename: letter_data.acc + DATE + "prelistingremedial.docx",
        savedfilename: savedfilename,
        objInfo: objInfo
      })

      //save to mino end
    }
  } catch (error) {
    console.log(error);
    res.status(500).json({
      success: false,
      error: error.message
    })
  }


});

module.exports = router;
