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
  port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
  useSSL: false,
  accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
  secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});

const { Document, Paragraph, Packer, TextRun } = docx;

var data = require('./data.js');

const LETTERS_DIR = data.filePath;
const IMAGEPATH = data.imagePath;

router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());

router.post('/download', async function (req, res) {
  const letter_data = req.body;
  const GURARANTORS = req.body.guarantors || [];
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

  document.createParagraph("Our Ref: DAY90/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
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
  document.createParagraph("Dear Sir/Madam ");
  document.createParagraph(" ");

  const headertext = new TextRun("RE: OUTSTANDING LIABILITIES DUE TO THE BANK ON ACCOUNT OF " + letter_data.acc + ": BASE NO. " + letter_data.custnumber);
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("We refer to our notice dated " + dateFormat(new Date(), 'fullDate') + ".");

  document.createParagraph(" ");
  const txt3 = new TextRun("As you are fully aware and despite the referenced notice, the above account is in arrears of " + letter_data.accounts[0].currency + ' ' + numeral(letter_data.accounts[0].oustbalance).format('0,0.0') + " dr as at (Date) which continues to accrue interest at xxx% per annum (equivalent to Kenya Bank's Reference Rate (KBRR) currently at xxxx% plus a margin of xxx% (K)) and late penalties of 0.5% per month and further the total outstanding sum due to the Bank as at " + DATE + " is " + letter_data.accounts[0].currency + ' ' + numeral(letter_data.accounts[0].oustbalance).format('0,0.0') + ". dr which continues to accrue interest at xxx% per annum (equivalent to Kenya Bank's Reference Rate (KBRR) currently at xxxx% plus a margin of xxx% (K)).");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("The liabilities are secured by way of a Legal charge over the properties: ");
  document.createParagraph("L.R.NO. " + letter_data.lrno.toUpperCase() + " registered in the name of " + letter_data.regowner.toUpperCase() + ".")
  document.createParagraph(" ");


  document.createParagraph("TAKE NOTICE that pursuant to the provisions of Section 90 of the Land Act, 2012, the Bank intends to take action and exercise remedies provided in this Section after the expiry of THREE (3) MONTHS from the date of service of this Notice upon yourself if you do not rectify the default by repaying the outstanding sum of " + letter_data.accounts[0].currency + ' ' + numeral(letter_data.accounts[0].oustbalance).format('0,0.0') + "dr which includes the ");
  document.createParagraph(" ");
  document.createParagraph("Please be advised that if you fail to remedy the default and repay the outstanding amount as stated above the Bank shall exercise any of the remedies as stipulated in Section 90 (3) of the Land Act, 2012 against you which includes:");
  document.createParagraph("•	Files suit against you for money due and owing ");
  document.createParagraph("•	Appoint a receiver of the income of the charged property ");
  document.createParagraph("•	Lease or sublease the charged property ");
  document.createParagraph("•	Enter into possession of the charged Property ");
  document.createParagraph("•	Sell the charged property. ");

  document.createParagraph(" ");
  document.createParagraph("FURTHER NOTE that pursuant to the provisions of Sections 90(2) (e) and 103 of the Land Act, 2012, you are at liberty to apply to the Court for any relief that the Court may deem fit to grant against the Bank's remedies. ");


  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph("Yours Faithfully, ");

  document.createParagraph(" ");
  document.createParagraph(" ");
  document.createParagraph("                                                                                                    JOYCE MBINDA");
  document.createParagraph(letter_data.arocode + "                                                                                     MANAGER REMEDIAL MANAGEMENT");
  document.createParagraph("CREDIT MANAGEMENT DIVISION.                                    CREDIT MANAGEMENT DIVISION.");


  if (GURARANTORS.length > 0) {
    document.createParagraph("CC: ");

    for (g = 0; g < GURARANTORS.length; g++) {
      document.createParagraph(" ");
      document.createParagraph(GURARANTORS[g].guarantorname);
      document.createParagraph(GURARANTORS[g].address);
    }
  }

  const packer = new Packer();

  try {
    const buffer = await packer.toBuffer(document);
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "day90.docx", buffer);
    // save to minio
    const filelocation = LETTERS_DIR + letter_data.acc + DATE + "day90.docx";
    const bucket = 'demandletters';
    const savedfilename = letter_data.acc + '_' + Date.now() + '_' + "day90.docx"
    var metaData = {
      'Content-Type': 'text/html',
      'Content-Language': 123,
      'X-Amz-Meta-Testing': 1234,
      'example': 5678
    }
    const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

    res.json({
      result: 'success',
      message: LETTERS_DIR + letter_data.acc + DATE + "day90.docx",
      filename: letter_data.acc + DATE + "day90.docx",
      savedfilename: savedfilename,
      objInfo: objInfo
    })

    //save to mino end
  } catch (error) {
    console.log(error);
    res.status(500).json({
      success: false,
      error: error.message
    })
  }

});

module.exports = router;
