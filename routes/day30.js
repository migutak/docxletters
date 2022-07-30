var express = require('express');
var router = express.Router();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const word2pdf = require('word2pdf-promises');
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

var data = require('./data.js');

const LETTERS_DIR = data.filePath;
const IMAGEPATH = data.imagePath;

const { Document, Paragraph, Packer, TextRun } = docx;

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

  document.createParagraph("Our Ref: DAY30/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'dd-mmm-yyyy'));
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

  const headertext = new TextRun("RE: DEMAND FOR OUTSTANDING BALANCES DUE ON LOAN FACILITIES ON ACCOUNT	" + letter_data.custname + " NO. " + letter_data.cust + " AND DEFAULT LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1)(b) OF THE BANKING (CRB) REGULATIONS, 2013");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  document.createParagraph("The above matter refers.");

  document.createParagraph(" ");
  const txt3 = new TextRun("We wish to notify you have breached terms of the letters of offer by defaulting in repaying your monthly repayments and as such your account is in arrears of " + letter_data.accounts[0].currency + ' ' + numeral(Math.abs(letter_data.accounts[0].oustbalance)).format('0,0.0') + "DR as at " + DATE + " which continues to accrue interest at 13% per annum (equivalent to Central Bank Rate (CBR) (currently at 9%) plus a margin of 4% per annum. Further, you owe the Bank the total sum of " + letter_data.accounts[0].currency + ' ' + numeral(Math.abs(letter_data.accounts[0].oustbalance)).format('0,0.0') + "DR as at " + DATE + " being the outstanding amount on the facility, which continues to accrue interest at 13% per annum (equivalent to Central Bank Rate (CBR) (currently at 9%) plus a margin of 4% per annum) until full payment, full particulars whereof are well within your knowledge. ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  ptxt3.justified();
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  document.createParagraph("Please note that if full payment of the outstanding amount is not made within the next Thirty (30) days from the date of this letter, then we shall take the necessary action to protect the Bank's interest at your own risk as to costs");
  document.createParagraph(" ");


  document.createParagraph("After revisions in 2012/2013 to the Banking Act (Cap 488), Central Bank Act, Microfinance Act, 2006 and the CRB Regulations, Banks and Microfinance Banks have been mandated to share information on all their borrowers, and their loan information with registered Credit Reference Bureaus (CRBs). This means that the CRBs will now hold information on both good and bad borrowers. A good loan repayment pattern will reflect in a borrower's credit report resulting in an attractive credit profile, which can allow a borrower to negotiate preferential loan agreements with lenders.  ");
  document.createParagraph(" ");
  document.createParagraph("Thus, in compliance to the law, and having borrowed with Co-operative Bank of Kenya Limited, we have forwarded your information to the Credit Reference Bureaus below ");

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

  const crb5 = new TextRun("Fax: +254(0)20 3751344                                Fax: +254 (0) 20273572 ");
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
  const txt = new TextRun("You are encouraged:");
  const ptxt = new Paragraph();
  txt.size(18);
  ptxt.addRun(txt);
  ptxt.justified();
  document.addParagraph(ptxt);

  const txt4 = new TextRun("• To ensure that your loan payments are always up-to date and ");
  const ptxt4 = new Paragraph();
  txt4.size(18);
  ptxt4.addRun(txt4);
  document.addParagraph(ptxt4);

  const txt5 = new TextRun("• Regularly obtain your credit report from the bureaus above to ascertain the accuracy of your information. ");
  const ptxt5 = new Paragraph();
  txt5.size(18);
  ptxt5.addRun(txt5);
  document.addParagraph(ptxt5);

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
      document.createParagraph(GURARANTORS[g].name);
      document.createParagraph(GURARANTORS[g].address);
    }
  }

  const packer = new Packer();

  try {
    const buffer = await packer.toBuffer(document);
    fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "day30.docx", buffer);
    // save to minio
    const filelocation = LETTERS_DIR + letter_data.acc + DATE + "day30.docx";
    const bucket = 'demandletters';
    const savedfilename = letter_data.acc + '_' + Date.now() + '_' + "day30.docx"
    var metaData = {
      'Content-Type': 'text/html',
      'Content-Language': 123,
      'X-Amz-Meta-Testing': 1234,
      'example': 5678
    }
    const objInfo = await minioClient.fPutObject(bucket, savedfilename, filelocation, metaData);

    res.json({
      result: 'success',
      message: LETTERS_DIR + letter_data.acc + DATE + "day30.docx",
      filename: letter_data.acc + DATE + "day30.docx",
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
