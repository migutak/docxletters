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

    const text = new TextRun(" ''Without Prejudice'' ")
    const paragraph = new Paragraph();
    text.bold();
    // document.createParagraph(" ''Without Prejudice'' ").title().heading3();
    paragraph.addRun(text).center();
    document.addParagraph(paragraph);

    document.createParagraph(" ");

    document.createParagraph("Our Ref: PRELISTING/"+ letter_data.branchcode + '/' + letter_data.arocode +'/'+ DATE);
    document.createParagraph(" ");
    const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
    const pddate = new Paragraph();
    ddate.size(20);
    pddate.addRun(ddate);
    document.addParagraph(pddate);

    document.createParagraph(" ");
    document.createParagraph(letter_data.custname);
    document.createParagraph(letter_data.address);
    document.createParagraph(letter_data.custname);
    document.createParagraph(" ");

    document.createParagraph("Dear sir/madam ");
    document.createParagraph(" ");

    const headertext = new TextRun("PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1) (a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013:");
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

    const table = document.createTable(DATA.length + 2, 7);
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
        table.getCell(row, 2).addContent(new Paragraph(numeral(DATA[i].oustbalance).format('0,0.00')));
        table.getCell(row, 3).addContent(new Paragraph(numeral(DATA[i].princarrears).format('0,0.00')));
        table.getCell(row, 4).addContent(new Paragraph(numeral(DATA[i].intarrears).format('0,0.00')));
        table.getCell(row, 5).addContent(new Paragraph(numeral(DATA[i].totalarrears).format('0,0.00')));
        table.getCell(row, 6).addContent(new Paragraph(numeral(DATA[i].oustbalance + DATA[i].totalarrears).format('0,0.00')));
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
    const txt2 = new TextRun("Kindly make the necessary arrangements to repay the outstanding balance within the next Fourteen (14) days from the date of this letter, i.e. on or before (Date), failure to which we shall have no option but to exercise any of the remedies below against you, to recover the said outstanding amount at your risk as to costs and expenses arising without further reference to you;.");
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

  const crb7 = new TextRun("Please text your name to 21272 and                                                  ");
  const pcrb7 = new Paragraph();
  crb7.size(20);
  pcrb7.addRun(crb7);
  document.addParagraph(pcrb7);

  const crb8 = new TextRun("follow instructions to secure a copy of your CRB report.                                                 ");
  const pcrb8 = new Paragraph();
  crb8.size(20);
  pcrb8.addRun(crb8);
  document.addParagraph(pcrb8);
  //stop crb

    document.createParagraph(" ");
    document.createParagraph("Yours Faithfully, ");

    document.createParagraph(" ");
    document.createParagraph( letter_data.manager );
    document.createParagraph("BRANCH MANAGER ");
    document.createParagraph( letter_data.branchname + " BRANCH");


    if (GURARANTORS) {
        document.createParagraph("cc: ");

        for (g = 0; g < GURARANTORS.length; g++) {
            document.createParagraph(" ");
            document.createParagraph( GURARANTORS[g].name );
            document.createParagraph( GURARANTORS[g].address );
        }
    }

    document.createParagraph(" ");
    document.createParagraph("This letter is valid without a signature ");

    const packer = new Packer();

    packer.toBuffer(document).then((buffer) => {
        fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "prelisting.docx", buffer);
        //conver to pdf
        // if pdf format
        if (letter_data.format == 'pdf') {
          const convert = () => {
            word2pdf.word2pdf(LETTERS_DIR + letter_data.cardacct + DATE + "prelisting.docx")
              .then(data => {
                fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + 'prelisting.pdf', data);
                res.json({
                  result: 'success',
                  message: LETTERS_DIR + letter_data.cardacct + DATE + "prelisting.pdf",
                  filename: letter_data.acc + DATE + "prelisting.pdf"
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
          // res.sendFile(path.join(LETTERS_DIR + letter_data.cardacct + DATE + 'prelisting.docx'));
          res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.cardacct + DATE + "prelisting.docx",
            filename: letter_data.acc + DATE + "prelisting.docx"
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