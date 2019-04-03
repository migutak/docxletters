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
const IMAGEPATH = data.imagePath;

const {
    Document,
    Paragraph,
    Packer,
    TextRun,
    HorizontalPositionRelativeFrom,
    HorizontalPositionAlign
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
    const GUARANTORS = req.body.guarantors;
    const INCLUDELOGO = req.body.showlogo;
    const DATA = req.body.accounts;
    const DATE = dateFormat(new Date(), "dd-mmm-yyyy");
    //
    // console.log(letter_data);
    //
    const document = new Document();

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
    if (INCLUDELOGO == true) {
        document.createImage(fs.readFileSync(IMAGEPATH + "coop.jpg"), 350, 60, {
            floating: {
                behindDocument: true,
                horizontalPosition: {
                    relative: HorizontalPositionRelativeFrom.LEFT_MARGIN,
                    // align: HorizontalPositionAlign.LEFT
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

    const text = new TextRun(" ''Without Prejudice'' ")
    const paragraph = new Paragraph();
    text.bold();
    paragraph.addRun(text).center();
    document.addParagraph(paragraph);

    document.createParagraph(" ");

    document.createParagraph("Our Ref: DEMAND1/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
    document.createParagraph(" ");
    const ddate = new TextRun(DATE);
    const pddate = new Paragraph();
    ddate.size(20);
    pddate.addRun(ddate);
    document.addParagraph(pddate);

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

    const headertext = new TextRun("RE: OUTSTANDING LIABILITIES A/C NO. " + letter_data.acc + " - " + letter_data.custname + " ");
    const paragraphheadertext = new Paragraph();
    headertext.bold();
    headertext.underline();
    paragraphheadertext.addRun(headertext);
    document.addParagraph(paragraphheadertext);

    document.createParagraph(" ");
    document.createParagraph("We write to notify you that your account/s is currently in arrears/overdrawn. Kindly note that your current balance is as indicated below and it continues to accrue interest until payment is made in full. ");
    document.createParagraph(" ");

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
        table.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
        table.getCell(row, 3).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR'));
        table.getCell(row, 4).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intarrears)).format('0,0.00') + ' DR'));
        table.getCell(row, 5).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].totalarrears)).format('0,0.00') + ' DR'));
        table.getCell(row, 6).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance + DATA[i].totalarrears)).format('0,0.00') + ' DR'));
    }

    document.createParagraph(" ");

    document.createParagraph(" ");
    document.createParagraph("The purpose of this letter therefore is to DEMAND immediate payment for the amount in arrears. Kindly ensure that the same is paid within fourteen days (14) from the date hereof. ");


    document.createParagraph(" ");
    document.createParagraph("In the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000. ");

    document.createParagraph(" ");
    document.createParagraph("Yours Faithfully, ");

    document.createParagraph(" ");
    document.createParagraph(letter_data.manager);
    document.createParagraph("BRANCH MANAGER ");
    document.createParagraph(letter_data.branchname);

    document.createParagraph(" ");

    if (GUARANTORS.length > 0) {
        document.createParagraph("CC: ");

        for (g = 0; g < GUARANTORS.length; g++) {
            document.createParagraph(" ");
            document.createParagraph(GUARANTORS[g].guarantorname);
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
        fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "demand1.docx", buffer);
        //conver to pdf
        // if pdf format
        if (letter_data.format == 'pdf') {
            const convert = () => {
                word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "demand1.docx")
                    .then(data => {
                        fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + 'demand1.pdf', data);
                        res.json({
                            result: 'success',
                            message: LETTERS_DIR + letter_data.acc + DATE + "demand1.pdf",
                            filename: letter_data.acc + DATE + "demand1.pdf"
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
            // res.sendFile(path.join(LETTERS_DIR + letter_data.acc + DATE + 'demand1.docx'));
            res.json({
                result: 'success',
                message: LETTERS_DIR + letter_data.acc + DATE + "demand1.docx",
                filename: letter_data.acc + DATE + "demand1.docx"
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
