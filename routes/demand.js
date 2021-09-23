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


router.post('/download', function (req, res) {
    const letter_data = req.body;
    const GURARANTORS = req.body.guarantors;
    const INCLUDELOGO = req.body.showlogo;
    const DATA = req.body.accounts;
    const DATE = dateFormat(new Date(), "isoDate");

    const document = new Document();
    if (INCLUDELOGO == 'Y') {
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

        document.createImage(fs.readFileSync("./coop.jpg"), 350, 60, {
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

    document.createParagraph(" ");

    const text = new TextRun(" ''Without Prejudice'' ")
    const paragraph = new Paragraph();
    text.bold();
    // document.createParagraph(" ''Without Prejudice'' ").title().heading3();
    paragraph.addRun(text).center();
    document.addParagraph(paragraph);

    document.createParagraph(" ");

    document.createParagraph("Our Ref: DEMAND1/" + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE);
    document.createParagraph(" ");
    const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
    const pddate = new Paragraph();
    ddate.size(20);
    pddate.addRun(ddate);
    document.addParagraph(pddate);

    // 


    document.createParagraph(" ");
    document.createParagraph("This letter is valid without a signature ");

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
})


module.exports = router;