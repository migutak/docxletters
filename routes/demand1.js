var express = require('express');
var router = express.Router();
const app = express();
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const cors = require('cors');

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
    TextRun,
    HorizontalPositionRelativeFrom,
    HorizontalPositionAlign
} = docx;

router.use(bodyParser.urlencoded({
    extended: true
}));

router.use(bodyParser.json());
router.use(cors());


router.get('/', function (req, res) {
    res.json({ message: 'Demand 1 letter is ready! on app01' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    const GUARANTORS = req.body.guarantors;
    const INCLUDELOGO = req.body.showlogo;
    const DATA = req.body.accounts;
    const DATE = dateFormat(new Date(), "dd-mmm-yyyy");

    let NOTICE = 'fourteen days (14)';
    //
    //
    const rawaccnumber = letter_data.acc;
    const memo = rawaccnumber.substr(2, 3);
    const first4 = rawaccnumber.substring(0, 9);
    const last4 = rawaccnumber.substring(rawaccnumber.length - 4);

    const mask = rawaccnumber.substring(4, rawaccnumber.length - 4).replace(/\d/g, "*");
    accnumber_masked = first4 + 'xxxxx';
    //
    if (memo == '6D0' || memo == '6E2' || memo == '6E3') {
        NOTICE = 'seven days (7)';
    };
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
        document.createImage(fs.readFileSync("coop.jpg"), 350, 60, {
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
    document.createParagraph("P.O. Box 48231-00100 GPO, Nairobi").right();
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

    const headertext = new TextRun("RE: OUTSTANDING LIABILITIES A/C NO. " + accnumber_masked + " - " + letter_data.custname + " ");
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
    table.getCell(0, 1).addContent(new Paragraph("Account Number"));
    table.getCell(0, 2).addContent(new Paragraph("Total Outstanding Amount"));
    table.getCell(0, 3).addContent(new Paragraph("Interest Arrears"));
    table.getCell(0, 4).addContent(new Paragraph("Principal Arrears"));
    table.getCell(0, 5).addContent(new Paragraph("Total Arrears"));
    table.getCell(0, 6).addContent(new Paragraph("Total Outstanding"));
    // table rows
    for (i = 0; i < DATA.length; i++) {
        row = i + 1
        table.getCell(row, 1).addContent(new Paragraph((DATA[i].accnumber).substring(0, 9) + 'xxxxx'));
        table.getCell(row, 2).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance)).format('0,0.00') + ' DR'));
        table.getCell(row, 3).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].intratearr)).format('0,0.00') + ' DR'));
        table.getCell(row, 4).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].princarrears)).format('0,0.00') + ' DR'));
        table.getCell(row, 5).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].instamount)).format('0,0.00') + ' DR'));
        table.getCell(row, 6).addContent(new Paragraph(DATA[i].currency + ' ' + numeral(Math.abs(DATA[i].oustbalance + DATA[i].instamount)).format('0,0.00') + ' DR'));
    }

    document.createParagraph(" ");

    document.createParagraph(" ");
    document.createParagraph("The purpose of this letter therefore is to DEMAND immediate payment for the amount in arrears. Kindly ensure that the same is paid within " + NOTICE + " from the date hereof. ");


    document.createParagraph(" ");
    document.createParagraph("In the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000. ");

    document.createParagraph(" ");
    document.createParagraph("Yours Faithfully, ");

    document.createParagraph(" ");
    document.createParagraph("BRANCH MANAGER ");
    document.createParagraph(letter_data.branchname);

    document.createParagraph(" ");

    if (GUARANTORS.length > 0) {
        document.createParagraph("cc: ");

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
        fs.writeFileSync(LETTERS_DIR + accnumber_masked + DATE + "demand1.docx", buffer);

        //conver to pdf
        // if pdf format
        if (letter_data.format == 'pdf') {
            /*const convert = () => {
               word2pdf.word2pdf(path.join(LETTERS_DIR + accnumber_masked + DATE + "demand1.docx"))
                    .then(data => {
                        fs.writeFileSync(LETTERS_DIR + accnumber_masked + DATE + 'demand1.pdf', data);
                        
                        // pipe to remote
                        /*client.scp(LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf", {
                            host: '172.16.204.71',
                            username: 'vomwega',
                            password: 'Stkenya.123',
                            path: '/tmp/demandletters/'
                        }, function(err) {
                            if (err) {
                                console.log(err);
                                res.json({
                                    result: 'error',
                                    message:  '/tmp/demandletters/' + accnumber_masked + DATE + "demand1.pdf",
                                    filename: accnumber_masked + DATE + "demand1.pdf",
                                    piped: false
                                })
                            } else {
                                console.log('file moved!');
                                res.json({
                                    result: 'success',
                                    message:  '/tmp/demandletters/' + accnumber_masked + DATE + "demand1.pdf",
                                    filename: accnumber_masked + DATE + "demand1.pdf",
                                    piped: true
                                })
                            }
                        })

                        res.json({
                            result: 'success',
                            message:  LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf",
                            filename: accnumber_masked + DATE + "demand1.pdf",
                            piped: true
                        })

                    }, error => {
                        console.log('error ...', error)
                        res.json({
                            result: 'error',
                            message: 'Exception occured'
                        });
                    })
            }*/
            // convert();
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
                    {
                        text: '\n\n\'\'Without Prejudice\'\'\n\n',
                        alignment: 'center',
                        bold: true
                    },
                    'Our Ref: DEMAND1/' + letter_data.branchcode + '/' + letter_data.arocode + '/' + DATE,
                    '\n' + DATE,
                    '\n' + letter_data.custname,
                    letter_data.address + '-' + letter_data.postcode,

                    '\nDear Sir/Madam',



                    {
                        text: 'RE: OUTSTANDING LIABILITIES A/C NO. ' + accnumber_masked + ' - ' + letter_data.custname + '',
                        style: 'subheader'
                    },

                    { text: '\nWe write to notify you that your account/s is currently in arrears/overdrawn. Kindly note that your current balance is as indicated below and it continues to accrue interest until payment is made in full.', alignment: 'justify' },
                    '\n',
                    {
                        alignment: 'justify',
                        fontSize: 9.5,
                        table: {
                            body: body,
                        },
                        layout: {
                            fillColor: function (rowIndex, node, columnIndex) {
                                return (rowIndex % 2 === 0) ? '#CCCCCC' : null;
                            }
                        }
                    },

                    {
                        text: '\nThe purpose of this letter therefore is to DEMAND immediate payment for the amount in arrears. Kindly ensure that the same is paid within ' + NOTICE + ' from the date hereof. ',
                        fontSize: 10, alignment: 'justify'
                    },

                    { text: '\nIn the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000. ', fontSize: 10, alignment: 'justify' },

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
            //pdfDoc.pipe(fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf"));
            //pdfDoc.end();

            // send response
            //res.json({
            //  result: 'success',
            // message: LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf",
            //filename: accnumber_masked + DATE + "demand1.pdf",
            //piped: true
            //})
            // ensures response is sent only after pdf is created
            writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf");
            pdfDoc.pipe(writeStream);
            pdfDoc.end();
            writeStream.on('finish', function () {
                // do stuff with the PDF file
                // send response
                res.json({
                    result: 'success',
                    message: LETTERS_DIR + accnumber_masked + DATE + "demand1.pdf",
                    filename: accnumber_masked + DATE + "demand1.pdf"
                })
            });
        } else {

            res.json({
                result: 'success',
                message: LETTERS_DIR + accnumber_masked + DATE + "demand1.docx",
                filename: accnumber_masked + DATE + "demand1.docx",
                piped: true
            });
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
