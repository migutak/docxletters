var express = require('express');
var router = express.Router();
const nodemailer = require("nodemailer");
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const cors = require('cors');

var data = require('./data.js');

const LETTERS_DIR = data.filePath;

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

router.use(bodyParser.urlencoded({
    extended: true
}));

router.use(bodyParser.json());
router.use(cors());


router.get('/', function (req, res) {
    res.json({ message: 'Ipf cancellation letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    console.log(letter_data)
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");

    var DATEEXPIRY = dateFormat(date1.setDate(date1.getDate() + 30), "dd-mmm-yyyy");
    var docDefinition = {
        pageSize: 'A4',
        pageOrientation: 'portrait',
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
            '',
            '\n' + DATE,
            '\n' + letter_data.storageyard,
            ' Head Office',

            '\nDear Sir/Madam',
            {
                text: '\nRe: RELEASE OF VEHICLE REG NO. ' + letter_data.regNumber ,
                style: 'subheader'
            },
            {
                text: [
                    '\nWe refer to the above matter. '
                ]
            },
            {
                text: [
                    '\nKindly release the Motor Vehicle Reg. No.  ',
                    { text: letter_data.regNumber, fontSize: 10, bold: true },
                    ' to the purchaser ',
                    { text: letter_data.customerName + ' ' +letter_data.nationID , fontSize: 10, bold: true },
                    ' upon proper identification.'
                ]
            },
            {
                text: [
                    '\nPlease note that we undertake to settle the storage charges up to release date  ',
                    { text: letter_data.releaseDate, fontSize: 10, bold: true },
                    ' i.e. any other charges there from should be borne by the purchaser.'
                ]
            },
            {
                text: [
                    '\nNote: Advise Re-possessor to raise their fee note to the department/Branch that issued repossession instructions, and copy remedial department for follow up. Storage fee to be send to Remedial department.',
                    
                ]
            },
            '\n',
            {
                table: {
                    headerRows: 1,
                    widths: [200, '*'],
                    body: [
                      [{text: '', style: 'tableHeader'}, {text: '', style: 'tableHeader'}],
                      ['\n\n---------------------------, ', ' \n\n---------------------------------------'],
                      ['REMEDIAL OFFICER', 'MANAGER, ADMIN & SUPPORT REMEDIAL MGT.']
                    ]
                },
                layout: 'noBorders',
                
            },


            { text: '\nCc ' },
            { text: '' + letter_data.customerName }
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
            },
            tableHeader: {
                bold: true,
                fontSize: 10,
                color: 'black'
            }
        },
        defaultStyle: {
            fontSize: 10
        }
    };
    var options = {
        // ..
    }

    var pdfDoc = printer.createPdfKitDocument(docDefinition, options);
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "release.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "release.pdf",
            filename: letter_data.accnumber + DATE + "release.pdf"
        })
    });
});

module.exports = router;
