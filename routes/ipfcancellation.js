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
const EMAIL = data.SENDEMAILURL; // 'http://172.16.204.71:8005/demandemail/email';
const emaildata = {};
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
            '\n' + letter_data.insurancecompany,
            '' + letter_data.insuranceaddress,
            '' + letter_data.postcode,
            ' Head Office',
            '\nNAIROBI',

            '\nAttn: Underwriting Manager',

            '\nDear Sir/Madam',
            {
                text: '\nRE: CANCELLATION \n' + letter_data.custname + '\nPOLICY NO: ' + letter_data.policynumber,
                style: 'subheader'
            },
            {
                text: [
                    '\nWe herein notify you to cancel the insurance policy for the above named client due to unpaid instalments. Give us a refund cheque for the unutilized premiums of ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.refundamount)).format('0,0.00'), fontSize: 10, bold: true }
                ]
            },
            '\n',
            {
                alignment: 'justify',
                fontSize: 10,
                table: {
                    body: [
                        [{ text: 'Premium Amount', style: 'tableHeader', alignment: 'left' }, { text: 'Broker', style: 'tableHeader', alignment: 'left' }, { text: 'Date paid', style: 'tableHeader', alignment: 'left' }, { text: 'Days utilized', style: 'tableHeader', alignment: 'left' }, { text: 'Days Unutilized', style: 'tableHeader', alignment: 'left' }, { text: 'Refund amount', style: 'tableHeader', alignment: 'left' }],
                        [numeral(Math.abs(letter_data.oustbalance)).format('0,0.00'), letter_data.broker, letter_data.origdate, letter_data.daysutilized, letter_data.daysunutilized, numeral(Math.abs(letter_data.refundamount)).format('0,0.00')]
                    ]
                },
                layout: {
                    fillColor: function (rowIndex, node, columnIndex) {
                        return (rowIndex % 2 === 0) ? '#CCCCCC' : null;
                    }
                }
            },

            {
                text: '\nThis notice expires on the ' + DATEEXPIRY.toUpperCase() + ' which will be the rate review date for any demand arrears that will be owing to the bank.',
                fontSize: 10, alignment: 'justify'
            },

            { text: '\nRefund payment should be made out to the below Co-operative Bank Account. ', fontSize: 10, alignment: 'justify' },
            { text: '\nAccount Name: ' + letter_data.custname, fontSize: 10, alignment: 'left' },
            { text: 'Branch: ' + letter_data.branchname, fontSize: 10, alignment: 'left' },
            { text: 'Account No: 01694' + letter_data.custnumber + '00 ', fontSize: 10, alignment: 'left' },

            { text: '\nYour co-operation is highly appreciated. ', fontSize: 11, alignment: 'left' },

            { text: '\nYours Faithfully,' },
            { text: '\n\nDAVID MITHIA,                                                                           JAMES KARANJA', style: 'tableHeader' },
            { text: 'REMEDIAL CREDIT DEPARTMENT                                          FOR HEAD – MSME REMEDIAL CREDIT DEPARTMENT ', style: 'tableHeader' },
            { text: '\n\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },


            { text: '\nCc ' },
            { text: '' + letter_data.custname },
            { text: '' + letter_data.branchname }
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
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "ipfcancellation.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "ipfcancellation.pdf",
            filename: letter_data.accnumber + DATE + "ipfcancellation.pdf"
        })
    });
});

module.exports = router;